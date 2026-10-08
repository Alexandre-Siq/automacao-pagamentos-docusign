import json
import os
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

import pymupdf

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.modules.batch_processor import processar_fatura
from src.modules.extrator_pdf import analisar_texto_fatura, extrair_dados_pdf

PASTA_TEMPLATES = os.path.join(ROOT, "data", "templates")

NFSE_INFOCOPPY = """
NOTA FISCAL ELETRONICA DE SERVICOS - NFS-e
Numero da Nota: 189
Data e Hora de Emissao: 24/08/2026

PRESTADOR DE SERVICOS
Souza & Santana Suprimentos e Soluções Técnicas Ltda – ME
CNPJ: 12.345.678/0001-90

TOMADOR DE SERVICOS
Complexo Hospitalar dos Estivadores
CNPJ: 99.999.999/0001-99

VALOR TOTAL DO SERVICO = R$ 24.654,98
Vencimento 10/09/2026
"""

CLARO = """
CLARO S/A
CNPJ 40.432.544/0001-47
Número da Fatura: 160995206
Data de Emissão: 11/08/2026
Vencimento: 24/08/2026
Total a pagar R$ 359,96
"""

VIVO = """
TELEFONICA BRASIL S.A.
CNPJ: 02.558.157/0001-62
Vivo Empresas
Nº Fatura: 686186
Emissão 29/07/2026
Data de Vencimento 01/09/2026
Valor Total da Fatura R$ 263,17
"""

PCTEC = """
PCTEC OUTSOURCING LTDA
PRESTADOR
CNPJ 11.222.333/0001-44
Nota Fiscal Nº 341
Emissão: 03/08/2026
Vencimento: 30/08/2026
Valor Total da Nota: R$ 5.171,89
"""

WTT = """
WTT TECNOLOGIA E CONSULTORIA
NFS-e Nº 55
Data de Emissão 01/08/2026
Vencimento 15/08/2026
Valor líquido da nota R$ 1.000,00
"""

TASY = """
TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA
Numero da NFS-e 2286
Emissao 01/08/2026
Vencimento 10/09/2026
Valor dos serviços R$ 2.186,34
"""


class TestAnaliseLocal(unittest.TestCase):
    def test_nfse_ignora_cnpj_do_tomador(self):
        dados = analisar_texto_fatura(NFSE_INFOCOPPY)
        self.assertEqual(dados["fornecedor_nome"], "Souza & Santana Suprimentos e Soluções Técnicas Ltda – ME")
        self.assertEqual(dados["cnpj_emissor"], "12.345.678/0001-90")
        self.assertEqual(dados["numero_nf"], "189")
        self.assertEqual(dados["valor_total"], "24.654,98")
        self.assertEqual(dados["data_emissao"], "24/08/2026")
        self.assertEqual(dados["data_vencimento"], "10/09/2026")

    def test_faturas_dos_fornecedores_mapeados(self):
        casos = [
            (CLARO, "Claro S/A", "160995206", "359,96", "40.432.544/0001-47"),
            (VIVO, "Telefônica Brasil S.A.", "686186", "263,17", "02.558.157/0001-62"),
            (PCTEC, "PCTEC OUTSOURCING LTDA", "341", "5.171,89", "11.222.333/0001-44"),
            (WTT, "WTT TECNOLOGIA E CONSULTORIA", "55", "1.000,00", None),
            (TASY, "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA", "2286", "2.186,34", None),
        ]
        for texto, fornecedor, numero, valor, cnpj in casos:
            with self.subTest(fornecedor=fornecedor):
                dados = analisar_texto_fatura(texto)
                self.assertEqual(dados["fornecedor_nome"], fornecedor)
                self.assertEqual(dados["numero_nf"], numero)
                self.assertEqual(dados["valor_total"], valor)
                self.assertEqual(dados["cnpj_emissor"], cnpj)

    def test_valor_sai_da_linha_digitavel_quando_nao_ha_rotulo(self):
        texto = """
        CLARO S/A
        Numero da Fatura 160995206
        23792.37001 60000.123456 00000.123456 1 12340002465498
        """
        dados = analisar_texto_fatura(texto)
        self.assertEqual(dados["valor_total"], "24.654,98")
        self.assertIn("12340002465498", dados["linha_digitavel"])

    def test_data_pode_vir_antes_do_rotulo(self):
        texto = """
        VIVO
        01/09/2026 Vencimento
        Total a pagar 80,00
        """
        dados = analisar_texto_fatura(texto)
        self.assertEqual(dados["data_vencimento"], "01/09/2026")
        self.assertEqual(dados["valor_total"], "80,00")

    def test_pdf_digital_preenche_a_planilha_sem_api(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "claro.pdf")
            documento = pymupdf.open()
            pagina = documento.new_page()
            pagina.insert_textbox(pymupdf.Rect(36, 36, 560, 800), CLARO, fontsize=12)
            documento.save(caminho)
            documento.close()

            def exportar(_origem, destino):
                documento_saida = pymupdf.open()
                pagina_saida = documento_saida.new_page()
                pagina_saida.insert_text((72, 72), "Solicitacao de pagamento")
                documento_saida.save(destino)
                documento_saida.close()
                return destino

            with patch("src.modules.batch_processor.converter_excel_para_pdf", side_effect=exportar):
                resultado = processar_fatura(caminho, PASTA_TEMPLATES, pasta)

            self.assertEqual(resultado["status"], "SUCESSO")
            self.assertEqual(resultado["dados"]["numero_nf"], "160995206")
            self.assertEqual(resultado["dados"]["valor_total"], "359,96")
            self.assertIn(os.path.join("CLARO", "fatura_NF_160995206.xlsx"), resultado["excel"])
            self.assertTrue(os.path.exists(resultado["docusign"]))

            original = pymupdf.open(caminho)
            pacote = pymupdf.open(resultado["docusign"])
            try:
                self.assertEqual(pacote.page_count, original.page_count + 1)
                self.assertIn("\\assinatura_dir_fin\\", pacote[-1].get_text())
            finally:
                original.close()
                pacote.close()

    def test_pdf_em_branco_nao_chama_servico_externo(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "scan.pdf")
            documento = pymupdf.open()
            documento.new_page()
            documento.save(caminho)
            documento.close()
            with self.assertRaises(ValueError) as ctx:
                extrair_dados_pdf(caminho)
            self.assertIn("não tem texto", str(ctx.exception))

    def test_pdf_so_com_imagem_le_a_nota_da_tasy(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "NF tasy.pdf")
            desenhada = pymupdf.open()
            pagina = desenhada.new_page()
            y = 80
            for linha in (
                "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA",
                "Numero da NFS-e 2286",
                "Emissao 01/08/2026",
                "Vencimento 10/09/2026",
                "Valor dos servicos R$ 2.186,34",
            ):
                pagina.insert_text((72, y), linha, fontsize=14)
                y += 32
            imagem = pagina.get_pixmap(matrix=pymupdf.Matrix(3, 3), alpha=False)
            desenhada.close()

            documento = pymupdf.open()
            pagina_imagem = documento.new_page()
            pagina_imagem.insert_image(pagina_imagem.rect, pixmap=imagem)
            documento.save(caminho)
            documento.close()

            conferido = pymupdf.open(caminho)
            try:
                self.assertEqual(conferido[0].get_text().strip(), "")
            finally:
                conferido.close()

            dados = json.loads(extrair_dados_pdf(caminho))
            self.assertEqual(dados["fornecedor_nome"], "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA")
            self.assertEqual(dados["numero_nf"], "2286")
            self.assertEqual(dados["valor_total"], "2.186,34")
            self.assertEqual(dados["data_emissao"], "01/08/2026")
            self.assertEqual(dados["data_vencimento"], "10/09/2026")

    def test_texto_incompleto_explica_o_que_faltou(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "incompleto.pdf")
            documento = pymupdf.open()
            pagina = documento.new_page()
            pagina.insert_text((72, 72), "Apenas um aviso interno sem valor.")
            documento.save(caminho)
            documento.close()
            with self.assertRaises(ValueError) as ctx:
                extrair_dados_pdf(caminho)
            self.assertIn("fornecedor", str(ctx.exception))

    def test_arquivo_ausente(self):
        with self.assertRaises(FileNotFoundError):
            extrair_dados_pdf(os.path.join(tempfile.gettempdir(), "fatura-inexistente.pdf"))

    def test_retorno_e_json(self):
        dados = json.loads(json.dumps(analisar_texto_fatura(CLARO), ensure_ascii=False))
        self.assertEqual(dados["fornecedor_nome"], "Claro S/A")

    def test_importar_o_lote_nao_carrega_o_gemini(self):
        codigo = (
            "import src.modules.extrator_pdf, src.modules.batch_processor, src.modules.gemini_extractor; "
            "import sys; assert 'google.genai' not in sys.modules and 'google' not in sys.modules"
        )
        subprocess.run([sys.executable, "-c", codigo], cwd=ROOT, check=True)


if __name__ == "__main__":
    unittest.main()
