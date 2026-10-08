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


def _gravar_pdf_imagem(caminho, texto):
    origem = pymupdf.open()
    pagina = origem.new_page(width=595, height=842)
    pagina.insert_textbox(pymupdf.Rect(36, 36, 560, 800), texto, fontsize=14)
    pix = pagina.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
    origem.close()
    saida = pymupdf.open()
    pagina_img = saida.new_page(width=595, height=842)
    pagina_img.insert_image(pagina_img.rect, pixmap=pix)
    saida.save(caminho)
    saida.close()

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

    def test_nfse_aceita_rotulo_com_espaco_e_rs_sem_cifrao(self):
        texto = """
        TASY
        Numero da NFS e 2286
        VALOR TOTAL DA NFS E RS 2.186,34
        Vencimento 10/09/2026
        """
        dados = analisar_texto_fatura(texto)
        self.assertEqual(dados["fornecedor_nome"], "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA")
        self.assertEqual(dados["valor_total"], "2.186,34")
        self.assertEqual(dados["numero_nf"], "2286")

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

    def test_pdf_sem_texto_nao_chama_servico_externo(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "scan.pdf")
            documento = pymupdf.open()
            documento.new_page()
            documento.save(caminho)
            documento.close()
            with patch("src.modules.extrator_pdf._texto_ocr", return_value=None) as ocr:
                with self.assertRaises(ValueError) as ctx:
                    extrair_dados_pdf(caminho)
            ocr.assert_called_once_with(caminho)
            mensagem = str(ctx.exception).lower()
            self.assertIn("não tem texto", mensagem)
            self.assertIn("imagem", mensagem)

    def test_pdf_imagem_e_lido_por_ocr(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "NF tasy.pdf")
            documento = pymupdf.open()
            documento.new_page()
            documento.save(caminho)
            documento.close()
            with patch("src.modules.extrator_pdf._texto_ocr", return_value=TASY) as ocr:
                dados = json.loads(extrair_dados_pdf(caminho))
            ocr.assert_called_once_with(caminho)
            self.assertEqual(dados["fornecedor_nome"], "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA")
            self.assertEqual(dados["valor_total"], "2.186,34")
            self.assertEqual(dados["numero_nf"], "2286")

    def test_nome_do_arquivo_completa_fornecedor_tasy(self):
        texto = """
        Numero da NFS-e 2286
        Emissao 01/08/2026
        Vencimento 10/09/2026
        Valor dos serviços R$ 2.186,34
        """
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "NF tasy.pdf")
            documento = pymupdf.open()
            documento.new_page()
            documento.save(caminho)
            documento.close()
            with patch("src.modules.extrator_pdf._texto_ocr", return_value=texto):
                dados = json.loads(extrair_dados_pdf(caminho))
            self.assertEqual(dados["fornecedor_nome"], "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA")
            self.assertEqual(dados["valor_total"], "2.186,34")

    def test_pdf_digital_nao_dispara_ocr(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "claro.pdf")
            documento = pymupdf.open()
            pagina = documento.new_page()
            pagina.insert_textbox(pymupdf.Rect(36, 36, 560, 800), CLARO, fontsize=12)
            documento.save(caminho)
            documento.close()
            with patch("src.modules.extrator_pdf._texto_ocr") as ocr:
                dados = json.loads(extrair_dados_pdf(caminho))
            ocr.assert_not_called()
            self.assertEqual(dados["fornecedor_nome"], "Claro S/A")
            self.assertEqual(dados["valor_total"], "359,96")

    def test_texto_incompleto_ainda_tenta_ocr(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "incompleto.pdf")
            documento = pymupdf.open()
            pagina = documento.new_page()
            pagina.insert_text((72, 72), "Apenas um aviso interno sem valor.")
            documento.save(caminho)
            documento.close()
            with patch("src.modules.extrator_pdf._texto_ocr", return_value=TASY) as ocr:
                dados = json.loads(extrair_dados_pdf(caminho))
            ocr.assert_called_once_with(caminho)
            self.assertEqual(dados["valor_total"], "2.186,34")
            self.assertEqual(dados["fornecedor_nome"], "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA")

    def test_texto_incompleto_explica_o_que_faltou(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "incompleto.pdf")
            documento = pymupdf.open()
            pagina = documento.new_page()
            pagina.insert_text((72, 72), "Apenas um aviso interno sem valor.")
            documento.save(caminho)
            documento.close()
            with patch("src.modules.extrator_pdf._texto_ocr", return_value=None):
                with self.assertRaises(ValueError) as ctx:
                    extrair_dados_pdf(caminho)
            self.assertIn("fornecedor", str(ctx.exception))

    def test_ocr_real_le_nfse_tasy_em_imagem(self):
        try:
            from rapidocr import RapidOCR  # noqa: F401
        except Exception:
            self.skipTest("rapidocr não está instalado")
        with tempfile.TemporaryDirectory() as pasta:
            caminho = os.path.join(pasta, "NF tasy.pdf")
            _gravar_pdf_imagem(caminho, TASY)
            dados = json.loads(extrair_dados_pdf(caminho))
            self.assertEqual(dados["fornecedor_nome"], "TUCANO DO BRASIL SISTEMAS DE INFORMACAO LTDA")
            self.assertEqual(dados["valor_total"], "2.186,34")
            self.assertEqual(dados["numero_nf"], "2286")

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
