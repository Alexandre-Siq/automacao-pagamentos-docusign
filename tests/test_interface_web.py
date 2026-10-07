import json
import os
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.interface_web import InterfacePagamentos, criar_servidor, nome_pdf_seguro


def _multipart(arquivos):
    limite = "----LimiteTeste"
    partes = []
    for nome, conteudo in arquivos:
        partes.append(
            (
                f"--{limite}\r\n"
                f'Content-Disposition: form-data; name="faturas"; filename="{nome}"\r\n'
                f"Content-Type: application/pdf\r\n\r\n"
            ).encode("utf-8")
            + conteudo
            + b"\r\n"
        )
    corpo = b"".join(partes) + f"--{limite}--\r\n".encode("utf-8")
    return corpo, f"multipart/form-data; boundary={limite}"


class ServidorTemporario:
    def __init__(self, aplicacao):
        self.httpd = criar_servidor(aplicacao, porta=0)
        self.thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self.thread.start()
        porta = self.httpd.server_address[1]
        self.url = f"http://127.0.0.1:{porta}"

    def fechar(self):
        self.httpd.shutdown()
        self.httpd.server_close()
        self.thread.join(timeout=3)

    def enviar(self, caminho, dados=None, tipo=None):
        requisicao = urllib.request.Request(self.url + caminho, data=dados, method="POST" if dados is not None else "GET")
        if tipo:
            requisicao.add_header("Content-Type", tipo)
        try:
            with urllib.request.urlopen(requisicao) as resposta:
                return resposta.status, resposta.headers, resposta.read()
        except urllib.error.HTTPError as erro:
            return erro.code, erro.headers, erro.read()


class TestNome(unittest.TestCase):
    def test_aceita_pdf_e_recusa_outro_tipo(self):
        self.assertEqual(nome_pdf_seguro(r"..\nota.PDF"), "nota.PDF")
        self.assertIsNone(nome_pdf_seguro("planilha.xlsx"))
        self.assertIsNone(nome_pdf_seguro(".pdf"))


class TestTela(unittest.TestCase):
    def setUp(self):
        self.temporario = tempfile.TemporaryDirectory()
        self.raiz = self.temporario.name
        os.makedirs(os.path.join(self.raiz, "data", "templates"))
        self.vistas = []

        def processar(pasta_input, pasta_output, pasta_templates):
            self.vistas.append(os.listdir(pasta_input))
            nome = sorted(os.listdir(pasta_input))[0]
            with open(os.path.join(pasta_input, nome), "rb") as arquivo:
                self.conteudo = arquivo.read()
            pasta = os.path.join(pasta_output, "CLARO")
            os.makedirs(pasta, exist_ok=True)
            excel = os.path.join(pasta, "fatura_NF_10.xlsx")
            pdf = os.path.join(pasta, "fatura_NF_10.pdf")
            with open(excel, "wb") as arquivo:
                arquivo.write(b"excel")
            with open(pdf, "wb") as arquivo:
                arquivo.write(b"%PDF")
            return [{
                "arquivo_origem": nome,
                "status": "SUCESSO",
                "dados": {"fornecedor_nome": "Claro S/A", "valor_total": "359,96", "numero_nf": "10"},
                "excel": excel,
                "pdf": pdf,
                "docusign": None,
                "erro": None,
            }]

        self.aplicacao = InterfacePagamentos(self.raiz, processar=processar)
        self.servidor = ServidorTemporario(self.aplicacao)

    def tearDown(self):
        self.servidor.fechar()
        self.temporario.cleanup()

    def test_pagina_inicial_tem_a_area_de_soltar(self):
        codigo, cabecalhos, corpo = self.servidor.enviar("/")
        self.assertEqual(codigo, 200)
        self.assertIn("text/html", cabecalhos.get("Content-Type"))
        texto = corpo.decode("utf-8")
        self.assertIn("Solte os PDFs aqui", texto)
        self.assertIn("Processar", texto)

    def test_processa_o_pdf_e_entrega_os_arquivos(self):
        conteudo = b"%PDF-1.4\n%\x00\xff fatura"
        corpo, tipo = _multipart([("fatura claro.pdf", conteudo)])
        codigo, _, resposta = self.servidor.enviar("/processar", corpo, tipo)
        self.assertEqual(codigo, 200)
        dados = json.loads(resposta.decode("utf-8"))
        self.assertEqual(dados["resumo"]["sucessos"], 1)
        self.assertEqual(dados["itens"][0]["fornecedor"], "Claro S/A")
        self.assertEqual(dados["itens"][0]["valor"], "359,96")
        self.assertIsNone(dados["itens"][0]["docusign"])
        self.assertEqual(self.conteudo, conteudo)

        codigo, cabecalhos, baixado = self.servidor.enviar(dados["itens"][0]["pdf"])
        self.assertEqual(codigo, 200)
        self.assertEqual(baixado, b"%PDF")
        self.assertIn("application/pdf", cabecalhos.get("Content-Type"))

        codigo, cabecalhos, baixado = self.servidor.enviar(dados["itens"][0]["excel"])
        self.assertEqual(baixado, b"excel")
        self.assertIn("attachment", cabecalhos.get("Content-Disposition"))

    def test_recusa_arquivo_que_nao_e_pdf(self):
        corpo, tipo = _multipart([("notas.txt", b"ola")])
        codigo, _, resposta = self.servidor.enviar("/processar", corpo, tipo)
        self.assertEqual(codigo, 400)
        self.assertIn("Nenhuma fatura", json.loads(resposta.decode("utf-8"))["erro"])

    def test_download_sem_token_responde_404(self):
        codigo, _, _ = self.servidor.enviar("/baixar/nao-existe")
        self.assertEqual(codigo, 404)


if __name__ == "__main__":
    unittest.main()
