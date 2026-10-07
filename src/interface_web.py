import json
import os
import re
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlparse

from src.modules.batch_processor import processar_lote_faturas

CAMINHO_HTML = os.path.join(os.path.dirname(__file__), "interface_web.html")
LIMITE_ARQUIVO = 25 * 1024 * 1024
LIMITE_ENVIO = 80 * 1024 * 1024


def diretorio_projeto():
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def nome_pdf_seguro(nome):
    bruto = (nome or "").replace("\x00", "").replace("\\", "/")
    base = os.path.basename(bruto)
    base = re.sub(r'[\\/:*?"<>|]', "", base).strip()
    if not base.lower().endswith(".pdf") or base.lower() == ".pdf":
        return None
    return base


def extrair_pdfs(corpo, content_type):
    """Separa os PDFs de um envio multipart sem alterar os bytes do arquivo."""
    encontrado = re.search(r'boundary="?([^";]+)"?', content_type or "", re.I)
    if not encontrado:
        raise ValueError("O envio não trouxe os arquivos.")
    limite = b"--" + encontrado.group(1).encode("utf-8", "replace")
    arquivos = []
    for pedaco in corpo.split(limite):
        if not pedaco or pedaco.startswith(b"--"):
            continue
        if pedaco.startswith(b"\r\n"):
            pedaco = pedaco[2:]
        if b"\r\n\r\n" not in pedaco:
            continue
        cabecalho, dados = pedaco.split(b"\r\n\r\n", 1)
        if dados.endswith(b"\r\n"):
            dados = dados[:-2]
        texto = cabecalho.decode("utf-8", "replace")
        if "filename=" not in texto.lower():
            continue
        nome = _nome_do_cabecalho(texto)
        seguro = nome_pdf_seguro(nome)
        if seguro:
            arquivos.append((seguro, dados))
    return arquivos


def _nome_do_cabecalho(cabecalho):
    codificado = re.search(r"filename\*=UTF-8''([^;\r\n]+)", cabecalho, re.I)
    if codificado:
        return unquote(codificado.group(1).strip().strip('"'))
    aspas = re.search(r'filename="([^"]*)"', cabecalho, re.I)
    if aspas:
        return aspas.group(1)
    simples = re.search(r"filename=([^;\r\n]+)", cabecalho, re.I)
    if simples:
        return simples.group(1).strip().strip('"')
    return ""


class InterfacePagamentos:
    def __init__(self, diretorio_raiz, processar=None):
        self.raiz = os.path.abspath(diretorio_raiz)
        self.processar = processar or processar_lote_faturas
        self.arquivos = {}
        self.trava = threading.Lock()
        self._sequencia = 0

    @property
    def pasta_saida(self):
        return os.path.join(self.raiz, "data", "output")

    @property
    def pasta_templates(self):
        return os.path.join(self.raiz, "data", "templates")

    def processar_envio(self, arquivos):
        if not arquivos:
            raise ValueError("Nenhuma fatura PDF foi enviada.")
        if sum(len(conteudo) for _, conteudo in arquivos) > LIMITE_ENVIO:
            raise ValueError("O envio passou de 80 MB.")
        for nome, conteudo in arquivos:
            if len(conteudo) > LIMITE_ARQUIVO:
                raise ValueError(f"{nome} passou de 25 MB.")

        import tempfile

        with tempfile.TemporaryDirectory(prefix="faturas_") as pasta_entrada:
            usados = set()
            for nome, conteudo in arquivos:
                destino = os.path.join(pasta_entrada, self._nome_unico(nome, usados))
                with open(destino, "wb") as arquivo:
                    arquivo.write(conteudo)
            with self.trava:
                resultados = self.processar(
                    pasta_input=pasta_entrada,
                    pasta_output=self.pasta_saida,
                    pasta_templates=self.pasta_templates,
                )
        return self._montar_resposta(resultados or [])

    def abrir_pasta_saida(self):
        os.makedirs(self.pasta_saida, exist_ok=True)
        if os.name == "nt":
            os.startfile(self.pasta_saida)
            return
        import subprocess
        import sys

        comando = ["open", self.pasta_saida] if sys.platform == "darwin" else ["xdg-open", self.pasta_saida]
        subprocess.Popen(comando, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def caminho_do_token(self, token):
        return self.arquivos.get(token)

    def _nome_unico(self, nome, usados):
        candidato = nome
        indice = 2
        while candidato.lower() in usados:
            stem, extensao = os.path.splitext(nome)
            candidato = f"{stem}_{indice}{extensao}"
            indice += 1
        usados.add(candidato.lower())
        return candidato

    def _montar_resposta(self, resultados):
        itens = []
        for item in resultados:
            dados = item.get("dados") or {}
            itens.append({
                "arquivo": item.get("arquivo_origem") or "Fatura",
                "status": item.get("status") or "ERRO",
                "fornecedor": dados.get("fornecedor_nome") or "Não identificado",
                "valor": dados.get("valor_total") or "",
                "numero_nf": dados.get("numero_nf") or "",
                "erro": item.get("erro"),
                "excel": self._publicar(item.get("excel")),
                "pdf": self._publicar(item.get("pdf")),
                "docusign": self._publicar(item.get("docusign")),
            })
        return {
            "itens": itens,
            "resumo": {
                "total": len(itens),
                "sucessos": sum(1 for item in itens if item["status"] == "SUCESSO"),
                "parciais": sum(1 for item in itens if item["status"] == "PARCIAL"),
                "falhas": sum(1 for item in itens if item["status"] == "ERRO"),
            },
        }

    def _publicar(self, caminho):
        if not caminho or not os.path.isfile(caminho):
            return None
        absoluto = os.path.abspath(caminho)
        saida = os.path.abspath(self.pasta_saida)
        try:
            if os.path.commonpath([saida, absoluto]) != saida:
                return None
        except ValueError:
            return None
        self._sequencia += 1
        token = f"{self._sequencia}"
        self.arquivos[token] = absoluto
        return "/baixar/" + token


def fabricar_handler(aplicacao, ao_encerrar=None):
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            caminho = urlparse(self.path).path
            if caminho in ("/", "/index.html"):
                self._enviar_pagina()
                return
            if caminho.startswith("/baixar/"):
                self._enviar_arquivo(caminho[len("/baixar/"):])
                return
            self.send_error(404)

        def do_POST(self):
            caminho = urlparse(self.path).path
            if caminho == "/processar":
                self._processar()
                return
            if caminho == "/abrir-pasta":
                self._abrir_pasta()
                return
            if caminho == "/encerrar":
                self._encerrar()
                return
            self.send_error(404)

        def log_message(self, formato, *args):
            return

        def _processar(self):
            try:
                tamanho = int(self.headers.get("Content-Length", "0"))
            except ValueError:
                tamanho = 0
            if tamanho <= 0 or tamanho > LIMITE_ENVIO + 1024 * 1024:
                _enviar_json(self, 400, {"erro": "Envio vazio ou grande demais."})
                return
            corpo = self.rfile.read(tamanho)
            try:
                arquivos = extrair_pdfs(corpo, self.headers.get("Content-Type", ""))
                resposta = aplicacao.processar_envio(arquivos)
            except ValueError as erro:
                _enviar_json(self, 400, {"erro": str(erro)})
                return
            except Exception as erro:
                _enviar_json(self, 500, {"erro": str(erro)})
                return
            _enviar_json(self, 200, resposta)

        def _abrir_pasta(self):
            try:
                aplicacao.abrir_pasta_saida()
            except Exception as erro:
                _enviar_json(self, 500, {"erro": str(erro)})
                return
            _enviar_json(self, 200, {"ok": True})

        def _encerrar(self):
            _enviar_json(self, 200, {"ok": True})
            if ao_encerrar:
                threading.Thread(target=ao_encerrar, daemon=True).start()

        def _enviar_pagina(self):
            try:
                with open(CAMINHO_HTML, "rb") as arquivo:
                    corpo = arquivo.read()
            except OSError:
                self.send_error(500, "Página não encontrada")
                return
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(corpo)

        def _enviar_arquivo(self, token):
            caminho = aplicacao.caminho_do_token(token)
            if not caminho or not os.path.isfile(caminho):
                self.send_error(404)
                return
            with open(caminho, "rb") as arquivo:
                corpo = arquivo.read()
            nome = os.path.basename(caminho)
            if nome.lower().endswith(".pdf"):
                tipo = "application/pdf"
                disposicao = "inline"
            else:
                tipo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                disposicao = "attachment"
            self.send_response(200)
            self.send_header("Content-Type", tipo)
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Content-Disposition", f'{disposicao}; filename="{nome}"')
            self.end_headers()
            self.wfile.write(corpo)

    return Handler


def _enviar_json(handler, codigo, payload):
    corpo = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(codigo)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(corpo)))
    handler.end_headers()
    handler.wfile.write(corpo)


def criar_servidor(aplicacao, porta=8765, ao_encerrar=None):
    handler = fabricar_handler(aplicacao, ao_encerrar)
    if porta == 0:
        return ThreadingHTTPServer(("127.0.0.1", 0), handler)
    ultimo_erro = None
    for candidata in range(porta, porta + 20):
        try:
            return ThreadingHTTPServer(("127.0.0.1", candidata), handler)
        except OSError as erro:
            ultimo_erro = erro
    raise ultimo_erro


def iniciar(abrir_navegador=True, porta=8765):
    aplicacao = InterfacePagamentos(diretorio_projeto())
    servidor = None

    def encerrar():
        if servidor:
            servidor.shutdown()

    servidor = criar_servidor(aplicacao, porta, encerrar)
    url = f"http://127.0.0.1:{servidor.server_address[1]}/"
    if abrir_navegador:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    print(f"Interface aberta em {url}")
    print("Feche esta janela, ou use Encerrar na tela, para sair.")
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        servidor.server_close()


def avisar_erro(mensagem):
    if os.name == "nt":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, mensagem, "Pagamentos", 0x10)
        return
    print(mensagem)
