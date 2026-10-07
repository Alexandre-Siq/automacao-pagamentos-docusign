"""A leitura da fatura é local. Este módulo só preserva o nome antigo da função."""

import os
import sys

diretorio_atual = os.path.dirname(os.path.abspath(__file__))
raiz_projeto = os.path.abspath(os.path.join(diretorio_atual, "..", ".."))
if raiz_projeto not in sys.path:
    sys.path.insert(0, raiz_projeto)

from src.modules.extrator_pdf import extrair_dados_pdf, extrair_texto_pdf_local

__all__ = ["extrair_dados_pdf", "extrair_texto_pdf_local"]


if __name__ == "__main__":
    caminho_teste = os.path.join(raiz_projeto, "data", "input", "fatura_real.pdf")
    try:
        print(extrair_dados_pdf(caminho_teste))
    except Exception as e:
        print(f"\nErro durante a execução: {e}")
