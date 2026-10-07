import os
from src.modules.batch_processor import processar_fatura


def main():
    print("=== INICIANDO A AUTOMAÇÃO INTEGRADA COM MULTI-TEMPLATES ===")

    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_pdf_entrada = os.path.join(diretorio_atual, "data", "input", "fatura_real.pdf")
    pasta_templates = os.path.join(diretorio_atual, "data", "templates")
    pasta_output_base = os.path.join(diretorio_atual, "data", "output")

    if not os.path.exists(caminho_pdf_entrada):
        print(f"-> PDF de entrada não encontrado: {caminho_pdf_entrada}")
        return

    print("\n1. A processar a fatura (extração, Excel e PDF)...")
    try:
        resultado = processar_fatura(caminho_pdf_entrada, pasta_templates, pasta_output_base)
    except Exception as e:
        print(f"-> Erro ao processar a fatura: {e}")
        return

    print("-> Dados extraídos:")
    print(resultado.get("dados"))
    print(f"-> Excel gerado em: {resultado.get('excel')}")
    if resultado.get("pdf"):
        print(f"-> PDF organizado em:\n{resultado.get('pdf')}")
    if resultado.get("erro"):
        print(f"-> Aviso: {resultado.get('erro')}")

    print(f"\n=== AUTOMAÇÃO CONCLUÍDA ({resultado.get('status')}) ===")

if __name__ == "__main__":
    main()
