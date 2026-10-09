import os
import sys
import json
import re

# Ajusta sys.path para garantir que o projeto reconheça 'src'
diretorio_raiz = os.path.dirname(os.path.abspath(__file__))
if diretorio_raiz not in sys.path:
    sys.path.insert(0, diretorio_raiz)

from src.modules.batch_processor import processar_lote_faturas
from src.modules.sheet_manager import salvar_na_planilha
from src.modules.pdf_exporter import converter_excel_para_pdf

def sanitizar_nome(texto):
    """Remove caracteres inválidos para nomes de pastas e arquivos no Windows (\ / : * ? " < > |)"""
    if not texto:
        return "OUTROS"
    texto_limpo = re.sub(r'[\\/*?:"<>|]', "", str(texto)).strip()
    return texto_limpo if texto_limpo else "OUTROS"

def identificar_pasta_fornecedor(fornecedor_raw):
    """Mapeia o nome do fornecedor retornado pela IA para a pasta correta entre os 6 homologados"""
    fornecedor_upper = str(fornecedor_raw).upper()
    
    if any(k in fornecedor_upper for k in ["INFOCOPPY", "SOUZA", "SANTANA"]):
        return "INFOCOPPY"
    elif "CLARO" in fornecedor_upper:
        return "CLARO"
    elif any(k in fornecedor_upper for k in ["VIVO", "TELEFONICA", "TELEFÔNICA"]):
        return "VIVO"
    elif "PCTEC" in fornecedor_upper:
        return "PCTEC"
    elif "WTT" in fornecedor_upper:
        return "WTT"
    elif any(k in fornecedor_upper for k in ["TASY", "TUCANO"]):
        return "TASY"
    else:
        return sanitizar_nome(fornecedor_raw)

def main():
    print("=== EXECUTANDO AUTOMAÇÃO DE PAGAMENTOS EM LOTE ===")
    
    # 1. Executa o processamento em lote para extrair os dados de todos os PDFs
    resultados = processar_lote_faturas()
    
    pasta_templates = os.path.join(diretorio_raiz, "data", "templates")
    pasta_output_base = os.path.join(diretorio_raiz, "data", "output")

    # 2. Para cada fatura extraída com sucesso, injeta no Excel e converte para PDF
    for item in resultados:
        if item["status"] != "SUCESSO" or not item["dados"]:
            continue

        dados = item["dados"]
        fornecedor_pasta = identificar_pasta_fornecedor(dados.get("fornecedor_nome"))
        numero_nf = sanitizar_nome(dados.get("numero_nf") or "SEM_NUMERO")

        pasta_destino = os.path.join(pasta_output_base, fornecedor_pasta)
        os.makedirs(pasta_destino, exist_ok=True)

        caminho_excel = os.path.join(pasta_destino, f"fatura_NF_{numero_nf}.xlsx")
        caminho_pdf = os.path.join(pasta_destino, f"fatura_NF_{numero_nf}.pdf")

        try:
            salvar_na_planilha(json.dumps(dados), pasta_templates, caminho_excel)
            converter_excel_para_pdf(caminho_excel, caminho_pdf)
            print(f"-> Sucesso ao gerar arquivos para NF {numero_nf} ({fornecedor_pasta})")
        except Exception as e:
            print(f"-> Erro ao gerar documentos da NF {numero_nf}: {e}")

    print("\n=== PROCESSO FINALIZADO ===")

if __name__ == "__main__":
    main()