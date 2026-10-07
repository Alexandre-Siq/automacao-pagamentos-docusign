import os
import json
import re
from src.modules.gemini_extractor import extrair_dados_pdf
from src.modules.sheet_manager import salvar_na_planilha
from src.modules.pdf_exporter import converter_excel_para_pdf

def sanitizar_nome(texto):
    """Remove caracteres inválidos para nomes de pastas no Windows (\ / : * ? " < > |)"""
    if not texto:
        return "OUTROS"
    texto_limpo = re.sub(r'[\\/*?:"<>|]', "", str(texto)).strip()
    return texto_limpo if texto_limpo else "OUTROS"

def identificar_pasta_fornecedor(fornecedor_raw):
    """Mapeia o nome do fornecedor retornado pela IA para a pasta correta entre os 6 homologados"""
    fornecedor_upper = str(fornecedor_raw).upper()
    
    if "INFOCOPPY" in fornecedor_upper or "SOUZA" in fornecedor_upper or "SANTANA" in fornecedor_upper:
        return "INFOCOPPY"
    elif "CLARO" in fornecedor_upper:
        return "CLARO"
    elif "VIVO" in fornecedor_upper or "TELEFONICA" in fornecedor_upper or "TELEFÔNICA" in fornecedor_upper:
        return "VIVO"
    elif "PCTEC" in fornecedor_upper:
        return "PCTEC"
    elif "WTT" in fornecedor_upper:
        return "WTT"
    elif "TASY" in fornecedor_upper or "TUCANO" in fornecedor_upper:
        return "TASY"
    else:
        return sanitizar_nome(fornecedor_raw)

def main():
    print("=== INICIANDO A AUTOMAÇÃO INTEGRADA COM MULTI-TEMPLATES ===")
    
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_pdf_entrada = os.path.join(diretorio_atual, "data", "input", "fatura_real.pdf")
    pasta_templates = os.path.join(diretorio_atual, "data", "templates")
    pasta_output_base = os.path.join(diretorio_atual, "data", "output")

    print("\n1. A enviar fatura para o Gemini...")
    try:
        dados_json_str = extrair_dados_pdf(caminho_pdf_entrada)
        print("-> Dados extraídos com sucesso!")
        print(dados_json_str)
    except Exception as e:
        print(f"-> Erro na extração: {e}")
        return

    # Extrai o nome do fornecedor e o número da NF
    try:
        dados_dict = json.loads(dados_json_str)
        fornecedor_raw = dados_dict.get("fornecedor_nome", "DESCONHECIDO")
        numero_nf_raw = dados_dict.get("numero_nf", "SEM_NUMERO")
        
        fornecedor_pasta = identificar_pasta_fornecedor(fornecedor_raw)
        numero_nf = sanitizar_nome(numero_nf_raw)
    except Exception:
        fornecedor_pasta = "GERAL"
        numero_nf = "NF"

    # Define a pasta específica do fornecedor em data/output/<FORNECEDOR>/
    pasta_fornecedor = os.path.join(pasta_output_base, fornecedor_pasta)
    os.makedirs(pasta_fornecedor, exist_ok=True)

    # Define os caminhos de saída organizados por pasta e número de NF
    caminho_excel_saida = os.path.join(pasta_fornecedor, f"fatura_NF_{numero_nf}.xlsx")
    caminho_pdf_saida = os.path.join(pasta_fornecedor, f"fatura_NF_{numero_nf}.pdf")

    print(f"\n2. A injetar dados no Template Excel (Pasta: output/{fornecedor_pasta})...")
    try:
        salvar_na_planilha(dados_json_str, pasta_templates, caminho_excel_saida)
        print(f"-> Excel gerado em: {caminho_excel_saida}")
    except Exception as e:
        print(f"-> Erro ao gerar Excel: {e}")
        return

    print("\n3. A converter silenciosamente para PDF...")
    try:
        converter_excel_para_pdf(caminho_excel_saida, caminho_pdf_saida)
        print(f"-> Sucesso! PDF organizado em:\n{caminho_pdf_saida}")
    except Exception as e:
        print(f"-> Erro na conversão para PDF: {e}")
        return
        
    print("\n=== AUTOMAÇÃO CONCLUÍDA ===")

if __name__ == "__main__":
    main()