import os
import sys
import json
import time
import warnings
from google import genai
from google.genai import types
from dotenv import load_dotenv

# Garante a resolução dos caminhos do projeto
diretorio_atual = os.path.dirname(os.path.abspath(__file__))
raiz_projeto = os.path.abspath(os.path.join(diretorio_atual, "..", ".."))
if raiz_projeto not in sys.path:
    sys.path.insert(0, raiz_projeto)

warnings.filterwarnings("ignore")

load_dotenv()
client = genai.Client()

def extrair_texto_pdf_local(caminho_pdf):
    """Extrai o texto do PDF em milissegundos localmente para evitar a fila do servidor."""
    try:
        import pypdf
        reader = pypdf.PdfReader(caminho_pdf)
        texto = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                texto += t + "\n"
        return texto.strip() if texto.strip() else None
    except Exception:
        return None

def extrair_dados_pdf(caminho_pdf):
    if not os.path.exists(caminho_pdf):
        raise FileNotFoundError(f"Ficheiro PDF não encontrado em: {caminho_pdf}")

    prompt_instrucao = """
    Você é um analista financeiro. Leia o documento a seguir.
    Extraia os seguintes dados e retorne ESTRITAMENTE em formato JSON com a seguinte estrutura:
    {
      "cnpj_emissor": "",
      "valor_total": "",
      "data_vencimento": "",
      "data_emissao": "",
      "fornecedor_nome": "",
      "numero_nf": "",
      "linha_digitavel": ""
    }
    Se o dado não existir no documento, defina o valor como null.
    """

    # Modelos estáveis com redundância
    modelos = [
        'gemini-2.0-flash',
        'gemini-1.5-flash',
        'gemini-2.0-flash-lite',
        'gemini-3.6-flash'
    ]

    # 1. Estratégia Principal: Envio de Texto Puro (Ultra-rápido, livre de 503)
    texto_local = extrair_texto_pdf_local(caminho_pdf)
    
    if texto_local:
        print("-> Texto do PDF lido localmente! Enviando requisição leve...")
        conteudo = f"{prompt_instrucao}\n\nTEXTO DO DOCUMENTO:\n{texto_local}"
        
        for modelo in modelos:
            for tentativa in range(1, 4):
                try:
                    res = client.models.generate_content(
                        model=modelo,
                        contents=conteudo,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    if res and res.text:
                        return res.text
                except Exception as e:
                    erro = str(e)
                    if "503" in erro or "429" in erro:
                        print(f"   [Aviso {modelo}] Servidor ocupado. Aguardando {tentativa * 2}s...")
                        time.sleep(tentativa * 2)
                    else:
                        break

    # 2. Estratégia Alternativa: Upload do PDF (Para faturas escaneadas/imagem)
    print("-> PDF em formato imagem/escaneado. Realizando upload do arquivo...")
    documento = client.files.upload(file=caminho_pdf)
    
    for modelo in modelos:
        for tentativa in range(1, 4):
            try:
                print(f"Analisando via {modelo} (Tentativa {tentativa}/3)...")
                res = client.models.generate_content(
                    model=modelo,
                    contents=[documento, prompt_instrucao],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                if res and res.text:
                    return res.text
            except Exception as e:
                erro = str(e)
                if "503" in erro or "429" in erro:
                    print(f"   [Aviso {modelo}] Ocupado. Aguardando {tentativa * 3}s...")
                    time.sleep(tentativa * 3)
                else:
                    break

    raise Exception("Falha na extração: Todos os modelos falharam após várias tentativas.")

if __name__ == "__main__":
    caminho_teste = os.path.join(raiz_projeto, "data", "input", "fatura_real.pdf")
    try:
        resultado = extrair_dados_pdf(caminho_teste)
        print("\n=== Resultado JSON ===")
        print(resultado)
    except Exception as e:
        print(f"\nErro durante a execução: {e}")