import os
import fitz  # Import da biblioteca PyMuPDF

def preparar_pdf_docusign(lista_caminhos_pdf, caminho_saida, texto_ancora=r"\assinatura_dir_fin\\"):
    # Cria um novo PDF vazio que receberá as páginas
    pdf_final = fitz.open()

    print("Mesclando documentos...")
    for caminho in lista_caminhos_pdf:
        if not os.path.exists(caminho):
            print(f"Aviso: Arquivo não encontrado e será ignorado: {caminho}")
            continue
        
        # Abre o PDF atual e adiciona ao final do nosso PDF principal
        pdf_atual = fitz.open(caminho)
        pdf_final.insert_pdf(pdf_atual)
        pdf_atual.close()

    if pdf_final.page_count == 0:
        raise Exception("Nenhum PDF válido foi encontrado para mesclar.")

    # O DocuSign geralmente coloca as assinaturas na última página do pacote
    ultima_pagina = pdf_final[-1]
    
    # Calcula as coordenadas para colocar a âncora no canto inferior direito
    retangulo_pagina = ultima_pagina.rect
    x = retangulo_pagina.width - 180
    y = retangulo_pagina.height - 50
    
    # Carimba a âncora. O segredo do AutoPlace é a cor (1, 1, 1) = Branco RGB.
    # O robô do DocuSign lê, mas fica invisível na folha impressa/tela.
    print(f"Carimbando âncora do DocuSign ({texto_ancora}) na última página...")
    ultima_pagina.insert_text(
        fitz.Point(x, y),
        texto_ancora,
        fontsize=8,
        color=(1, 1, 1) 
    )

    # Salva o arquivo final mesclado e carimbado na pasta de output
    pdf_final.save(caminho_saida)
    pdf_final.close()
    print(f"PDF pronto para o DocuSign gerado em: {caminho_saida}")

if __name__ == "__main__":
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    
    # Vamos usar o mesmo PDF de teste para simular a mesclagem. 
    # Colocamos ele duas vezes na lista para simular "NF" + "Boleto".
    caminho_teste = os.path.abspath(os.path.join(diretorio_atual, "..", "..", "data", "input", "teste_pagamento.pdf"))
    caminho_saida = os.path.abspath(os.path.join(diretorio_atual, "..", "..", "data", "output", "documento_docusign_pronto.pdf"))
    
    try:
        # Passa a lista de arquivos de entrada e onde deve salvar a saída
        preparar_pdf_docusign([caminho_teste, caminho_teste], caminho_saida)
    except Exception as e:
        print(f"Erro: {e}")