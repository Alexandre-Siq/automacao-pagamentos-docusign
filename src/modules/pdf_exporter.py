import os
import win32com.client

def converter_excel_para_pdf(caminho_excel, caminho_pdf):
    # Garante caminhos absolutos exigidos pelo Windows COM
    caminho_excel_abs = os.path.abspath(caminho_excel)
    caminho_pdf_abs = os.path.abspath(caminho_pdf)
    
    if not os.path.exists(caminho_excel_abs):
        raise FileNotFoundError(f"Ficheiro Excel não encontrado: {caminho_excel_abs}")

    # Remove o PDF antigo se já existir para evitar conflitos de escrita
    if os.path.exists(caminho_pdf_abs):
        try:
            os.remove(caminho_pdf_abs)
        except Exception:
            pass

    # Usamos DispatchEx para criar uma instância totalmente limpa e isolada do Excel
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False 
    
    wb = None
    try:
        # Abre o ficheiro em modo de leitura
        wb = excel.Workbooks.Open(caminho_excel_abs, ReadOnly=True)
        
        if wb is None:
            raise RuntimeError(f"O Excel não conseguiu carregar o ficheiro: {caminho_excel_abs}")
            
        # Exporta o livro inteiro (Capa + Rateio) para PDF
        wb.ExportAsFixedFormat(0, caminho_pdf_abs)
        
    except Exception as e:
        raise RuntimeError(f"Falha ao exportar para PDF: {str(e)}")
        
    finally:
        # Garante o fecho limpo do Excel da memória
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass