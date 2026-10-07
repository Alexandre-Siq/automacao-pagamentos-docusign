import os
import shutil
import subprocess
import tempfile


def converter_excel_para_pdf(caminho_excel, caminho_pdf):
    """Converte o Excel preenchido em PDF.

    No Windows usa o Excel via win32com. Nos demais sistemas usa o
    LibreOffice em modo headless, para o lote não depender só do COM.
    """
    caminho_excel_abs = os.path.abspath(caminho_excel)
    caminho_pdf_abs = os.path.abspath(caminho_pdf)

    if not os.path.exists(caminho_excel_abs):
        raise FileNotFoundError(f"Ficheiro Excel não encontrado: {caminho_excel_abs}")

    if os.path.exists(caminho_pdf_abs):
        try:
            os.remove(caminho_pdf_abs)
        except Exception:
            pass

    pasta_pdf = os.path.dirname(caminho_pdf_abs)
    if pasta_pdf:
        os.makedirs(pasta_pdf, exist_ok=True)

    if os.name == "nt":
        _converter_via_excel(caminho_excel_abs, caminho_pdf_abs)
    else:
        _converter_via_libreoffice(caminho_excel_abs, caminho_pdf_abs)

    return caminho_pdf_abs


def _converter_via_excel(caminho_excel_abs, caminho_pdf_abs):
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(caminho_excel_abs, ReadOnly=True)

        if wb is None:
            raise RuntimeError(f"O Excel não conseguiu carregar o ficheiro: {caminho_excel_abs}")

        wb.ExportAsFixedFormat(0, caminho_pdf_abs)

    except Exception as e:
        raise RuntimeError(f"Falha ao exportar para PDF: {str(e)}")

    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass


def _converter_via_libreoffice(caminho_excel_abs, caminho_pdf_abs):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            "Conversão para PDF indisponível neste sistema. "
            "No Windows a exportação usa o Excel; aqui é preciso o LibreOffice (comando soffice)."
        )

    with tempfile.TemporaryDirectory() as pasta_tmp:
        resultado = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                pasta_tmp,
                caminho_excel_abs,
            ],
            capture_output=True,
            timeout=180,
        )
        if resultado.returncode != 0:
            detalhe = (resultado.stderr or resultado.stdout or b"").decode("utf-8", "replace").strip()
            raise RuntimeError(f"LibreOffice falhou ao converter para PDF: {detalhe}")

        nome_gerado = os.path.splitext(os.path.basename(caminho_excel_abs))[0] + ".pdf"
        caminho_gerado = os.path.join(pasta_tmp, nome_gerado)
        if not os.path.exists(caminho_gerado):
            raise RuntimeError("LibreOffice terminou sem gerar o PDF.")

        shutil.move(caminho_gerado, caminho_pdf_abs)
