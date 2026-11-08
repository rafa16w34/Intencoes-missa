import io
import os
import tempfile
import subprocess
from .docx_generator import gerar_docx_missa


def gerar_pdf_missa(missa):
    """
    Gera o documento PDF formatado para a Missa informada,
    mantendo rigorosamente a mesma formatação visual do arquivo Word (.docx).
    Utiliza o LibreOffice em modo headless para renderização oficial idêntica.
    Retorna um buffer io.BytesIO pronto para download.
    """
    docx_buffer = gerar_docx_missa(missa)

    with tempfile.TemporaryDirectory() as temp_dir:
        docx_path = os.path.join(temp_dir, "documento.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_buffer.getvalue())

        user_profile_dir = os.path.join(temp_dir, "lo_profile")
        os.makedirs(user_profile_dir, exist_ok=True)

        cmd = [
            "libreoffice",
            f"-env:UserInstallation=file://{user_profile_dir}",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", temp_dir,
            docx_path
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        pdf_path = os.path.join(temp_dir, "documento.pdf")

        if result.returncode != 0 or not os.path.exists(pdf_path):
            raise RuntimeError(f"Erro na conversão do documento para PDF via LibreOffice: {result.stderr}")

        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

    buffer = io.BytesIO(pdf_bytes)
    buffer.seek(0)
    return buffer
