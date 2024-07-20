from PyPDF2 import PdfReader, PdfWriter

def encrypt_pdf(file_path: str, password: str) -> str:
    """Encrypt a PDF file with a password."""
    pdf = PdfReader(file_path)
    output = PdfWriter()
    for page_num in range(len(pdf.pages)):
        output.add_page(pdf.pages[page_num])
    output_filename = f"{file_path[:-4]}.pdf"
    output.encrypt(password)
    with open(output_filename, "wb") as output_pdf:
        output.write(output_pdf)
    return output_filename