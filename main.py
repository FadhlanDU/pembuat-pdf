import os
from fpdf import FPDF
from PIL import Image
import comtypes.client
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io

def convert_text_to_pdf(file_path, output_dir):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    with open(file_path, "r", encoding="utf-8") as file:
        for line in file:
            pdf.cell(200, 10, txt=line.strip(), ln=True)
    
    output_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}.pdf")
    pdf.output(output_file)
    return output_file

def convert_image_to_pdf(file_path, output_dir):
    image = Image.open(file_path)
    pdf_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}.pdf")
    image.convert("RGB").save(pdf_path, "PDF")
    return pdf_path

import comtypes.client

def convert_word_to_pdf(file_path, output_dir):
    word = comtypes.client.CreateObject("Word.Application")
    doc = word.Documents.Open(file_path)
    output_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}.pdf")
    doc.SaveAs(output_file, FileFormat=17)  # 17 is the format ID for PDF
    doc.Close()
    word.Quit()
    return output_file

def add_centered_image_watermark(input_pdf, output_pdf, watermark_image):
    """
    Tambahkan watermark gambar di bawah teks halaman PDF.

    :param input_pdf: Path file PDF input
    :param output_pdf: Path file PDF output
    :param watermark_image: Path gambar watermark
    """
    # Baca PDF input
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        # Dapatkan dimensi halaman
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        # Atur ukuran watermark (besar seperti di Word)
        watermark_width = page_width * 0.7  # 70% lebar halaman
        watermark_height = page_height * 0.5  # 50% tinggi halaman

        # Hitung posisi gambar agar berada di tengah
        x_center = (page_width - watermark_width) / 2
        y_center = (page_height - watermark_height) / 2

        # Buat watermark dengan reportlab
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(page_width, page_height))
        c.drawImage(watermark_image, 
                    x=x_center, 
                    y=y_center, 
                    width=watermark_width, 
                    height=watermark_height, 
                    mask='auto')
        c.save()

        # Tambahkan watermark di bawah teks
        packet.seek(0)
        watermark_pdf = PdfReader(packet)
        watermark_page = watermark_pdf.pages[0]
        watermark_page.merge_page(page)  # Gabungkan halaman asli di atas watermark

        # Tambahkan halaman yang dimodifikasi ke writer
        writer.add_page(watermark_page)

    # Simpan hasil PDF
    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)



def batch_convert_to_pdf(input_dir, output_dir, watermark_image):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for root, _, files in os.walk(input_dir):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().endswith(".txt"):
                output_pdf = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.pdf")
                convert_text_to_pdf(file_path, output_dir)
                add_centered_image_watermark(output_pdf, output_pdf, watermark_image)
            elif file.lower().endswith((".png", ".jpg", ".jpeg")):
                output_pdf = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.pdf")
                convert_image_to_pdf(file_path, output_dir)
                add_centered_image_watermark(output_pdf, output_pdf, watermark_image)
            elif file.lower().endswith(".docx"):
                output_pdf = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.pdf")
                convert_word_to_pdf(file_path, output_dir)
                add_centered_image_watermark(output_pdf, output_pdf, watermark_image)
            else:
                print(f"Skipping unsupported file: {file}")
    
    print(f"All files converted and watermarked. Check output directory: {output_dir}")


# Example usage:
input_directory = r"D:\Pembuat PDF\word"  # Ganti dengan folder file Anda
output_directory = r"D:\Pembuat PDF\pdf"     # Ganti dengan folder tujuan
watermark_image = r"D:\Pembuat PDF\image\logo ubp.png"

batch_convert_to_pdf(input_directory, output_directory, watermark_image)
