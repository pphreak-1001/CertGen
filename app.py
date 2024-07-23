import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PyPDF2 import PdfReader, PdfWriter
import io

def draw_text(can, text, x, y, max_width, max_font_size, min_font_size):
    font_name = 'CustomFont'
    font_size = max_font_size
    
    while font_size >= min_font_size:
        can.setFont(font_name, font_size)
        text_width = can.stringWidth(text, font_name, font_size)
        if text_width <= max_width:
            can.drawString(x, y, text)
            return font_size
        font_size -= 1

    wrapped_text = simpleSplit(text, font_name, font_size, max_width)
    text_object = can.beginText(x, y)
    text_object.setFont(font_name, font_size)
    for line in wrapped_text:
        text_object.textLine(line)
    can.drawText(text_object)
    return font_size

# Function to generate certificate
def generate_certificate(name, paper, template_pdf):
    reader = PdfReader(template_pdf)
    writer = PdfWriter()
    packet = io.BytesIO()
    
    can = canvas.Canvas(packet, pagesize=letter)
    
    # Coordinates and maximum width for text
    name_x, name_y = 380, 305  # Adjust as needed
    paper_x, paper_y = 230, 240  # Adjust as needed
    max_width = 400  # Adjust based on your template
    max_font_size = 16
    min_font_size = 10
    

    draw_text(can, name, name_x, name_y, max_width, max_font_size, min_font_size)
    

    draw_text(can, paper, paper_x, paper_y, max_width, max_font_size, min_font_size)
    
    can.save()
    
    # Move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page = reader.pages[0]
    page.merge_page(new_pdf.pages[0])
    writer.add_page(page)
    output_filename = f'certificate_{name.replace(" ", "_")}.pdf'
    with open(output_filename, 'wb') as outputStream:
        writer.write(outputStream)
    print(f'Generated certificate for {name}')

def main():
    font_path = input("Enter the path to your custom font (.ttf): ")
    participant_file = input("Enter the path to the participant Excel file (.xlsx): ")
    template_pdf = input("Enter the path to the certificate template (.pdf): ")
    participants = pd.read_excel(participant_file)
    pdfmetrics.registerFont(TTFont('CustomFont', font_path))
    for index, row in participants.iterrows():
        name = row['name']
        paper = row['paper']
        generate_certificate(name, paper, template_pdf)

if __name__ == "__main__":
    main()
