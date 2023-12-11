import io
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def add_text_box(input_pdf, output_pdf, text, x, y, font_name, font_size):
    # Read the original PDF file
    pdf_reader = PdfReader(input_pdf)
    pdf_writer = PdfWriter()

    # Register the Kaiu font
    pdfmetrics.registerFont(TTFont(font_name, "kaiu.ttf"))

    # Get the page size
    width, height = float(pdf_reader.pages[0].mediabox.upper_right[0]), float(
        pdf_reader.pages[0].mediabox.upper_right[1]
    )

    # Calculate text width and height
    text_width = pdfmetrics.stringWidth(text, font_name, font_size)
    text_height = font_size

    # Calculate the position of the text box
    x_position = width - x - text_width
    y_position = height - y - text_height

    # Add the text box to the first page
    for i, page in enumerate(pdf_reader.pages):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(width, height))
        can.setFont(font_name, font_size)
        can.setFillColorRGB(0, 0, 0)  # Set fill color to black
        can.drawString(x_position, y_position, text)
        can.save()

        packet.seek(0)
        new_pdf = PdfReader(packet)

        # Only merge the text box on the first page
        if i == 0:
            page.merge_page(new_pdf.pages[0])

        pdf_writer.add_page(page)

    # Save the modified PDF file
    with open(output_pdf, "wb") as output_file:
        pdf_writer.write(output_file)


# Specify the PDF file path and the text to add
input_pdf_path = "test.pdf"
output_pdf_path = "output.pdf"
text_to_add = "黃志祥弟兄 賢伉儷"
font_name = "標楷體"
font_size = 16

# Specify the position of the text box (offset from the upper right corner)
x_offset = 20
y_offset = 10

# Call the function to add the text box
add_text_box(
    input_pdf_path,
    output_pdf_path,
    text_to_add,
    x_offset,
    y_offset,
    font_name,
    font_size,
)
