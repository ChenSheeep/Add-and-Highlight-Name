import io
import os
import fitz  # PyMuPDF
import re
import os
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def parse_member_file(filepath):
    """
    Parse the member file and return a dictionary containing filenames and associated members or representatives.
    """
    # Open the file and read its content
    with open(filepath, "r", encoding="utf-8") as file:
        content = file.readlines()

    # Initialize a dictionary to store filenames and corresponding members or representatives
    name_dict = {}

    # Iterate through each line in the file
    for line in content:
        # Split the line based on ":"
        parts = line.split(":")
        filename = parts[0].strip()

        representative = filename.split()[0]
        # Check if the filename contains English characters
        contains_english = re.search("[a-zA-Z]", filename)

        if not contains_english:
            # If no English characters, take the second word onward as the representative
            if len(representative) == 3:
                representative = representative[1:]
            elif len(representative) == 4:
                representative = representative[2:]

        # Check if the line contains member information
        if len(parts) == 2:
            # If yes, split the members by ", " and store in the dictionary
            members = parts[1].strip().split(", ")
            name_dict[filename] = members
        else:
            # If no member information, store the representative in the dictionary
            name_dict[filename] = [representative]
    return name_dict


def highlight_keywords(pdf_path, keywords, output_path):
    """
    Highlight keywords in the first page of a PDF and save the modified PDF.
    """
    # Open the PDF document
    pdf_document = fitz.open(pdf_path)

    # Access the first page
    page = pdf_document[0]

    # Initialize a flag to check if any instances of any keyword were found
    keyword_found = False

    # Iterate through the list of keywords
    for keyword in keywords:
        # Search for the keyword on the page
        keyword_instances = page.search_for(keyword)

        # Check if any instances of the keyword were found
        if keyword_instances:
            # Set the flag to indicate that at least one keyword was found
            keyword_found = True

            # Iterate through each instance and add a highlight annotation
            for inst in keyword_instances:
                highlight = page.add_highlight_annot(inst)

                # Set highlight color (here set to yellow), only set stroke color
                highlight.set_colors({"stroke": (1, 1, 0)})

    # Save the modified PDF only if at least one keyword was found
    if keyword_found:
        pdf_document.save(output_path)

    pdf_document.close()


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

    # Add the text box to the first page only
    for i, page in enumerate(pdf_reader.pages):
        if i == 0:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(width, height))
            can.setFont(font_name, font_size)
            can.setFillColorRGB(0, 0, 0)  # Set fill color to black
            can.drawString(x_position, y_position, text)
            can.save()

            packet.seek(0)
            new_pdf = PdfReader(packet)

            # Merge the text box on the first page only
            page.merge_page(new_pdf.pages[0])

        pdf_writer.add_page(page)

    # Save the modified PDF file
    with open(output_pdf, "wb") as output_file:
        pdf_writer.write(output_file)


def process_pdfs(folder):
    # Process each PDF file in the output folder
    for filename in os.listdir(folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder, filename)

            # Remove leading and trailing whitespaces, remove file extension
            text_to_add = os.path.splitext(filename)[0].strip()

            # Specify the position of the text box (offset from the upper right corner)
            x_offset = 20
            y_offset = 10

            # Call the function to add the text box and overwrite the existing file
            add_text_box(
                pdf_path,
                pdf_path,
                text_to_add,
                x_offset,
                y_offset,
                font_name="標楷體",
                font_size=16,
            )


# Print the final dictionary
filepath = "members.txt"
pdf_path = "test.pdf"

# Create the output folder if it doesn't exist
output_folder = "output"
os.makedirs(output_folder, exist_ok=True)
name_dict = parse_member_file(filepath)

# Iterate through the filenames and associated members or representatives
for filename, members_or_representative in name_dict.items():
    # Generate the output path for the modified PDF
    output_path = f"{output_folder}/{filename}.pdf"

    # Highlight keywords on the first page and save the modified PDF
    highlight_keywords(pdf_path, members_or_representative, output_path)


# Specify the output folder
folder = "output"

# Call the function to process PDF files in the output folder and overwrite existing files
process_pdfs(folder)
