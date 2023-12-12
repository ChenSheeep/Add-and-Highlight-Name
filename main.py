import io
import os
import fitz  # PyMuPDF
import re
import sys
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyQt5 import QtWidgets, QtCore
from scheduleUI import Ui_widget


def load_name():
    global filepath, name_dict
    try:
        abs_path = QtWidgets.QFileDialog.getOpenFileName(
            None, "選擇輸出名單", ".", "Text Files (*.txt)"
        )[0]
        current_dir = os.getcwd()

        # 計算相對路徑
        filepath = os.path.relpath(abs_path, current_dir)
        ui.name_label.setText(filepath)

        # Parse the member file and create a dictionary
        name_dict = parse_member_file(filepath)
    except:
        pass


def load_schedule():
    global pdf_path, output_folder
    try:
        abs_path = QtWidgets.QFileDialog.getOpenFileName(
            None, "選擇輸出名單", ".", "PDF Files (*.pdf)"
        )[0]
        current_dir = os.getcwd()

        # 計算相對路徑
        pdf_path = os.path.relpath(abs_path, current_dir)
        ui.schedule_label.setWordWrap(True)
        ui.schedule_label.setText(pdf_path)

        output_folder = pdf_path.replace(".pdf", "")
    except:
        pass


def parse_member_file(filepath):
    """
    Parse the member file and return a dictionary containing filenames and associated members or representatives.
    """
    with open(filepath, "r", encoding="utf-8") as file:
        content = file.readlines()

    name_dict = (
        {}
    )  # Initialize a dictionary to store filenames and corresponding members or representatives

    for line in content:
        parts = line.split(":")
        filename = parts[0].strip()

        fullname = filename.split()[0]
        contains_english = re.search("[a-zA-Z]", filename)

        if contains_english:
            # Extract English name and representative
            filename = filename.split("(")[1].replace(") ", " ")
            representative = [fullname.split("(")[0][-2:]]
            representative.append(fullname.split("(")[0])
        else:
            # Extract representative from non-English filename
            representative = [fullname[-2:]]
            representative.append(fullname)

        if len(parts) == 2:
            # Extract members if present in the line
            m_fullname = parts[1].strip().split(", ")
            member = []
            for name in m_fullname:
                member.append(name[-2:])
                member.append(name)
                if len(name) == 2:
                    spacename = f"{name[0]}  {name[1]}"
                    member.append(spacename)
            name_dict[filename] = member
        else:
            # If no member information, store the representative in the dictionary
            if len(fullname) == 2:
                spacename = f"{fullname[0]}  {fullname[1]}"
                representative.append(spacename)
            name_dict[filename] = representative
    return name_dict


def highlight_keywords(pdf_path, keywords, output_path):
    """
    Highlight keywords in the first page of a PDF and save the modified PDF.
    """
    pdf_document = fitz.open(pdf_path)
    page = pdf_document[0]
    keyword_found = False

    for keyword in keywords:
        # Search for the keyword on the page
        keyword_instances = page.search_for(keyword)

        if keyword_instances:
            # Set the flag to indicate that at least one keyword was found
            keyword_found = True

            # Iterate through each instance and add a highlight annotation
            for inst in keyword_instances:
                highlight = page.add_highlight_annot(inst)
                # Set highlight color (here set to yellow)
                highlight.set_colors({"stroke": (1, 1, 0)})

    # Save the modified PDF only if at least one keyword was found
    if keyword_found:
        pdf_document.save(output_path)

    pdf_document.close()


def add_text_box(target_pdf, text, x, y, font_name, font_size):
    """
    Add a text box to the specified location in a PDF.
    """
    pdf_reader = PdfReader(target_pdf)
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
            # Create a PDF canvas to add text
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
    with open(target_pdf, "wb") as output_file:
        pdf_writer.write(output_file)

    # Rename the file by extracting the first word from the original filename
    directory = os.path.dirname(target_pdf)
    filename = os.path.basename(target_pdf).replace(".pdf", "")
    new_filename = filename.split()[0] + ".pdf"
    renamed_pdf = os.path.join(directory, new_filename)
    os.rename(target_pdf, renamed_pdf)


def process_pdfs(folder):
    """
    Process each PDF file in the specified folder.
    """
    for filename in os.listdir(folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder, filename)

            # Remove leading and trailing whitespaces, remove file extension
            text_to_add = os.path.splitext(filename)[0].strip()

            # Specify the position of the text box (offset from the upper right corner)
            x_offset, y_offset = 20, 10

            # Call the function to add the text box and overwrite the existing file
            add_text_box(
                pdf_path, text_to_add, x_offset, y_offset, font_name="標楷體", font_size=16
            )


# File paths
# pdf_path = "names.txt"
# pdf_path = "template.pdf"
# output_folder = "output"

# # Create the output folder if it doesn't exist
# os.makedirs(output_folder, exist_ok=True)

# # Delete all files in the output folder
# for file_name in os.listdir(output_folder):
#     pdf_path = os.path.join(output_folder, file_name)
#     try:
#         if os.path.isfile(pdf_path):
#             os.remove(pdf_path)
#     except Exception as e:
#         print(f"Error deleting {pdf_path}: {e}")


# # Parse the member file and create a dictionary
# name_dict = parse_member_file(pdf_path)
def main():
    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Delete all files in the output folder
    for file_name in os.listdir(output_folder):
        filepath = os.path.join(output_folder, file_name)
        try:
            if os.path.isfile(filepath):
                os.remove(filepath)
        except Exception as e:
            print(f"Error deleting {filepath}: {e}")

    # Iterate through the filenames and associated members or representatives
    for filename, members_or_representative in name_dict.items():
        filename += ".pdf"
        output_path = os.path.join(output_folder, filename)

        # Highlight keywords on the first page and save the modified PDF
        highlight_keywords(pdf_path, members_or_representative, output_path)

    # Process PDF files in the output folder and overwrite existing files
    process_pdfs(output_folder)

    mbox = QtWidgets.QMessageBox(widget)
    mbox.information(widget, "提示訊息", f"全數輸出完畢！")
    sys.exit()


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)
    widget = QtWidgets.QWidget()
    ui = Ui_widget()
    ui.setupUi(widget)

    ui.load_name.clicked.connect(load_name)
    ui.load_schedule.clicked.connect(load_schedule)
    ui.start.clicked.connect(main)

    widget.show()
    app.exec_()
    # sys.exit(app.exec_())
