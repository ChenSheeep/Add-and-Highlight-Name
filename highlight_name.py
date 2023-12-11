import fitz  # PyMuPDF
import re
import os


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

        # Check if the filename contains English characters
        contains_english = re.search("[a-zA-Z]", filename)
        if contains_english:
            # If English characters are present, take the first word as the representative
            representative = filename.split()[0]
        else:
            # If no English characters, take the second word onward as the representative
            representative = filename.split()[0][1:]

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

    # Iterate through the list of keywords
    for keyword in keywords:
        # Search for the keyword on the page
        keyword_instances = page.search_for(keyword)

        # Check if any instances of the keyword were found
        if keyword_instances:
            # Iterate through each instance and add a highlight annotation
            for inst in keyword_instances:
                highlight = page.add_highlight_annot(inst)

                # Set highlight color (here set to yellow), only set stroke color
                highlight.set_colors({"stroke": (1, 1, 0)})

            # Save the modified PDF and return (no need to continue checking other keywords)
            pdf_document.save(output_path)
            pdf_document.close()
            return

    # If no instances of any keyword were found, do not save the modified PDF
    pdf_document.close()


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
