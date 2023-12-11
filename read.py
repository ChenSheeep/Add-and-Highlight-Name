import re

# Open the file and read its content
with open("members.txt", "r", encoding="utf-8") as file:
    content = file.readlines()

# Initialize a dictionary to store names and members
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
        member = parts[1].strip().split(", ")
        name_dict[filename] = member
    else:
        # If no member information, store the representative in the dictionary
        name_dict[filename] = representative

# Print the final dictionary
print(name_dict)
