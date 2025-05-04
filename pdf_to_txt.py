import fitz  # PyMuPDF
import re
import os
def read_pdf_to_txt(pdf_path):
    doc = fitz.open(pdf_path)
    record_list = []
    for page_num, page in enumerate(doc, start=1):
        text = page.get_text()
        lines = text.splitlines()
        matches = re.findall(r'(Line# \d+\s+\d+)', text)
        for match in matches:
            line_number = int(re.findall(r'Line# (\d+)\s+\d+', match)[0])
            if len(record_list) == 0:
                record_list.append([line_number,match])
            elif line_number - 1 == record_list[-1][0] or line_number == record_list[-1][0]:
                record_list.append([line_number, match])
            else:
                break
    record_list2 = []
    for line in record_list:
        message_id = re.findall(r'Line# \d+\s+(\d+)', line[1])[0]
        if message_id != "628":
            record_list2.append(line[1].replace("\n", " "))
    for line in record_list:
        message_id = re.findall(r'Line# \d+\s+(\d+)', line[1])[0]
        if message_id == "628":
            record_list2.append(line[1].replace("\n", " "))

    file = open(pdf_path[:-4] + ".txt", "w")
    file.truncate()
    file.close
    file = open(pdf_path[:-4] + ".txt", "a")
    for line in record_list2:
        file.write(line + "\n")
    file.close


if __name__ == "__main__":
    pdf_file = "sample.pdf"
    read_pdf_to_txt(pdf_file)
