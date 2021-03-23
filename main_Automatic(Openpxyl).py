from openpyxl import load_workbook
from tkinter import filedialog

file = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("Any file", "*")))
try:
    workbook = load_workbook(file)
    sheet = workbook.active
    # publisher = input("Enter the publisher name: ")
    # book_list = []
    # publisher_list = []
    for i in range(2, sheet.max_row + 1):
        n = f"D{i}"
        # b = f"B{i}"

        publisher = sheet[n].value
        for j in range(3, sheet.max_row + 1):
            n1 = f"D{j}"
            b1 = f"B{i}"
            if publisher in sheet[n1].value:
                sheet[n1] = publisher
                # publisher_list.append(sheet[n1].value)
                # book_list.append(sheet[b1].value)
            else:
                continue

    workbook.save("saved.xlsx")
    # print("List of books whos' Publishers are changed: ", book_list)
    # print(publisher_list)

except FileNotFoundError:
    print("Please enter a valid file path!")


