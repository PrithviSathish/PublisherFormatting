import xlrd
# from xlutils.copy import copy
import xlsxwriter


def getting_publisher_list(x, y):
    file = x + ".xlsx"
    try:
        wb = xlrd.open_workbook(file)
        # rb = copy(wb)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0, 1)

        publisher_list = []
        for i in range(sheet.nrows):
            if i == 0:
                continue
            else:
                publisher_list.append(sheet.cell_value(i, 3))

        print("Original list: ", publisher_list)
        # s = rb.get_sheet(0)
        workbook = xlsxwriter.Workbook(file)
        worksheet = workbook.add_worksheet()
        for j in publisher_list:
            if y in j:
                index = publisher_list.index(j)
                publisher_list[index] = y
                # s.write(3, index + 1, y)
                worksheet.write(index + 1, 3, y)

                # rb.save("SampleSheet.xlsx")

        workbook.close()
        print("Modified list: ", publisher_list)

    except FileNotFoundError:
        print("Enter a valid file name!\n")
        main()


def main():
    name = input("Enter the File Path: ")
    publisher = input("Enter the name of the publisher: ")
    getting_publisher_list(name, publisher)


main()
