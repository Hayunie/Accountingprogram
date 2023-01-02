from openpyxl import *
from tkinter import *

from openpyxl.styles import Alignment, Font, PatternFill

wb = load_workbook("C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\test1.xlsx")
sheet = wb.active
def excel():
    # Set column dimensions


    # Merge cells
    sheet.merge_cells('A1:F1')
    sheet.merge_cells('A2:A3')
    sheet.merge_cells('B2:B3')
    sheet.merge_cells('C2:D2')
    sheet.merge_cells('A4:B4')

    # Give the headers values
    sheet['A1'].value = "Fördelning av"
    sheet['A2'].value = "Dag"
    sheet['A4'].value = 'SUMMA'
    sheet['B2'].value = "Köpare, säljare, varuslag etc."
    sheet['C2'].value = "Kassa"
    sheet['C3'].value = "Inbetalningar"
    sheet['D3'].value = "Utbetalningar"


    # Change the font
    wb.font = Font(size = 12)
    for cell in sheet[4:4]:
        cell.font = Font(size = 14, bold = True)
    sheet['C2'].font = Font(size = 12, bold = True)

    # Center the values
    wb.alignment = Alignment(horizontal='left')
    for cell in sheet[1]:
        cell.alignment = Alignment(horizontal='center')
    for cell in sheet[2]:
        cell.alignment = Alignment(horizontal='center')

    # Background color
    for x in range(5, sheet.max_row+1):
        c = sheet.cell(row=x, column=1)
        if x % 2 != 0:
            c.fill = PatternFill(start_color="ADD8E6", fill_type= "solid")


# Set focus(event) for every field
def focus1(event):
    day_field.focus_set()
def focus2(event):
    month_field.focus_set()
def focus3(event):
    main_type_field.focus_set()
def focus4(event):
    sub_type_field.focus_set()
def focus5(event):
    price_field.focus_set()
def focus6(event):
    moms_field.focus_set()

#def clear():
    # Clear every field in the GUI
    receipt_name_field.delete(0, END)
    day_field.delete(0, END)
    month_field.delete(0, END)
    main_type_field.delete(0, END)
    sub_type_field.delete(0, END)
    price_field.delete(0, END)
    moms_field.delete(0, END)

#def insert():
    # Take the data from the GUI and write to excel file

    # Get method to return current text as strinf and write it into excel
    # at particular location

    #wb.save("C:\\Users\\snoew\\Desktop\\demo.xlsx")

    # set focus at first first field with focus_set()

    # clear()

# def newFile()
    # add sheets for every month
    # excel()

# def openFile()

# def showFile()

# Driver code
if __name__ == "__main__":
    root = Tk()
    # Set background color
    root.configure(background='light blue')

    # Set title of GUI Window
    root.title("Bokföringsprogram")
    # Set the config of GUI window
    root.geometry("800x600")
    excel()

    # Create labels for every data entry
    header = Label(root, text="Kvitto", font=14, bg="light blue")
    receipt_name = Label(root, text="Köpare, Säljare, Varumärke etc.", font=12, bg="light blue")
    day = Label(root, text="Dag", font=12, bg="light blue", compound="left")
    month = Label(root, text="Månad", font=12, bg="light blue")
    main_type = Label(root, text="Inkomst eller Utgift", font=12, bg="light blue")
    sub_type = Label(root, text="Sort", font=12, bg="light blue")
    price = Label(root, text="Pris", font=12, bg="light blue")
    kr = Label(root, text="Kr", font=12, bg="light blue")
    moms = Label(root, text="Moms", font=12, bg="light blue")
    percent = Label(root, text="%", font=12, bg="light blue")

    # Grid method to place widgets at respective positions
    header.grid(row=0, column = 3)
    receipt_name.grid(row=1, column=0, sticky="w", columnspan=5)
    day.grid(row=2, column=0, sticky="w", columnspan=1)
    month.grid(row=2, column=3, sticky="w", columnspan=1)
    main_type.grid(row=3, column=0, sticky="w", columnspan=3)
    sub_type.grid(row=4, column=0, sticky="w", columnspan=1)
    price.grid(row=5, column=0, sticky="w", columnspan=1)
    kr.grid(row=5, column=3, sticky="w")
    moms.grid(row=6, column=0, sticky="w", columnspan=1)
    percent.grid(row=6, column=3,sticky="w")

    # Create a text entrybox for every data entry
    receipt_name_field = Entry(root, font=12)
    day_field = Entry(root, width=5, font=12)
    month_field = Entry(root, width=5, font=12)
    main_type_field = Entry(root, width=10, font=12)
    sub_type_field = Entry(root, width=5, font=12)
    price_field = Entry(root, width=5, font=12)
    moms_field = Entry(root, width=5, font=12)

    # Bind method to call for the focus function
    receipt_name_field.bind("<Return>", focus1)
    day_field.bind("<Return>", focus2)
    month_field.bind("<Return>", focus3)
    main_type_field.bind("<Return>", focus4)
    sub_type_field.bind("<Return>", focus5)
    price_field.bind("<Return>", focus6)

    # Grid method to place entry
    receipt_name_field.grid(row=1, column=5)
    day_field.grid(row=2, column=1, sticky="w", ipadx=2)
    month_field.grid(row=2, column=4, sticky="w", ipadx=2)
    main_type_field.grid(row=3, column=3, sticky="w", ipadx=6, columnspan=2)
    sub_type_field.grid(row=4, column=1, sticky="w", ipadx=2)
    price_field.grid(row=5, column=1, sticky="w", ipadx=2)
    moms_field.grid(row=6, column=1, sticky="w", ipadx=2)


    excel()
    wb.save("C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\test1.xlsx")

    # Save button
    save = Button(root, text="Spara", bg="Yellow")
    save.grid(row=7, column=3)

    root.mainloop()

