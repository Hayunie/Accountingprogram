from openpyxl import *
from tkinter import *

from openpyxl.styles import Alignment, Font, PatternFill
path = "C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\test1.xlsx"
wb = load_workbook(path)
sheet = wb.active
def sheets():
    sheet.title = "Januari"
    wb.create_sheet(title="Februari")
    wb.create_sheet(title="Mars")
    wb.create_sheet(title="April")
    wb.create_sheet(title="Maj")
    wb.create_sheet(title="Juni")
    wb.create_sheet(title="Juli")
    wb.create_sheet(title="Augusti")
    wb.create_sheet(title="September")
    wb.create_sheet(title="Oktober")
    wb.create_sheet(title="November")
    wb.create_sheet(title="December")

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

    # Formula for the total
    sheet['C4'] = '=SUM(C5:INDEX(C:C,ROWS(C:C)))'

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
    brutto_field.focus_set()
def focus6(event):
    momsPercent_field.focus_set()

#def clear():
    # Clear every field in the GUI
    receipt_name_field.delete(0, END)
    day_field.delete(0, END)
    month_field.delete(0, END)
    main_type_field.delete(0, END)
    sub_type_field.delete(0, END)
    brutto_field.delete(0, END)
    momsPercent_field.delete(0, END)

#def insert():
    # Take the data from the GUI and write to excel file
    if (receipt_name_field.get() == "" and
        day_field.get() == "" and
        month_field.get() == "" and
        main_type_field.get() == "" and
        sub_type_field.get() == "" and
        brutto_field.get() == "" and
        momsPercent_field.get() == ""):
        print("empty input")

    else:
        current_month = month_field.get()
       # Set active sheet to current_month

        current_row = sheet.max_row
        current_column = sheet.max_column
        current_main_type = main_type_field.get()
        brutto = float(brutto_field.get())
        momsPer = float(momsPercent_field.get())
        momsKr = brutto*momsPer
        netto = brutto-momsKr


        sheet.cell(row=current_row + 1, column=2).value = receipt_name_field.get()
        sheet.cell(row=current_row + 1, column=1).value = day_field.get()
        if current_main_type == "Inkomst":
            sheet.cell(current_row + 1, column=3).value = brutto_field.get()
        elif current_main_type == "Utgift":
            sheet.cell(current_row + 1, column=4).value = brutto_field.get()





    # Get method to return current text as strinf and write it into excel
    # at particular location

    #wb.save("C:\\Users\\snoew\\Desktop\\demo.xlsx")

    # set focus at first first field with focus_set()

    # clear()

# def newFile()
    # sheets()
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
    brutto = Label(root, text="Pris", font=12, bg="light blue")
    kr = Label(root, text="Kr", font=12, bg="light blue")
    momsPercent = Label(root, text="Moms", font=12, bg="light blue")
    percent = Label(root, text="%", font=12, bg="light blue")

    # Grid method to place widgets at respective positions
    header.grid(row=0, column = 3)
    receipt_name.grid(row=1, column=0, sticky="w", columnspan=5)
    day.grid(row=2, column=0, sticky="w", columnspan=1)
    month.grid(row=2, column=3, sticky="w", columnspan=1)
    main_type.grid(row=3, column=0, sticky="w", columnspan=3)
    sub_type.grid(row=4, column=0, sticky="w", columnspan=1)
    brutto.grid(row=5, column=0, sticky="w", columnspan=1)
    kr.grid(row=5, column=3, sticky="w")
    momsPercent.grid(row=6, column=0, sticky="w", columnspan=1)
    percent.grid(row=6, column=3,sticky="w")

    # Create a text entrybox for every data entry
    receipt_name_field = Entry(root, font=12)
    day_field = Entry(root, width=5, font=12)
    month_field = Entry(root, width=5, font=12)
    main_type_field = Entry(root, width=10, font=12)
    sub_type_field = Entry(root, width=5, font=12)
    brutto_field = Entry(root, width=5, font=12)
    momsPercent_field = Entry(root, width=5, font=12)

    # Bind method to call for the focus function
    receipt_name_field.bind("<Return>", focus1)
    day_field.bind("<Return>", focus2)
    month_field.bind("<Return>", focus3)
    main_type_field.bind("<Return>", focus4)
    sub_type_field.bind("<Return>", focus5)
    brutto_field.bind("<Return>", focus6)

    # Grid method to place entry
    receipt_name_field.grid(row=1, column=5)
    day_field.grid(row=2, column=1, sticky="w", ipadx=2)
    month_field.grid(row=2, column=4, sticky="w", ipadx=2)
    main_type_field.grid(row=3, column=3, sticky="w", ipadx=6, columnspan=2)
    sub_type_field.grid(row=4, column=1, sticky="w", ipadx=2)
    brutto_field.grid(row=5, column=1, sticky="w", ipadx=2)
    momsPercent_field.grid(row=6, column=1, sticky="w", ipadx=2)


    excel()
    wb.save(path)

    # Save button
    save = Button(root, text="Spara", bg="Yellow")
    save.grid(row=7, column=3)

    root.mainloop()

