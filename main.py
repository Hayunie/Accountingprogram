from openpyxl import *
from tkinter import *
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

path = "C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\test1.xlsx"
wb = load_workbook(path)

def sheets():
    months = ["Januari", "Februari", "Mars", "April", "Maj", "Juni",
              "Juli", "Augusti", "September", "November", "December"]
    headers = ['Dag', 'Köpare, Säljare, Varuslag, etc.', 'Verif.\nnr', 'Inbetalningar', 'Utbetalningar', 'Utgående\nmoms',
               'Inventarier\nförsäljning', 'Nötkreatur', 'Svin', 'Skog och\nskogs-\nprodukter', 'Övriga \ninbetalningar',
               'Ingående\nmoms', 'Inventarier\ninköp', 'Inköp djur', 'Omkostnader\nskogen', 'Omkostnader\ndjurskötseln',
               'Drivmedel,\neldnings och\nsmörolja', 'Underhåll\ninventarier', 'Kontors-\nkostnader,\nbokföring,\ntelefon',
               'Försäkrings\npremier', 'Övriga\nutbetalningar', ' Underhåll\nnärings-\nfastigheter\nekonomi-\nbyggnader',
               'Underhåll\nnärings-\nfastigheter\nbostäder\n(inkl moms)', 'Underhåll\nmark-\nanläggning']

    for m in range(len(months)):
        temp = months[m]
        sheets = wb.create_sheet(title=temp)

        # Set Dimensions
        sheets.column_dimensions['A'].width = 5
        sheets.column_dimensions['B'].width = 30
        sheets.column_dimensions['C'].width = 7
        column = 4
        while column < 25:
            i = get_column_letter(column)
            sheets.column_dimensions[i].width = 15
            column += 1

        # Merge cells
        sheets.merge_cells('A1:F1')
        sheets.merge_cells('A3:B3')
        sheets.merge_cells('G1:K1')
        sheets.merge_cells('L1:X1')

        # Give the headers values
        sheets['A1'].value = "Fördelning av"
        sheets['A3'].value = "SUMMA"
        sheets['C3'].value = "---"
        sheets['G1'].value = "Inbetalningar"
        sheets['L1'].value = "Utbetalningar"
        for h in range(len(headers)):
            tempHead = headers[h]
            sheets.cell(row=2, column=h + 1).value = tempHead

        # Formula for the total
        for c in range(sheets.max_column - 3):
            formula = '=SUM(D5:INDEX(D:D,ROWS(D:D)))'
            sheets.cell(row=3, column=c + 4).value = formula

        # Set the font
        wb.font = Font(size=12)
        thick = Side(border_style="thick", color="000000")
        thin = Side(border_style="thin", color="000000")
        for row in sheets:
            for cell in row:
                cell.border = Border(left=thin, right=thin)
        for cell in sheets[1:1]:
            cell.font = Font(size=14, bold=True)
        for cell in sheets[2:2]:
            cell.border = Border(bottom=thin, left=thin, right=thin, top=thin)
        for cell in sheets[3:3]:
            cell.font = Font(size=14, bold=True)
            cell.border = Border(bottom=thick, top=thin, left=thin, right=thin)

        # Center the values
        wb.alignment = Alignment(horizontal='left')
        for cell in sheets[1]:
            cell.alignment = Alignment(horizontal='center')
        for cell in sheets[2]:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
        sheets['C3'].alignment = Alignment(horizontal='center')

        # Background color
        for x in range(5, sheets.max_row + 1):
            c = sheets.cell(row=x, column=1)
            if x % 2 != 0:
                c.fill = PatternFill(start_color="ADD8E6", fill_type="solid")

    del wb["Sheet1"]

def excel():
    thin = Side(border_style="thin", color="000000")
    for row in wb.active:
        for cell in row:
            cell.border = Border(left=thin, right=thin)

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

def clear():
    # Clear every field in the GUI
    receipt_name_field.delete(0, END)
    day_field.delete(0, END)
    month_field.delete(0, END)
    main_type_field.delete(0, END)
    sub_type_field.delete(0, END)
    brutto_field.delete(0, END)
    momsPercent_field.delete(0, END)

def insert():
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
        wb.active = wb[current_month]

        current_row = wb.active.max_row
        current_column = wb.active.max_column
        current_main_type = main_type_field.get()
        current_sub_type = sub_type_field.get()
        brutto = float(brutto_field.get())
        momsPercent = float(momsPercent_field.get())
        momsKr = float(brutto * (momsPercent / 100))
        netto = float(brutto - momsKr)
        current_entries = current_row - 4

        wb.active.cell(row=current_row + 1, column=2).value = receipt_name_field.get()
        wb.active.cell(row=current_row + 1, column=1).value = day_field.get()
        wb.active.cell(row=current_row + 1, column=3).value = current_entries
        if current_main_type == "Inkomst":
            wb.active.cell(current_row + 1, column=3).value = brutto_field.get()
            # Moms
            # Subtype
            # Netto
        elif current_main_type == "Utgift":
            wb.active.cell(current_row + 1, column=4).value = brutto_field.get()
            # Moms
            # Subtype
            # Netto
        # if current_sub_type == "":

    # Get method to return current text as strinf and write it into excel
    # at particular location

    # wb.save("C:\\Users\\snoew\\Desktop\\demo.xlsx")

    # set focus at first first field with focus_set()

    # clear()


# def newFile()
# sheets()


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
    header.grid(row=0, column=3)
    receipt_name.grid(row=1, column=0, sticky="w", columnspan=5)
    day.grid(row=2, column=0, sticky="w", columnspan=1)
    month.grid(row=2, column=3, sticky="w", columnspan=1)
    main_type.grid(row=3, column=0, sticky="w", columnspan=3)
    sub_type.grid(row=4, column=0, sticky="w", columnspan=1)
    brutto.grid(row=5, column=0, sticky="w", columnspan=1)
    kr.grid(row=5, column=3, sticky="w")
    momsPercent.grid(row=6, column=0, sticky="w", columnspan=1)
    percent.grid(row=6, column=3, sticky="w")

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

    sheets()
    wb.save(path)

    # Save button
    save = Button(root, text="Spara", bg="Yellow")
    save.grid(row=7, column=3)

    root.mainloop()
