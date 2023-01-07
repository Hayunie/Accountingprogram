import os

import openpyxl
from openpyxl import *
from tkinter import *
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

path = "C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\test1.xlsx"
wb = load_workbook(path)
#wb = openpyxl.Workbook()


def sheets():
    months = ["Januari", "Februari", "Mars", "April", "Maj", "Juni",
              "Juli", "Augusti", "September", "Oktober", "November", "December"]
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
    brutto_field.focus_set()

def focus3(event):
    momsPercent_field.focus_set()

def clear():
    # Clear every field in the GUI
    receipt_name_field.delete(0, END)
    day_field.delete(0, END)
    month_field.delete(0, END)
    var1.set(main_type_headers[0])
    sub_type_field.delete(0, END)
    brutto_field.delete(0, END)
    momsPercent_field.delete(0, END)

def insert():
    # Take the data from the GUI and write to excel file
    if (receipt_name_field.get() == "" and
            day_field.get() == "" and
            month_field.get() == "" and
            brutto_field.get() == "" and
            momsPercent_field.get() == ""):
        print("empty input")

    else:
        current_month = receipt_entries.month_field.get()
        # Set active sheet to current_month
        wb.active = wb[current_month]

        current_row = wb.active.max_row
        current_column = wb.active.max_column
        current_main_type = var1.get()
        current_sub_type = sub_type_field.get()
        brutto = float(brutto_field.get())
        momsPercent = float(momsPercent_field.get())
        momsKr = float(brutto * (momsPercent / 100))
        netto = float(brutto - momsKr)
        verif_nr =+ 1


        wb.active.cell(row=current_row + 1, column=2).value = receipt_name_field.get()
        wb.active.cell(row=current_row + 1, column=1).value = day_field.get()
        wb.active.cell(row=current_row + 1, column=3).value = verif_nr
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

    # wb.save("C:\\Users\\snoew\\Desktop\\demo.xlsx")

    # set focus at first first field with focus_set()

    # clear()


# def newFile()
# wb = openpyxl.Workbook()
# sheets()
# file_name = input("Name of file: ")
#path = file_name.xlsx
#saveas(path)

# Save file
# def saveFile():
    # if path == None:
        # path = asksaveasfilename(initialfile='Untitled.xlsx', defaultextension=".xlsx",
                                #filetypes=[("All Files", "*.*"),("Excel Documents","*.xlsx")])
        # if path == "":
            #path = None
        # else:
            ## Try to save the file
            # file = open(path,"w")
            # file.write(


# def openFile()

# def showFile()

# Driver code
if __name__ == "__main__":
    def set_jan():
        month = months_headers[0]
        wb.active = 1
        receipt_entries(month)
    def set_feb():
        month = months_headers[1]
        wb.active = 2
        receipt_entries(month)
    def set_mar():
        month = months_headers[2]
        wb.active = 3
        receipt_entries(month)
    def set_apr():
        month = months_headers[3]
        wb.active = 4
        receipt_entries(month)
    def set_may():
        month = months_headers[4]
        wb.active = 5
        receipt_entries(month)
    def set_jun():
        month = months_headers[5]
        wb.active = 6
        receipt_entries(month)
    def set_jul():
        month = months_headers[6]
        wb.active = 7
        receipt_entries(month)
    def set_aug():
        month = months_headers[7]
        wb.active = 8
        receipt_entries(month)
    def set_sep():
        month = months_headers[8]
        wb.active = 9
        receipt_entries(month)
    def set_oct():
        month = months_headers[9]
        wb.active = 10
        receipt_entries(month)
    def set_nov():
        month = months_headers[10]
        wb.active = 11
        receipt_entries(month)
    def set_dec():
        month = months_headers[11]
        wb.active = 12
        receipt_entries(month)

    class Default_Label( Label):
        def __init__(self, text, *args, **kwargs):
            super().__init__()
            self['bg'] = 'light blue'
            self['font'] = 12
            self['text'] = text

    class Default_Button(Button):
        def __init__(self, text, command, *args, **kwargs):
            super().__init__()
            self['text'] = text
            self['command'] = command
            self['font'] = 14
            self['width'] = 15

    root = Tk()

    main_type_headers = ['----', 'Inbetalningar', 'Utbetalningar']
    sub_type_headers = ['----', 'Inventarier försäljning', 'Nötkreatur', 'Svin', 'Skog och skogsprodukter', 'Övriga inbetalningar',
               'Inventarier inköp', 'Inköp djur', 'Omkostnader skogen', 'Omkostnader djurskötseln',
               'Drivmedel, eldnings och smörolja', 'Underhåll inventarier', 'Kontorskostnader, bokföring, telefon',
               'Försäkrings premier', 'Övriga utbetalningar', ' Underhåll näringsfastigheter ekonomibyggnader',
               'Underhåll näringsfastigheter bostäder (inkl moms)', 'Underhåll markanläggning']

    months_headers = ["Januari", "Februari", "Mars", "April", "Maj", "Juni",
              "Juli", "Augusti", "September", "Oktober", "November", "December"]

    # Set background color
    root.configure(background='light blue')

    # Set title of GUI Window
    root.title(os.path.basename(path))
    # Set the config of GUI window
    root.geometry("375x500")
    # Create frame

    def receipt_entries(month):

        # Set the config of GUI window
        root.geometry("800x600")

        # Create labels for every data entry
        header = Default_Label(month)

        receipt_name = Default_Label("Köpare, Säljare, Varumärke etc.")
        receipt_name.pack(anchor='w')
        day = Default_Label("Dag")
        month = Default_Label("Månad")
        main_type = Default_Label("Inkomst eller Utgift")
        sub_type = Default_Label("Sort")
        brutto = Default_Label("Pris")
        kr = Default_Label("Kr")
        momsPercent = Default_Label("Moms")
        percent = Default_Label("%")

        # Grid method to place widgets at respective positions
        """header.grid(row=0, column=3)
        #receipt_name.grid(row=1, column=0, sticky="w", columnspan=5)
        day.grid(row=2, column=0, sticky="w", columnspan=1)
        month.grid(row=2, column=3, sticky="w", columnspan=1)
        main_type.grid(row=3, column=0, sticky="w", columnspan=3)
        sub_type.grid(row=4, column=0, sticky="w", columnspan=1)
        brutto.grid(row=5, column=0, sticky="w", columnspan=1)
        kr.grid(row=5, column=3, sticky="w")
        momsPercent.grid(row=6, column=0, sticky="w", columnspan=1)
        percent.grid(row=6, column=3, sticky="w")"""

        # Create a text entrybox for every data entry
        receipt_name_field = Entry(root, font=12)
        day_field = Entry(root, width=5, font=12)
        month_field = Entry(root, width=5, font=12)
        sub_type_field = Entry(root, width=5, font=12)
        brutto_field = Entry(root, width=5, font=12)
        momsPercent_field = Entry(root, width=5, font=12)

        # Dropdown menus

        var1 = StringVar()
        var1.set(main_type_headers[0])
        drop1 = OptionMenu(root, var1, *main_type_headers)
        #drop1.pack()
        drop1.config(width=15, font=12)
        #drop1.grid(row=3, column=4, sticky="w", columnspan=2)


        var2 = StringVar()
        var2.set(sub_type_headers[0])


        # Bind method to call for the focus function
        receipt_name_field.bind("<Return>", focus1)
        day_field.bind("<Return>", focus2)
        brutto_field.bind("<Return>", focus3)

        # Grid method to place entry
        """receipt_name_field.grid(row=1, column=5)
        day_field.grid(row=2, column=1, sticky="w", ipadx=2)
        month_field.grid(row=2, column=4, sticky="w", ipadx=2)
        sub_type_field.grid(row=4, column=1, sticky="w", ipadx=2)
        brutto_field.grid(row=5, column=1, sticky="w", ipadx=2)
        momsPercent_field.grid(row=6, column=1, sticky="w", ipadx=2)"""

        # sheets()
        wb.save(path)

        # Save button
        save = Button(root, text="Spara", bg="Yellow")
        #save.grid(row=7, column=3)
        # Back Button
        back = Button(root, text="Tillbaka", bg ="Pink")
        #back.grid(row=7, column=0)




    choose_month = Label(root, text="Välj månad", font=16, bg="light blue")
    choose_month.grid(row=0, column=1)
    empty = Label(root, text="", bg="light blue", width=15)
    empty.grid(row=0, column=0)

    jan = Button(root, text=months_headers[0], font=14, width=15, command=set_jan)
    feb = Button(root, text=months_headers[1], font=14, width=15, command=set_feb)
    mar = Button(root, text=months_headers[2], font=14, width=15, command=set_mar)
    apr = Button(root, text=months_headers[3], font=14, width=15, command=set_apr)
    may = Button(root, text=months_headers[4], font=14, width=15, command=set_may)
    jun = Button(root, text=months_headers[5], font=14, width=15, command=set_jun)
    jul = Button(root, text=months_headers[6], font=14, width=15, command=set_jul)
    aug = Button(root, text=months_headers[7], font=14, width=15, command=set_aug)
    sep = Button(root, text=months_headers[8], font=14, width=15, command=set_sep)
    oct = Button(root, text=months_headers[9], font=14, width=15, command=set_oct)
    nov = Button(root, text=months_headers[10], font=14, width=15, command=set_nov)
    dec = Button(root, text=months_headers[11], font=14, width=15, command=set_dec)

    jan.grid(row=1, column=1)
    feb.grid(row=2, column=1)
    mar.grid(row=3, column=1)
    apr.grid(row=4, column=1)
    may.grid(row=5, column=1)
    jun.grid(row=6, column=1)
    jul.grid(row=7, column=1)
    aug.grid(row=8, column=1)
    sep.grid(row=9, column=1)
    oct.grid(row=10, column=1)
    nov.grid(row=11, column=1)
    dec.grid(row=12, column=1)


    root.mainloop()
