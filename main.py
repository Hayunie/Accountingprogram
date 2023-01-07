import os
from tkinter import filedialog

import openpyxl
from openpyxl import *
from tkinter import *
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

path = "C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\test1.xlsx"
wb = load_workbook(path)


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
    receipt_entries.day_field.focus_set()

def focus2(event):
    receipt_entries.brutto_field.focus_set()

def focus3(event):
    receipt_entries.momsPercent_field.focus_set()

def clear():
    # Clear every field in the GUI
    receipt_entries.receipt_name_field.delete(0, END)
    receipt_entries.day_field.delete(0, END)
    receipt_entries.brutto_field.delete(0, END)
    receipt_entries.momsPercent_field.delete(0, END)

def insert():
    # Take the data from the GUI and write to excel file
    if (receipt_entries.receipt_name_field.get() == "" and
            receipt_entries.day_field.get() == "" and
            receipt_entries.month_field.get() == "" and
            receipt_entries.brutto_field.get() == "" and
            receipt_entries.momsPercent_field.get() == ""):
        print("empty input")

    else:
        current_row = wb.active.max_row
        current_column = wb.active.max_column
        current_sub_type = receipt_entries.sub_type_field.get()
        brutto = float(receipt_entries.brutto_field.get())
        momsPercent = float(receipt_entries.momsPercent_field.get())
        momsKr = float(brutto * (momsPercent / 100))
        netto = float(brutto - momsKr)
        verif_nr =+ 1

        wb.active.cell(row=current_row + 1, column=2).value = receipt_entries.receipt_name_field.get()
        wb.active.cell(row=current_row + 1, column=1).value = receipt_entries.day_field.get()
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

    def browseFiles():
        filename = filedialog.askopenfilename(initialdir= "/", title="Select a file",
                                              filetypes=(("Excel files","*.xlsx"),("all files","*.*")))
        wb = load_workbook(filename)
        root.title(os.path.basename(filename))


    def newFileWindow():
        """def newfile():
            wb = Workbook()
            filepath = f"C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\{entry}.xlxs"
            wb.save(filepath)
            #sheets()
            nf.destroy()"""
        nf= Toplevel(root)
        nf.geometry("300x300")
        nf.configure(bg='light blue')
        Default_Label(nf, text="Not yet implemented").pack()
        Default_Button(nf, text="Okej", command=nf.destroy).pack()
        """Default_Label(nf, text="Välj namn på filen: ").pack()
        entry = Entry(nf, font=12).pack()
        Default_Button(nf, text="Okej!", command=newfile).pack()"""

    root = Tk()
    menu = Menu(root)
    root.config(menu=menu)
    filemenu = Menu(menu)
    menu.add_cascade(label='File', menu=filemenu)
    filemenu.add_command(label='New', command=newFileWindow)
    filemenu.add_command(label='Open...', command=browseFiles)
    filemenu.add_separator()
    filemenu.add_command(label='Exit', command=root.quit)

    # Set title of GUI Window
    root.title()
    # Set the config of GUI window
    root.geometry("500x550")

    root.configure(bg='light blue')


    sub_type_headers1 = ['----', 'Inventarier försäljning', 'Nötkreatur', 'Svin', 'Skog och skogsprodukter',
                         'Övriga inbetalningar']
    sub_type_headers2 = ['----', 'Inventarier inköp', 'Inköp djur', 'Omkostnader skogen', 'Omkostnader djurskötseln',
               'Drivmedel, eldnings och smörolja', 'Underhåll inventarier', 'Kontorskostnader, bokföring, telefon',
               'Försäkrings premier', 'Övriga utbetalningar', ' Underhåll näringsfastigheter ekonomibyggnader',
               'Underhåll näringsfastigheter bostäder (inkl moms)', 'Underhåll markanläggning']

    months_headers = ["Januari", "Februari", "Mars", "April", "Maj", "Juni",
              "Juli", "Augusti", "September", "Oktober", "November", "December"]


    # Create frame
    mainframe = Frame(root, bg='light blue')
    mainframe.pack(expand=True, fill="both")

    frame2 = Frame(root, bg='light blue')
    frame2.pack(expand=True, fill="both")

    frame3 = Frame(root, bg='light blue')
    frame3.pack(expand=True, fill="both")

    frame4 = Frame(root, bg='light blue')
    frame4.pack(expand=True, fill="both")


    class Default_Label(Label):
        def __init__(self, *args, **kwargs):
            Label.__init__(self, *args, **kwargs)
            self['bg'] = 'light blue'
            self['font'] = 12


    class Default_Button(Button):
        def __init__(self, *args, **kwargs):
            Button.__init__(self, *args, **kwargs)
            self['font'] = 14
            self['width'] = 15

    Default_Label(mainframe, text='Välj månad').pack(side="top", pady=5)
    Default_Button(mainframe, text= months_headers[0], command= set_jan).pack(pady=5)
    Default_Button(mainframe, text= months_headers[1], command= set_feb).pack(pady=5)
    Default_Button(mainframe, text= months_headers[2], command= set_mar).pack(pady=5)
    Default_Button(mainframe, text= months_headers[3], command= set_apr).pack(pady=5)
    Default_Button(mainframe, text= months_headers[4], command= set_may).pack(pady=5)
    Default_Button(mainframe, text= months_headers[5], command= set_jun).pack(pady=5)
    Default_Button(mainframe, text= months_headers[6], command= set_jul).pack(pady=5)
    Default_Button(mainframe, text= months_headers[7], command= set_aug).pack(pady=5)
    Default_Button(mainframe, text= months_headers[8], command= set_sep).pack(pady=5)
    Default_Button(mainframe, text= months_headers[9], command= set_oct).pack(pady=5)
    Default_Button(mainframe, text= months_headers[10], command= set_nov).pack(pady=5)
    Default_Button(mainframe, text= months_headers[11], command= set_dec).pack(pady=5)

    def clearFrame():
        for widget in mainframe.winfo_children():
            widget.destroy()

    def receipt_entries(month):
        clearFrame()

        # Set the config of GUI window
        root.geometry("500x550")

        # Create labels for every data entry
        Default_Label(mainframe, text=month).pack(padx=5, pady=5, side=TOP, anchor='n')
        v = IntVar()
        Radiobutton(mainframe, text="Inbetalning", font=14, variable=v, value=1, bg='light blue').pack(side=TOP, anchor='n', padx=50)
        Radiobutton(mainframe, text="Utbetalning", font=14, variable=v, value=2, bg='light blue').pack(side=TOP, anchor='n', padx=10)

        Default_Label(frame2, text='Köpare, Säljare, Varumärke etc.').pack(padx=5, pady=5, anchor='nw')
        receipt_name_field = Entry(frame2, font=12).pack(padx=5, pady=5)
        Default_Label(frame2, text='Dag').pack(side=LEFT, anchor='nw', padx=5, pady=5)
        day_field = Entry(frame2, width=5, font=12).pack(side=LEFT, anchor='nw', padx=5, pady=5)

        Default_Label(frame3, text="Sort").pack(anchor='nw', padx=5, pady=5)
        Default_Label(frame3, text="Dropdown").pack(anchor='nw', padx=5, pady=5)
        Default_Label(frame3, text="Pris").pack(anchor='nw', padx=5, pady=5)
        brutto_field = Entry(frame3, width=5, font=12).pack(padx=5, pady=5)
        Default_Label(frame3, text="Kr").pack(padx=5, pady=5)

        Default_Label(frame4, text="Moms").pack(anchor='nw', padx=5, pady=5)
        momsPercent_field = Entry(frame4, width=5, font=12).pack(padx=5, pady=5)
        Default_Label(frame4, text="%").pack(padx=5, pady=5)
        # Save button
        Default_Button(frame4, text= "Spara", command= insert).pack(pady=5)

        # Back Button
        Default_Button(frame4, text= "Tillbaka", command= root.mainloop).pack(pady=5)

        # Create a text entrybox for every data entry





        # Dropdown menus




        # Bind method to call for the focus function
        receipt_name_field.bind("<Return>", focus1)
        day_field.bind("<Return>", focus2)
        brutto_field.bind("<Return>", focus3)


        # sheets()
        wb.save(path)




    root.mainloop()
