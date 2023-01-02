from openpyxl import *
from tkinter import *

from openpyxl.styles import Alignment, Font

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
    for cell in sheet[1]:
        cell.alignment = Alignment(horizontal='center')
    for cell in sheet[2]:
        cell.alignment = Alignment(horizontal='center')








# Set focus(event) for every field
# def focus(event):

#def clear():
    # Clear every field in the GUI

#def insert():
    # Take the data from the GUI and write to excel file

    # Get method to return current text as strinf and write it into excel
    # at particular location

    #wb.save("C:\\Users\\snoew\\Desktop\\demo.xlsx")

    # set focus at first first field with focus_set()

    # clear()

# def newFile()
    # add sheets for every month
    # add all the headers

# def openFile()

# def showFile()

# Driver code
if __name__ == "__main__":
    root = Tk()
    # Set background color
    # Set title of GUI Window
    # Set the config of GUI window
    # excel()

    # Create labels for every data entry
    # Grid method to place widgets at respective positions

    # Create a text entrybox for every data entry
    # Bind method to call for the focus function
    # Grid method to place entry boxes

    excel()
    wb.save("C:\\Users\\snoew\\OneDrive\\Skrivbord\\Projekt\\test1.xlsx")

    # Save button

    root.mainloop()

