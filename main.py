from openpyxl import *
from tkinter import *

from openpyxl.styles import Alignment

wb = load_workbook("C:\\Users\\snoew\\Desktop\\demo.xlsx")

sheet = wb.active

def excel():
    # Set column dimensions



    # Merge cells
    sheet.merge_cells('A1:A2')
    sheet.merge_cells('B1:B2')
    sheet.merge_cells('C1:D1')

    # Give them values
    sheet['A1'].value = "Dag"
    sheet['B1'].value = "Köpare, säljare, varuslag etc."
    sheet['C1'].value = "Kassa"
    sheet['C2'].value = "Inbetalningar"
    sheet['D2'].value = "Utbetalningar"

    # Center the values
    sheet['A1'].alignment = Alignment(horizontal='center')
    sheet['B1'].alignment = Alignment(horizontal='center')
    sheet['C1'].alignment = Alignment(horizontal='center')
    sheet['C2'].alignment = Alignment(horizontal='center')
    sheet['D2'].alignment = Alignment(horizontal='center')



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
    wb.save("C:\\Users\\snoew\\Desktop\\demo.xlsx")

    # Save button

    root.mainloop()

