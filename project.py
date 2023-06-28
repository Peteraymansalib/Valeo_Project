# ---------------------------------------------------
# Author: Peter Ayman
# Date: 3 Nov 2022
# Valeo Testing Academy
# Project: Replace words at XML file from excel sheet
# ---------------------------------------------------

# Import libraries
import shutil
import fileinput
import openpyxl

# Import excel file
wb = openpyxl.load_workbook('Generationfile.xlsx')

# Read the excel sheet
ws1 = wb['Sheet1']

# Search on range to copy the xml file with new names
for nRow in ws1['A2':'X6']:
    colnumber = 0
    while colnumber < ws1.max_column:
        if colnumber == 0:
            sourcerPath = nRow[colnumber].value+'.xml'
            filename = nRow[colnumber+1].value +'.xml'
            shutil.copyfile(sourcerPath, filename) # Coping a new xml file
        else:
            if nRow[colnumber].value == None: # If you don't find values go to next line
                break
            else:
                with fileinput.FileInput(filename, inplace=True) as f:
                    for line in f:
                        if str(nRow[colnumber].value) in line: # Search at the xml file
                            print(line.replace(str(nRow[colnumber].value), str(nRow[colnumber+1].value)), end='') # Replace with new values
                        else:
                            print(line, end='')
        colnumber += 2 # Goes to the next sample at excel sheet

