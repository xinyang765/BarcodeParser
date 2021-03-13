from tkinter import *
from tkinter import filedialog
import os
import pandas as pd
from datetime import date 
from openpyxl import load_workbook


root = Tk()
root.configure(background='green')
  
frame = Frame(root, padx=30, pady=30, bg='green')
frame.grid()

heading = Label(frame, text='Barcode Parser', font=('Arial', 30), fg='white', bg='green')  
heading.grid(row=0) 

barcodeEntryLabel = Label(frame, text ='Enter Barcode:', font=('Arial', 20), fg='white', bg='green')
barcodeEntryLabel.grid(row=1, pady=(50, 5))

barcodeText = Text(frame)
barcodeText.grid(row=2)


outputFileName = '/Products.xlsx'
outputFile = ''
barcodesInfoFile = './BarcodesInfo.xlsx'
outputSheetName = date.today().strftime('%m-%d-%Y')
PRODUCT_ID = 'Product ID'
PRODUCT_NAME = 'Product Name'
WEIGHT_POSITION = 'Weight Position (inclusive)'
WEIGHT = 'Weight'
NEXT_INDICATOR = 'CN'
parsedBarcodes = {}
saved = False


def getBarcodesInfo(filePath):
    if os.path.isfile(filePath):
        bcsInfo = pd.read_excel(filePath, dtype=str).to_dict('list')

        return bcsInfo


def getWeight(wp, barcode):
    wpIdx = wp.split('-')
    startIdx = int(wpIdx[0]) - 1
    endIdx = int(wpIdx[1])
    weight = barcode[startIdx:endIdx]
    return weight

def process(productNames, weights, totalWeight):
    productNames.append('')
    weights.append(totalWeight)
    productNames.append('')
    weights.append('')


def parse(barcodes):
    barcodesArr = barcodes.split('\n')
    print(barcodesArr)
    bcsInfo = getBarcodesInfo(barcodesInfoFile)
    productNames = []
    weights = []
    totalWeight = 0

    for bc in barcodesArr:
        if bc == NEXT_INDICATOR:
            process(productNames, weights, totalWeight)
            totalWeight = 0
        else:
            for pid, pName, wp in zip(bcsInfo[PRODUCT_ID], bcsInfo[PRODUCT_NAME], bcsInfo[WEIGHT_POSITION]):
                if (bc.find(pid, 0, len(pid)) != -1):
                    productNames.append(pName)
                    weight = getWeight(wp, bc)
                    weights.append(weight)
                    totalWeight += int(weight)
    
    if productNames and weights:
        process(productNames, weights, totalWeight)
        parsedBarcodes[PRODUCT_NAME] = productNames
        parsedBarcodes[WEIGHT] = weights

    if parsedBarcodes:
        print(parsedBarcodes)
        writeToExcel()
        saved = True
                
    
def getBarcodesFromInput():
    barcodes = barcodeText.get(1.0, 'end-1c')
    parse(barcodes)


def getBarcodesFromFile():
    filePath = filedialog.askopenfilename(initialdir='/', title='Select File', 
                                          filetypes=(('Text Files', '*.txt'), ('All Files', '*.*')))
    if os.path.isfile(filePath):
        with open(filePath, 'r') as inputFile:
            barcodes = inputFile.read()
            parse(barcodes)


def writeToExcel():
    global outputFile

    if not outputFile:
        outputFile = '.' + outputFileName

    pbDf = pd.DataFrame(parsedBarcodes)
    pbDf.to_excel(outputFile, sheet_name=outputSheetName, index=False)


def selectDirectory():
    global outputFile
    outputDirectory = filedialog.askdirectory()
    outputFile = outputDirectory + outputFileName


parseButton = Button(frame, text ='Parse Barcodes', fg='green', bg='white', command=getBarcodesFromInput) 
parseButton.grid(row=3, pady=(5, 50))

openFileButton = Button(frame, text ='Select A File To Parse', fg='green', bg='white', command=getBarcodesFromFile)
openFileButton.grid(row=4)

selectDirectoryLabel = Label(frame, text ='The parsed file will be saved to same location as the BarcodeParser Application unless you select a location.', font=('Arial', 12), fg='white', bg='green')
selectDirectoryLabel.grid(row=5, pady=(20, 5))

selectDirectoryButton = Button(frame, text ='Select A Location', fg='green', bg='white', command=selectDirectory)
selectDirectoryButton.grid(row=6)

root.mainloop()

# if not saved:
#     writeToExcel(outputFile)