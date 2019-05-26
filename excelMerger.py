#!/usr/bin/python
import sys, getopt
from pathlib import Path
from openpyxl import Workbook, load_workbook
from Modules import excelReader, pathReader, groupReader

def createExcelInPath(path):
    wb = Workbook()
    wb.save(path)
    return wb

def createFileIfRequired(name, outputPath):
    excelName = name + ".xlsx"
    excelFilesNames = excelReader.getExcelFileNamesinPath(outputPath)
    finalPath = outputPath + '\\' + excelName
    if len(excelFilesNames) < 1:
        return createExcelInPath(finalPath), finalPath
    for excelFileName in excelFilesNames:
        if excelFileName.find(excelName) > -1:
           return load_workbook(finalPath), finalPath
    return createExcelInPath(finalPath), finalPath
    
def createFileByGroups(groups, file, outputPath):
    for group in groups:
      excelFile, path = createFileIfRequired(group, outputPath)
      rows = extractAllRowsWithGroup(file, group)
      appendRowsToTable(file, excelFile, path, rows)
    return "lastRow"

def createExcelTables(managers):
    excelTables = managers
    return excelTables

def extractAllRowsWithGroup(file, group):
    return groupReader.findRowWithValue(group, file.active)

def appendRowsToTable(inputFile, outputFile, path, rows, offsetRow = 0, offsetColumn = 1):
    sheet = outputFile.active
    initialRow = sheet.max_row
    for rowIndex, keys in enumerate(rows):
        for columnIndex, value in enumerate(rows[keys]):
            rowCoordinate = rowIndex + initialRow + offsetRow
            columnCoordinate = columnIndex + offsetColumn
            sheet.cell(row=rowCoordinate, column=columnCoordinate, value=value)
    outputFile.save(path)
    #printTables(inputFile, outputFile, path, rows)
    return "Success"

def createExcelFileForGroup(inputFiles, outputPath, columnHeader):
    groups = []
    for file in inputFiles:
        groups = groupReader.getDifferentGroupsFile(file, columnHeader)
        createFileByGroups(groups, file, outputPath)
    return groups

def main(arguments):
    inputPath, outputPath, columnHeader = pathReader.getFilePaths(arguments)
    inputExcelFiles = excelReader.openExcelAt(inputPath)
    createExcelFileForGroup(inputExcelFiles, outputPath, columnHeader)
    print(inputPath)
    print(outputPath)

if __name__ == "__main__":
    # Give me all the arguments if the program is the main
    # Arguments from 1 to size
    main(sys.argv[1:])
