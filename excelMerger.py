#!/usr/bin/python
import sys, getopt
from pathlib import Path
from openpyxl import Workbook, load_workbook
from Modules.excelReader import openExcelAt
from Modules.excelUtils import getDifferentGroupsFile, getRowWithValue
from Modules.excelWriter import createFileByGroups, appendRowsToFile
from Modules import excelReader, excelWriter, pathReader
import Modules.Constants as constants

def extractAllRowsWithGroup(file, excelFilesByGroups):
    for group in excelFilesByGroups:
        excelFilesByGroups[group][constants.ROWS_KEY] = getRowWithValue(group, file.active)
    return excelFilesByGroups

def appendRowsToEachGroupFile(file, excelFilesByGroups):
    for group in excelFilesByGroups:
      rows = excelFilesByGroups[group][constants.ROWS_KEY]
      file = excelFilesByGroups[group][constants.WORKBOOK_KEY]
      path = excelFilesByGroups[group][constants.PATH_KEY]
      appendRowsToFile(file, path, rows)
    return "Success"

def createExcelFileForGroup(inputFiles, outputPath, columnHeader):
    for file in inputFiles:
        groups = getDifferentGroupsFile(file, columnHeader)
        excelFilesByGroups = createFileByGroups(groups, file, outputPath)
        excelFilesByGroups = extractAllRowsWithGroup(file, excelFilesByGroups)
        appendRowsToEachGroupFile(file, excelFilesByGroups)
    return "Success"

def main(arguments):
    inputPath, outputPath, columnHeader = pathReader.getFilePaths(arguments)
    inputExcelFiles = excelReader.openExcelAt(inputPath)
    createExcelFileForGroup(inputExcelFiles, outputPath, columnHeader)
    print("Success")

if __name__ == "__main__":
    # Give me all the arguments if the program is the main
    # Arguments from 1 to size
    main(sys.argv[1:])