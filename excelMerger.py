#!/usr/bin/python
import sys, getopt
from pathlib import Path
from openpyxl import Workbook, load_workbook
from Modules.excelReader import openExcelAt, extractAllRowsWithGroup
from Modules.excelWriter import createFileByGroups, appendRowsToTable
from Modules import excelReader, excelWriter, pathReader
import Modules.Constants as constants

def createExcelFileForGroup(inputFiles, outputPath, columnHeader):
    for file in inputFiles:
        groups = excelReader.getDifferentGroupsFile(file, columnHeader)
        excelFilesByGroups = createFileByGroups(groups, file, outputPath)
        excelFilesByGroups = extractAllRowsWithGroup(file, excelFilesByGroups)
        appendRowsToTable(file, excelFilesByGroups)
    return groups

def main(arguments):
    inputPath, outputPath, columnHeader = pathReader.getFilePaths(arguments)
    inputExcelFiles = excelReader.openExcelAt(inputPath)
    createExcelFileForGroup(inputExcelFiles, outputPath, columnHeader)
    print("Success")

if __name__ == "__main__":
    # Give me all the arguments if the program is the main
    # Arguments from 1 to size
    main(sys.argv[1:])