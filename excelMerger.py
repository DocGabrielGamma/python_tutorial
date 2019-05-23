#!/usr/bin/python
import sys, getopt
from pathlib import Path
from Readers import excelReader, pathReader

def getDifferentManagersInTable(table):
    managers = ["1113", "5166", "515"]
    return managers

def createExcelTables(managers):
    excelTables = managers
    return excelTables

def extractAllRowsFromTableWithManagerId(table, managerID):
    return "rows"

def appendRowsToTable(table, rows):
    return "sucessfull"

def main(arguments):
    inputPath, outputPath = pathReader.getFilePaths(arguments)
    excelFiles = excelReader.openExcelAt(inputPath)
    print(inputPath)
    print(outputPath)

if __name__ == "__main__":
    # Give me all the arguments if the program is the main
    # Arguments from 1 to size
    main(sys.argv[1:])
