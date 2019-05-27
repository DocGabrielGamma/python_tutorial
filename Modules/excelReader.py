from os import listdir
from os.path import isfile, join
from openpyxl import load_workbook
import Modules.Constants as constants

def isExcelFileName(completePath):
    if isfile(completePath) and completePath.endswith('.xlsx'):
        return True
    return False

def constructFileNamesList(excelFiles, file, path):
    completePath = join(path, file)
    if isExcelFileName(completePath):
        excelFiles.append(completePath)
    return excelFiles

def getExcelFileNamesinPath(path):
    excelFileNames = []
    files = listdir(path)
    if len(files) > 0: 
        for file in files:
            excelFileNames = constructFileNamesList(excelFileNames, file, path)
        return excelFileNames
    return []

def constructFileList(excelFilenames):
    excelFiles = []
    for filename in excelFilenames:
        excelFiles.append(load_workbook(filename))
    return excelFiles

def openExcelAt(path):
    excelFilenames = getExcelFileNamesinPath(path)
    return constructFileList(excelFilenames)