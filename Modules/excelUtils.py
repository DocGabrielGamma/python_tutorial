import Modules.Constants as constants

def findColumnLetterWithValue(columnHeader, sheet):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == columnHeader:
                return cell.column_letter

def getRowWithValue(value, sheet):
    rows = {}
    for row in sheet.iter_rows():
        values = []
        for cell in row:
            values.append(cell.value)
            if str((cell.value)) == value:
                rows[cell.row] = values
    return rows

def addToArrayifUnique(array, value, ignoredValues):
    valueToIntroduce = ""
    if type(value) != str:
        valueToIntroduce = str((value))
    else:
        valueToIntroduce = value
    if valueToIntroduce not in array and valueToIntroduce not in ignoredValues:
        array.append(valueToIntroduce)
    return array

def getDifferentGroupsFile(file, columnHeader):
    sheet = file.active
    groups = []
    columnLetter = findColumnLetterWithValue(columnHeader , sheet)
    for cell in sheet[columnLetter]:
        groups = addToArrayifUnique(groups, cell.value, [columnHeader])
    return groups