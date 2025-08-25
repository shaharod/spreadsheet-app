
# Classes for spreadsheet errors
class SpreadsheetError(Exception):
    """ Generic error in spreadsheet """
    VALUE = '#ERROR#'

class CircularReferenceError(SpreadsheetError):
    """ Circular reference from a cell to itself """
    VALUE ='#CIRCULAR#'

class NonExistingFunctionError(SpreadsheetError):
    """ Unknown function in use """
    VALUE = '#NoSuchFunction#'

class CellLocationError(SpreadsheetError):
    """ Cell does not exsist """
    VALUE = '#NoCell#'

class ZeroDivision(SpreadsheetError):
    """ Division by 0 """
    VALUE = '#ZeroDiv#'

class FormulaValueError(SpreadsheetError):
    """ Formula value error """
    VALUE = '#FuncValue#'

class FunctionSyntaxError(SpreadsheetError):
    """ Function syntax being used worng """
    VALUE = '#FuncSyntax#'


# Classes for files handling errors
class FileError(Exception):
    """ Generic file error """
    pass

class ExportTypeError(FileError):
    """ Exporting wrong file type """
    pass

class FileTypeError(FileError):
    """ Using wrong file type """
    pass

class FileFormatError(FileError):
    """ File not formatted coerrectly """
    pass

class ImportFileError(FileError):
    """ Error in importing a file """
    pass

class ImportTypeError(FileError):
    """ Importing wrong file type """
    pass

class EmptyFileImportError(FileError):
    """ Importing an empty file """
    pass