from enum import Enum


# Types for files
class FileType(Enum):
    JSON = '.json'
    YAML = '.yml'
    CSV = '.csv'
    EXCEL = '.xlsx'
    PDF = '.pdf'

# Types for exporting
class ExportType(Enum):
    INCLUDE_INFO = 'i'
    VALUES_ONLY = 'v'

# Types for importing
class ImportType(Enum):
    FORMULAS = 'f'
    VALUES_ONLY = 'v'
