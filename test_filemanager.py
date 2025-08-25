import pytest
from libraries import *
from spreadsheet import Spreadsheet
from file_manager import FileManager
from spreadsheet_types import FileType, ExportType, ImportType
from spreadsheet_errors import *

import os

@pytest.fixture
def initialized_spreadsheet():
    # Initialize a spreadsheet with 10 rows and 10 columns
    spreadsheet = Spreadsheet(rows=10, cols=10)
    return spreadsheet



def test_export_to_file_invalid_extension(initialized_spreadsheet):
    """ Testing invalid extension given to export """
    file_manager = FileManager(initialized_spreadsheet)
    with pytest.raises(FileTypeError):
        file_manager.export_to_file("test.txt", FileType.JSON, ExportType.VALUES_ONLY)


def test_export_to_file_invalid_mode(initialized_spreadsheet):
    """ Testing invalide export type given to export """

    file_manager = FileManager(initialized_spreadsheet)
    with pytest.raises(ExportTypeError):
        file_manager.export_to_file("test.json", FileType.JSON, "invalid_mode")


def test_ex_imp_to_file_json_values_only(initialized_spreadsheet):
    """ Test export and import json values """

    initialized_spreadsheet.edit_cell('a1', '3')
    file_manager = FileManager(initialized_spreadsheet)

    file_manager.export_to_file("test_v.json", FileType.JSON, ExportType.VALUES_ONLY)    
    data = file_manager.__import_from_json__('test_v.json', ImportType.VALUES_ONLY)

    assert isinstance(data, List)
    assert isinstance(data[0], int)
    assert isinstance(data[1], int)
    assert isinstance(data[2], Dict)

    assert data[0] == 10
    assert data[1] == 10
    assert data[2]['A']['1'] == '3'
    

def test_ex_imp_to_file_json_include_info(initialized_spreadsheet):
    """ Test export and import json info """


    initialized_spreadsheet.edit_cell('a1', '=1')
    file_manager = FileManager(initialized_spreadsheet)

    file_manager.export_to_file("test_i.json", FileType.JSON, ExportType.INCLUDE_INFO)
    data = file_manager.__import_from_json__('test_i.json', ImportType.FORMULAS)

    assert isinstance(data, Dict)
    assert isinstance(data['rows'], int)
    assert isinstance(data['cols'], int)
    assert isinstance(data['cells'], List)

    assert data['rows'] == 10
    assert data['cols'] == 10
    assert data['cells'][0]['name'] == 'A1'
    assert data['cells'][0]['index'] == [0,1]
    assert data['cells'][0]['value'] == '1'
    assert data['cells'][0]['formula'] == '=1'
    assert data['cells'][0]['bg color'] == 'white'
    assert data['cells'][0]['fg color'] == 'black'


def test_ex_imp_to_file_excel_include_info(initialized_spreadsheet):
    """ Test export and import excel info"""

    initialized_spreadsheet.edit_cell('a1', '=1')
    file_manager = FileManager(initialized_spreadsheet)

    file_manager.export_to_file("test_i.xlsx", FileType.EXCEL, ExportType.INCLUDE_INFO)
    data = file_manager.__import_from_excel__('test_i.xlsx', ImportType.FORMULAS)

    assert isinstance(data, Dict)
    assert isinstance(data['rows'], int)
    assert isinstance(data['cols'], int)
    assert isinstance(data['cells'], List)

    assert data['rows'] == 10
    assert data['cols'] == 10
    assert data['cells'][0]['name'] == 'A1'
    assert data['cells'][0]['index'] == [0,1]
    assert data['cells'][0]['value'] == '=1'
    assert data['cells'][0]['formula'] == '=1'
    assert data['cells'][0]['bg color'] == 'white'
    assert data['cells'][0]['fg color'] == 'black'



def test_ex_to_file_yaml(initialized_spreadsheet):
    """ Testing exporting to yaml """
    initialized_spreadsheet.edit_cell('a1', '3')
    file_manager = FileManager(initialized_spreadsheet)
    file_manager.export_to_file("test_yaml_values.yml", FileType.YAML, ExportType.VALUES_ONLY)

    file_manager.export_to_file("test_yaml_info.yml", FileType.YAML, ExportType.INCLUDE_INFO)

    assert os.path.exists('test_yaml_values.yml')
    assert os.path.exists('test_yaml_info.yml')


def test_ex_to_file_csv(initialized_spreadsheet):
    """ Testing exporting to csv """

    initialized_spreadsheet.edit_cell('a1', '3')
    file_manager = FileManager(initialized_spreadsheet)
    file_manager.export_to_file("test_yaml_values.csv", FileType.CSV, ExportType.VALUES_ONLY)

    assert os.path.exists('test_yaml_values.csv')
    

def test_ex_to_file_pdf(initialized_spreadsheet):
    """ Testing exporting pdf """

    initialized_spreadsheet.edit_cell('a1', '3')
    file_manager = FileManager(initialized_spreadsheet)
    file_manager.export_to_file("test_yaml_values.pdf", FileType.PDF, ExportType.VALUES_ONLY)
    assert os.path.exists('test_yaml_values.pdf')


def test_import_file_invalid_extension(initialized_spreadsheet):
    """ Testing invlid extension to import  """
    
    file_manager = FileManager(initialized_spreadsheet)
    with pytest.raises(FileTypeError):
        file_manager.import_file("test.txt", FileType.JSON, ImportType.VALUES_ONLY)

def test_import_file_invalid_mode(initialized_spreadsheet):
    """ Testing invlid type of import  """

    file_manager = FileManager(initialized_spreadsheet)
    with pytest.raises(ImportTypeError):
        file_manager.import_file("test.json", FileType.JSON, "invalid_mode")


def test_non_existing_file(initialized_spreadsheet):
    """ Testing import non existing file  """

    file_manager = FileManager(initialized_spreadsheet)

    with pytest.raises(ImportFileError):
        file_manager.import_file('No such file.json', FileType.JSON, ImportType.VALUES_ONLY)
