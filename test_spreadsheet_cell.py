import pytest
from spreadsheet import Spreadsheet
from cell import Cell
from spreadsheet_errors import *


def test_cell_creation():
    """
    Testing cell creation with all defult values
    """
    cell = Cell(0, 0)
    assert cell.get_name() == ''
    assert cell.get_index() == [0,0]
    assert cell.get_formula() == ''
    assert cell.get_dependent_cell() == []
    assert cell.value == ''
    assert cell.bg == 'white'
    assert cell.fg == 'black'

def test_set_formula():
    """
    Testing setters functions in cell
    """
    cell = Cell(0, 0)
    cell.set_formula('=A1+B1')
    assert cell.get_formula() == '=A1+B1'

    cell.set_value(5)
    assert cell.value == 5
    
    cell.set_design(bg='red', fg='blue')
    assert cell.bg == 'red'
    assert cell.fg == 'blue'



@pytest.fixture
def empty_spreadsheet():
    return Spreadsheet()

@pytest.fixture
def initialized_spreadsheet():
    spreadsheet = Spreadsheet()

    # Set some initial values in cells
    spreadsheet.edit_cell("A1", "5")
    spreadsheet.edit_cell("A2", "=A1 + 10")
    spreadsheet.edit_cell("B1", "7")
    spreadsheet.edit_cell("B2", "=A2 * B1")
    return spreadsheet


def test_initialization(empty_spreadsheet):
    """
    Testing initialzation works, defult values are 50x26
    """
    assert empty_spreadsheet.rows == 50
    assert empty_spreadsheet.cols == 26


def test_edit_cell(empty_spreadsheet):
    """
    Testing edit_cell function
    * uppercase lowercase sensetivity
    * formulas and values
    * strings - non numeral values
    """
    empty_spreadsheet.edit_cell('A1', '42')
    empty_spreadsheet.edit_cell('a2', '=a1')
    empty_spreadsheet.edit_cell('a3','=5+10')
    empty_spreadsheet.edit_cell('a4','=a1+a2')
    empty_spreadsheet.edit_cell('a5', 'string')

    assert empty_spreadsheet.get_value_from_cell('a1') == '42'
    assert empty_spreadsheet.get_value_from_cell('A2') == '42'
    assert empty_spreadsheet.get_value_from_cell('a3') == '15'
    assert empty_spreadsheet.get_value_from_cell('a4') == '84'
    assert empty_spreadsheet.get_value_from_cell('a5') == 'string'

    assert empty_spreadsheet.get_cell_formula('a1') == '42'
    assert empty_spreadsheet.get_cell_formula('a2') == '=a1'
    assert empty_spreadsheet.get_cell_formula('a3') == '=5+10'
    assert empty_spreadsheet.get_cell_formula('a4') == '=a1+a2'
    assert empty_spreadsheet.get_cell_formula('a5') == 'string'


def test_functions(initialized_spreadsheet):
    """
    Function testing:
    * sum, nesting sum
    * sum_by_range
    * average, nesting average
    * countnums
    * countif
    """
    initialized_spreadsheet.edit_cell('a5', '=sum(a1,a2,b1,3)')
    initialized_spreadsheet.edit_cell('a6', '=sum(a1,a2,b1,sum(a1,a2))')
    initialized_spreadsheet.edit_cell('a7', '=sum_by_range(a1,b2)')
    initialized_spreadsheet.edit_cell('a8', '=average(a1, a2, average(3, 5))')
    initialized_spreadsheet.edit_cell('a9', '=average_by_range(a1,b2)')
    initialized_spreadsheet.edit_cell('a10', '=countnums(a1,b2)')
    initialized_spreadsheet.edit_cell('a11', '=countif(a1,b2,a1)')

    assert initialized_spreadsheet.get_value_from_cell('a5') == '30'
    assert initialized_spreadsheet.get_value_from_cell('a6') == '47'
    assert initialized_spreadsheet.get_value_from_cell('a7') == '132'
    assert initialized_spreadsheet.get_value_from_cell('a8') == '8.0'
    assert initialized_spreadsheet.get_value_from_cell('a9') == '33.0'
    assert initialized_spreadsheet.get_value_from_cell('a10') == '4'
    assert initialized_spreadsheet.get_value_from_cell('a11') == '1'
    


def test_evaluate_cell_value(initialized_spreadsheet: Spreadsheet):
    """
    Testin evalutaing a cell is working
    """
    assert initialized_spreadsheet.evaluate_cell_value("5", "A1") == '5'
    assert initialized_spreadsheet.evaluate_cell_value("=A1 + 10", "A2") == '15'
    assert initialized_spreadsheet.evaluate_cell_value("=A2 * B1", "B2") == '105'


def test_edit_cell_with_formula(initialized_spreadsheet: Spreadsheet):
    """
    Testing dependencies between cells
    """
    initialized_spreadsheet.edit_cell("A1", "10")
    assert initialized_spreadsheet.get_value_from_cell("A2") == "20"
    assert initialized_spreadsheet.get_value_from_cell("B2") == "140"


def test_edit_cell_circular_reference(initialized_spreadsheet: Spreadsheet):
    """
    Testing error circular reference
    """
    with pytest.raises(CircularReferenceError):
        initialized_spreadsheet.edit_cell("A1", "=A2")


def test_edit_cell_invalid_formula(initialized_spreadsheet: Spreadsheet):
    """
    Testing error invlid formula
    """
    with pytest.raises(FormulaValueError):
        initialized_spreadsheet.edit_cell("A1", "=4 +")


def test_edit_cell_non_existing_function(initialized_spreadsheet: Spreadsheet):
    """
    Testing no such function exsist
    """
    with pytest.raises(NonExistingFunctionError):
        initialized_spreadsheet.edit_cell("A1", "=NONEXISTENT(a3)")


def test_edit_cell_zero_division(initialized_spreadsheet: Spreadsheet):
    """
    Testing zero division error, both number and a cell values 0
    """
    initialized_spreadsheet.edit_cell("A1", "0")

    with pytest.raises(ZeroDivision):
        initialized_spreadsheet.edit_cell("B1", "=5/A1")

    with pytest.raises(ZeroDivision):
        initialized_spreadsheet.edit_cell("B1", "=5/0")



import string
def test_int_to_letter_conversion(empty_spreadsheet: Spreadsheet):
    """
    Testing the names of the cells are formatted correctly
    """
    # Generate cell names for single-letter columns (A to Z)
    single_letter_names = list(string.ascii_uppercase)

    # Generate cell names for two-letter columns (AA to ZZ)
    two_letter_names = [f"{first}{second}" for first in string.ascii_uppercase for second in string.ascii_uppercase]

    # Generate cell names for three-letter columns (AAA to AZZ)
    three_letter_names = [f"{first}{second}{third}" for first in string.ascii_uppercase for second in string.ascii_uppercase for third in string.ascii_uppercase]

    # Combine all cell names
    expected_letters = single_letter_names + two_letter_names + three_letter_names

    
    # Iterate through expected letters and compare with actual conversion
    for i, expected_letter in enumerate(expected_letters, start=1):
        actual_letter = empty_spreadsheet.__int_to_letter__(i)
        assert actual_letter == expected_letter, f"Failed for index {i}: Expected {expected_letter}, but got {actual_letter}"