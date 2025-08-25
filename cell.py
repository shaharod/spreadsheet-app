from libraries import *
from copy import deepcopy

class Cell:
    """
    Class Cell: representation of cell in a spreadsheet. has its own attributes such as name, color, font, etc.
    """

    # Class Constants
    ASCII_A = 65
    EMPTY_CELL = ''

    # Defult values
    DEFULT_BG_COLOR = 'white'
    DEFULT_FG_COLOR = 'black'


    def __init__(self, row: int, col: int, name='', formula=EMPTY_CELL, value=EMPTY_CELL, bg=DEFULT_BG_COLOR, fg=DEFULT_FG_COLOR) -> None:

        # Setting up cell name and location
        self.index: List[int] = [row, col]
        self.name: str = name

        # Setting up background color, text color and font
        self.set_design(bg=bg, fg=fg)

        # Setting values
        self.value = value
        self.formula = formula
        self.dependent_cells: List[str] = []

    def get_info(self) -> Dict[str, str]:
        """
        Function get_info
        Returns a dictionary containing all of the cell's info
        """
        return deepcopy({'name': self.name, 'index': self.index, 'formula': self.formula, 'value': self.value, 'bg color': self.bg, 'fg color': self.fg})


    def get_name(self) -> str:
        """
        Function get_name
        Getter function for cell's class - returns the name of the cell in the form: col_name row_number, for example: A3
        """
        return self.name

    def get_index(self) -> Tuple[int, int]:
        """
        Function get_index
        Getter function for cell's class - returns the index of the cell in the form (row, col)
        """

        return self.index

    def get_formula(self) -> str:
        """
        Function get_formula
        Getter function for cell's class - returns the formula of the cell, for example: '=A1+2'
        """

        return self.formula

    def get_dependent_cell(self) -> List[str]:
        """
        Function get_dependent_cell
        Returns the list of cells names of cells who depend on the current cell's value
        """

        return self.dependent_cells[:]

    def add_dependent_cell(self, cell_name: str):
        """
        Function add_dependent_cell
        
        """
        if cell_name.upper() not in self.dependent_cells:
            self.dependent_cells.append(cell_name.upper())


    def set_formula(self, formula: str) -> None:
        """
        Function set_formula
        Setter function for cell's class - updates the formula of the current cell
        Parameters:
            * formula: str - the new formula of the cell
        """

        self.formula = formula


    def set_value(self, value) -> None:
        """
        Function set_value
        Setter function for cell's class - updates the value of the current cell

        Parameters:
            * value: str - the new value of the cell
        """
        self.value = value

    def set_design(self, bg=DEFULT_BG_COLOR, fg=DEFULT_FG_COLOR) -> None:
        """
        Function set_design
        The function sets cell's design values
        
        Parameters:
            * bg: str - background color of cell
            * fg: str - text color of cell
            * fnt: font - font of the text of the cell (bold or normal)
        
        Return Value: None
        """

        self.bg = bg
        self.fg = fg

        if not self.bg:
            self.bg = Cell.DEFULT_BG_COLOR
        
        if not self.fg:
            self.fg = Cell.DEFULT_BG_COLOR
