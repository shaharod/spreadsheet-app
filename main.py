from libraries import *
import argparse
from spreadsheet import Spreadsheet
from spreadsheet_gui import SpreadsheetGui


class SpreadsheetApp:
    def __init__(self):
        self.spreadsheet = self.get_spreadsheet()

    def handle_args(self) -> Optional[Tuple[int, int]]:
        """
        Function handle_args
        The function creates the command line arguments and help menu, returns the arguments from the user

        Parameters: None

        Return Value: Tuple[int, int] if user chose to decide row and column size, None if user chose to use the defult values.
        """

        app_description = """
        Option 1: main.py -s rows columns
                  Integer values only, No commas. Will open an empty sheet in the mentioned size (rows x columns)
                  If You type illegal values, the program will start with defult size - 50x26
        
        Option 2: main.py
                  Will open an empty sheet with defult size - 50x26

        Notice:     ----  You Cannot Change File Size While Editing  ----


        Once you have decided upon the spreadsheet size, The program will run using graphical user interface.

        Sections:
            * Toolbar - Export File, Import File, Cell Formatting -> All of them have matching buttons
            * Formula Bar - readonly -> Shows the formula for a cell that is focused
            * The Grid - That is the actual spreadsheet where you can edit

        Keyboard and mouse use:
            * Double Click, F2 and Return keys - allow you to edit the current cell you're at, just press one of those keys (or double click on a cell with the mouse)
            * Arrows keys, Tab and Shift-Tab allow you to navigate. They are also what's saving you info in a cell, once youv'e started editing a cell, you can leave it
              with one of the navigation keys and the value will be saved.
    

        Supported Functions and Usage:
            --- All functions and cell references ARE NOT CASE-SENSITVIE, you can use a1 as well as A1 ---
        
        In oreder to type in a formula (function or cell reference), start with '=' sign, and only then type in you formula.
        When you're done editing you will see the formula in the formula bar and the result in the cell.
        If any error occures, both formula and value will change to the error type value

        Regular Functions:
            1) SQRT(arg) - returns squre root of a given numbre (only one)
            2) POW(base, power) - returns base**power
            3) SUM(arg1, arg2, *args) - returns sum of given values, must accepts at least two numbers
            4) AVERAGE(arg1, arg2, *args) - returns sum of given values, must accepts at least two numbers
            
            Note: Numbers can also be cell referenc. A function can be used inside another function: sqrt(sum(90, 10)) - will return 10

        Ranged Functions:
            1) SUM_BY_RANGE(start, stop) - returns sum of ranged cells from starting cell to end.
            2) AVERAGE_BY_RANGE(start, stop) - returns average of ranged cells from starting cell to end.
            3) COUNTNUMS(start, stop) - returns the amount of cells that contain a numeral value in the range
            4) COUNTIF(start, stop, condition) - returns the amount of cells that are equaled by value to the value in the condition

            Must be used only with cell names, any other value will cause an error.



                                                            FILES HANDLING AREA
        Importing Files:
            In the interface there is a button called 'import file' with two options - values and informed.
            This means that you can load a file that has only cell values or a file that contains more info like cell color or formula.

            Supported file type for importing: json and excel only.
        
            Json file format supported for importing:
            Informed import:
            {'rows': int, 'cols': int, 'cells': [{'name': cell_name, 'index': [row, col], 'formula': cell_formula, 'value': cell_value, 'bg color': background_color, 'fg color': foreground_color}, ...]}
            
            Notes
                1) cell name - of the form {Column name}{Row number}, for example: A1, or C3, and also: AB4, AAB10 etc
                2) index - must be integers, length of 2 in the format [row, col] (list of size 2), must match the cell name by values, A1 - index [0,1] (A = 0)
                3) The rest are up to you, if any illegal value exsits an error will be shown.
            
            Values only import:
            {col_name: {row: value, row: value,...}, col_name: {row: value, row: value,...}, ...}
            
            Notes
                1) col_name - Capital letter/s only
                2) row - string numeral value, base is 1. row 0 does not exist
            
                IMPORTANT: All keys must be exactly like in the format, or else the file will not be imported. You will be informed for any error that has occured if any.

        Exporting Files:
            These are the available files for exporting, the menu in the app is clear.
        
            FILE TYPE | VALUES EXPORT |    INFO EXPORT 
            ----------------------------------------------
            JSON      |       X       |        X
            ----------------------------------------------
            YAML      |       X       |        X
            ----------------------------------------------
                      |               | possible, except
            Excel     |       X       | for average_by_range
                      |               | (error in excel only)
            ----------------------------------------------
            CSV       |       X       |
            ----------------------------------------------
            PDF       |       X       |
        """

        parser = argparse.ArgumentParser(usage=app_description)

        # Adding sheet dimention option
        parser.add_argument("-s", "--size", nargs=2, metavar=("Rows", "Columns"), help="Create a new spreadsheet with specified dimensions: rows columns - Must be integers")

        try:
            args = parser.parse_args()
        except argparse.ArgumentError as e:
            print("Please check your commnad line arguments and try again. type main.py --help for help.")

        if args:
            return args.size
        else:
            return None


    def get_spreadsheet(self) -> Spreadsheet:
        """
        Function get_spreadsheet
        The function takes care of the creation of the app's spreadsheet according to the user's choice

        Parameters: None

        Return Value: Spreadsheet object
        """
        # Getting user's choice
        args = self.handle_args()

        # If no arguments are found the spreadsheet will be created using defult values only
        if args:
            try:
                # Parsing arguments and creating spreadsheet based on that
                if isinstance(args, list):
                    rows = int(args[0])
                    cols = int(args[1])

                    spreadsheet = Spreadsheet(rows, cols)

            except Exception as e:
                print(f"{e}\nLoading empty spreadsheet instead")
                spreadsheet = Spreadsheet()
        else:
            spreadsheet = Spreadsheet()

        return spreadsheet


    def run(self):
        self.gui = SpreadsheetGui(self.spreadsheet)
        self.gui.run()


if __name__ == "__main__":
    app = SpreadsheetApp()
    app.run()

