from libraries import *
from spreadsheet_errors import *
from cell import Cell
from re import findall
from math import sqrt, pow


class Spreadsheet:

    # Class's constants
    __FORMULA_PREFIX = '='
    __CELL_NAME_RGULAR_EXPRESSION = r'[A-Za-z]+[1-9]\d*'

    def __init__(self, rows=50, cols=26) -> None:
        
        # Size of the spreadsheet
        self.rows = rows
        self.cols = cols

        # Dataframe representation of the spreadsheet
        self.cells_df = pd.DataFrame(index=[str(i + 1) for i in range(rows)], columns=[self.__int_to_letter__(i) for i in range(1, self.cols+1)])
        self.cells_df.index.name = 'Row'
        self.cells_df.columns.name = 'Column'

        # Making sure no Nan values are in initial dataframe
        self.cells_df.fillna('', inplace=True)

        # Cells objects list
        self.cells_objs: List[List[Cell]] = [[Cell(row, col, name=(self.__int_to_letter__(col)+str(row+1))) for col in range(1, self.cols+1)] for row in range(rows)]

        # Supported functions
        self.SUPPORTED_FUNCTIONS = {"SUM": self.__sum__, "AVERAGE": self.__mean__, "SQRT": self.__sqrt__, "POWER": self.__pow__}
        self.RANGE_FUNCTIONS = {"SUM_BY_RANGE": self.__sum_by_range__, "AVERAGE_BY_RANGE": self.__mean_by_range__, "COUNTNUMS": self.__countnums__, "COUNTIF": self.__count_if__}


        self.original_vals = {'rows': rows, 'cols': cols, 'df': self.cells_df, 'cells': self.cells_objs[:]}


    def get_cells_info(self, cell_id: str) -> Dict:
        """
        Function get_cell_info
        The function returns the wanted cell's information dictionary

        Parameters: None

        Return value: Dict
        """
        # Getting cell's object
        cell = self. __get_cell_by_name__(cell_id)

        return cell.get_info()


    def __int_to_letter__(self, n: int) -> str:
        """
        Function __int_to_letter__
        The function converts integer value of column into a capital literal value, for example - 0 turns into A.

        Parameters:
        * n: int - The integer value to convert

        Return Value: str - the converted string value of the column name
        """

        # Constants values
        CAPITAL_A_ASCII = 65
        ALPHBET_LETTERS = 26

        result = ''

        # Run until n is 0
        while n:

            # Calculate quotient and remainder using divmod, subtract 1 from n to handle 0-indexing
            n, remainder = divmod(n-1, ALPHBET_LETTERS)

            # Convert remainder to corresponding capital letter and prepend to result
            result = chr(CAPITAL_A_ASCII + remainder) + result
        
        return result


    def get_value_from_cell(self, cell_id: str) -> Union[str, int]:
        """
        Function get_value_from_cell
        The function returns the value of a given cell by its id

        Parameters:
            * cell_id: str - string represented by a capital letter and number, for example A4
        
        Return Value: string or int, depends on cell's value stored in the class's dataframe
        """

        # Getting index from cell id
        row, col = self.__cell_name_tuple__(cell_id)
        
        return self.cells_df.at[row, col]


    def __cell_name_tuple__(self, cell_id: str) -> Tuple[str, str]:
        """
        Function __cell_name_tuple__
        The function seperates row and column value from a given cell id and returns a tuple with those values

        Parameters:
            * cell_id: str
        
        Return Value: Tuple containing row and column values as strings (row, col)
        """

        # Getting index for seperation
        for i, char in enumerate(cell_id):
            if char.isdigit():
                break

        # Seperating row and column value
        col = cell_id[:i]
        row = cell_id[i:]

        return row, col.upper()


    def get_cell_formula(self, cell_name) -> str:
        """
        Function get_cell_formula
        The function returns a given cell's formula using it's id

        Parameters:
            * cell_name: str
        
        Return Value: str - the formula of the cell, usually starts with '=', unless user just put a value and then its just the value
        """
        
        # Getting cell by it's id
        cell = self.__get_cell_by_name__(cell_name)

        return cell.get_formula()


    def __update_cell_dependencies__(self, cell: Cell) -> None:
        """
        Function __update_cell_dependencies__
        The function recives a cell and gets its dependent cells - the ones that depend on its value, and updates every one of them with the new value recursively.
        In case of an error, it changes the value to the corresponding error value.

        Parameters:
            * cell: Cell - an actual cell object
        
        Return value: None

        Exceptions: Function raises an exception of the type SpreadsheetError. It raises the specific error it cateches.
        """

        # Getting cells that depend on current cell
        deps = cell.get_dependent_cell()

        # Going over the tokens of the cells
        for token in deps:

            # Gettingn actual cell object for each token
            curr_cell = self.__get_cell_by_name__(token)
            row, col = self.__cell_name_tuple__(token)

            try:
                # Getting cell new value after the original cell has changed
                value = self.evaluate_cell_value(curr_cell.formula, token)
                curr_cell.set_value(value)

                # Recursively updating the current cell's dependent cells
                self.__update_cell_dependencies__(curr_cell)

                # Updating the dataframe
                self.cells_df.at[row, col] = value

            # Taking care of an error if any has occured
            except SpreadsheetError as error:

                # Getting corresponding error's value
                error_val =  error.VALUE   #Spreadsheet.__ERRORS_INFO[type(error)]

                # Updating cell's value with error value
                curr_cell.set_value(error_val)
                row, col = self.__cell_name_tuple__(token)
                self.cells_df.at[row, col] = error_val

                raise error


    def edit_cell(self, cell_name: str, new_val: str) -> None:
        """
        Function edit_cell
        The function edits the cell's value while making sure it's possible. The function takes care of errors if any occur.

        Parameters:
            * cell_name: str
            * new_val: str
        
        Return Value: None
        """
        # Gettin current cell and updating its formula to the new value
        cell = self.__get_cell_by_name__(cell_name)
        cell.set_formula(new_val)
        row, col = self.__cell_name_tuple__(cell_name)
        
        final_value = SpreadsheetError.VALUE

        try:
            # Evalueating new cell value based on the formula if any
            final_value = self.evaluate_cell_value(new_val, cell_name)

            # If no errors occured, actually changing the value and updating dependencies
            cell.set_value(final_value)
            self.__update_cell_dependencies__(cell)

            # Updating grid
            self.cells_df.at[row, col] = str(final_value)

        # Taking care of errors that may occur
        except SpreadsheetError as error:

            # Getting error value and updating cell's value and formula accordingly
            error_value = error.VALUE #Spreadsheet.__ERRORS_INFO[type(error)]
            cell.set_formula(error_value)
            cell.set_value(error_value)

            # Updating dataframe as well
            self.cells_df.at[row, col] = error_value
            raise error


    def evaluate_cell_value(self, expression: str, original_cell_name: str) -> Optional[str]:
        """
        Function evaluate_cell_value
        The function recursovley going through cells and their dependencies and reaching the final value. If any errors occur, the function raises a matching SpreadSheetError.

        Parameters:
            * expression: str - may be a fomula or just a regular value
            * original_cell_name: str - the cell to store the new value in
        
        Return Value: Optional[str] - The final value of the cell, might raise an error if any occures.
        """

        # Base-case checks
        if expression is None or expression == Cell.EMPTY_CELL:
            return expression

        # If expression is not a formula there is nothing to evaluate - another type of base case
        if not self.is_formula(expression):
            return expression

        # Uppering all of the values in order to evaluate porperly
        original_cell_name = original_cell_name.upper()
        expression = expression.upper()

        # Getting rid of any spaces and getting rid of '=' sign
        expression = expression[1:].replace(' ', '')

        # Calculating any supported function if there are any, updating the formula accordingly
        formula = self.__calculate_ranged_funcs__(expression, original_cell_name)


        # Getting all of the tokens of cell names that appear in the formula in order to evaluate their values
        tokens: List[str] = findall(Spreadsheet.__CELL_NAME_RGULAR_EXPRESSION, formula)
        
        # Checking for circular reference, raising matching error if found
        if original_cell_name in tokens:
            raise CircularReferenceError(f"Circular Reference at cell: {original_cell_name}")
        
        # Calculate each value for every token in the formula if any
        if tokens:
            values = {}
            for token in tokens:

                # Getting actual cell object by its name
                cell = self.__get_cell_by_name__(token)

                # Recursively creating a dictionary of cells names and thier values, updating cell's dependent cells
                values[token] = self.evaluate_cell_value(cell.formula, original_cell_name)
                cell.add_dependent_cell(original_cell_name)

            # Replacing empty cells with 0 for calcuation purpuses
            values = {key: value if value != Cell.EMPTY_CELL else '0' for key, value in values.items()}

            # Covering a case where the user inserts only one cell and that cell contains a string value
            if len(tokens) == 1:
                string_to_check = formula.replace(tokens[0], "").strip()
                if not string_to_check:
                    return values[tokens[0].upper()]

            # Changin values in the dictionary to numeral values if possible, raising an error if not
            if self.__str_values_to_numeric__(values):

                # Adding supported functions to the values dictionary
                values.update(self.SUPPORTED_FUNCTIONS)

                try:
                   # Evaluating mathematical value based on formulas an actions
                   return_value = str(eval(formula, values))
                
                # Taking care of errors that may occure
                except NameError:
                    raise NonExistingFunctionError("Unknown Cell or Function being used")
                except ZeroDivisionError:
                    raise ZeroDivision('Zero Division Error')
                except:
                    raise FunctionSyntaxError("Invalid syntax for function")
            else:
                raise FormulaValueError(f"Invalid literal for int() or float() in cell: {original_cell_name}")

        else:
            # In case there arent any tokens, evaluating value for mathematical value
            try:
                return_value = str(eval(formula, self.SUPPORTED_FUNCTIONS))
            
            # Taking care of errors that may occure
            except ZeroDivisionError:
                raise ZeroDivision("Cannot divide by 0")
            except:
                raise FormulaValueError("Forula format is invalid. Can only contain cell names, numbers and supported functions. Formula must contain at least two values!")

        return return_value


    def is_formula(self, expression: str) -> bool:
        """
        Function _is_formula
        An auxiliary function that checks if a given expression is a fomula (meaning its value starts with '=' and contains other values)

        Parameters:
            * expression: str
        
        Return Value: bool - True if expression is a fomula, False otherwise
        """
        if expression:
            return expression[0] == Spreadsheet.__FORMULA_PREFIX and len(expression) > 1
        
        return False


    def __get_cell_by_name__(self, cell_name: str) -> Cell:
        """
        Function __get_cell_by_name__
        An auxiliary function that returns a Cell object based on its id.

        Parameters:
            * cell_name: str
        
        Return Value: Cell object
        """

        # Speprating row and col values from the id
        row, col = self.__cell_name_tuple__(cell_name)

        cols = self.cells_df.columns.to_list()

        # Setting row index
        row_index = int(row) - 1

        found = False

        # Finding column index based on its name
        for col_index, column in enumerate(cols):
            if column == col:
                found = True
                break
        
        # Validating rows and cols
        if row_index < 0 or row_index > len(self.cells_objs) - 1 or not found:
            raise CellLocationError("Cell Does not exist")

        return self.cells_objs[row_index][col_index]


    def __str_values_to_numeric__(self, values_dict: Dict[str,str]) -> bool:
        """
        Function __str_values_to_numeric__
        An auxiliary function that converts numeral strings into their matching numeral type - int or float. Its done for calculation purpuses using eval() function.

        Parameters:
            *  values_dict: Dict[str, str] - dictionary of cell names and values.
        
        Return Value: bool - True if conversion has succeded, False otherwise
        """

        # Going over the values
        for cell in values_dict:

            value = values_dict[cell]

            if isinstance(value, str):

                # Trying conversion, will succeed if value is numeral
                try:

                    if value != '':

                        # Checking both float and int
                        if '.' in value:
                            values_dict[cell] = float(value)
                        else:
                            values_dict[cell] = int(value)
                except:
                    if value in '+-/*':
                        continue
                    else:
                        # Returning False for any error that may have occured - it means the value is not numeral convertable
                        return False
        return True




    #######                      SUPPORTED FUNCTIONS SECTION                    #######

    def __calculate_ranged_funcs__(self, formula: str, target_cell: str) -> str:
        """
        Function __calculate_ranged_funcs__
        An auxiliary function that calls the right range function that the user has asked for. It returns a new formula without the functions in it, only values.

        Parameters:
            * formula: str
        
        Return Value: str - the formula without the functions in it, values only
        """

        # Ranges cell's indexes in the matched tuple
        START_CELL = 0
        STOP_CELL = 1
        CONDITION = 2

        # Extracting needed functions from supported range functions dictionary
        range_funcs = [func for func in self.RANGE_FUNCTIONS.keys() if func in formula]

        # Going over every found function
        for func in range_funcs:

            if func == 'COUNTIF':
                pattern = rf'COUNTIF\(({self.__CELL_NAME_RGULAR_EXPRESSION}),\s*({self.__CELL_NAME_RGULAR_EXPRESSION}),\s*(.*)\)'
            else:

                # Creating pattern to look for per function and getting all matches
                pattern = rf'{func}\(({self.__CELL_NAME_RGULAR_EXPRESSION}),\s*({self.__CELL_NAME_RGULAR_EXPRESSION})\)'
        

            matches = findall(pattern, formula)

            # Taking care of errors
            if not matches:
                raise FormulaValueError("Function not being called correctly")
            
            # Getting function to call
            call_func = self.RANGE_FUNCTIONS[func]


            # For each match, calling function and calculating value
            for attrs in matches:
                
                if call_func == self.RANGE_FUNCTIONS['COUNTIF']:
                    # New df with range
                    ranged_df: pd.DataFrame = self.__get_df_by_range__(attrs[START_CELL], attrs[STOP_CELL], target_cell, count_func=True)
                    new_val = self.__count_if__(ranged_df, attrs[CONDITION])
                    func_replace = func + '(' + str(attrs[0]) + ',' + str(attrs[1]) + ',' + str(attrs[2]) + ')'
                else:
                    count_func = call_func == self.RANGE_FUNCTIONS['COUNTNUMS']
                    # New df with range
                    ranged_df: pd.DataFrame = self.__get_df_by_range__(attrs[START_CELL], attrs[STOP_CELL], target_cell, count_func=count_func)
                    new_val = call_func(ranged_df)

                    # Creating the function again in order to remove it from the string
                    func_replace = func + '(' + str(attrs[0]) + ',' + str(attrs[1]) +')'

                # Updating cells dependencies
                rows = ranged_df.index.to_list()
                cols = ranged_df.columns.to_list()

                for row in rows:
                    for col in cols:
                        cell_id = col+row
                        cell = self.__get_cell_by_name__(cell_id)
                        self.__update_cell_dependencies__(cell)

                
                # Replacing function with actual value
                formula = formula.replace(func_replace, str(new_val))

        return formula


    def __sqrt__(self, arg: Union[int, float]) -> float:
        """
        Function __sqrt__
        Calculates the square root of a given number

        Parameters:
            * arg: Union[int, float] - the nunber to calculate from
        Return Value: Union[int, float] - the square root of the number
        """
        return sqrt(arg)


    def __pow__(self, base: Union[int, float], power: Union[int, float]) -> Union[int, float]:
        """
        Function __pow__
        Calculates base ** power - 'base' to the power of 'power'

        Parameters:
            * arg1, arg2: Union[int, float] - numeral values, mandatory.

        Return Value: Union[int, float] - the result of the power
        """
        return pow(base, power)
    

    def __sum__(self, arg1: Union[int, float], arg2: Union[int, float], *args: Union[int, float]) -> Union[int, float]:
        """
        Function __sum__
        The function accepts minimum of two arguments to sum, and can get as many arguments byond that. Calculates the sum of the given values

        Parameters:
            * arg1, arg2: Union[int, float] - numeral values, mandatory.
            * *args: Union[int, float] - additional values to calculate
        
        Return Value: Union[int, float] - the final sum of all of the values
        """

        # Getting base
        result = arg1+arg2

        # Adding the rest of the numbers
        for num in args:
            result += num
        
        return result


    def __mean__(self, arg1: str, arg2: str, *args: str) -> float:
        """
        Function __mean__
        The function accepts minimum of two arguments to sum, and can get as many arguments byond that. Calculates the average of the given values

        Parameters:
            * arg1, arg2: Union[int, float] - numeral values, mandatory.
            * *args: Union[int, float] - additional values to calculate

        Return Value: float - the calculated average
        """

        # Getting sum and divisor (number of arguments in total)
        args__sum__ = self.__sum__(arg1, arg2, *args)
        division = 2 + len(args)

        # Calculating average
        return args__sum__ / division


    def __sum_by_range__(self, range_df: pd.DataFrame) -> Union[int, float]:
        """
        Function __sum_by_range__
        The Function gets two parameters only: start and stop. ignores any cells with errors. Calculates the sum of the cells in the given range

        Parameters:
            * range_df: DataFrame - The cut dataframe to calculate
        
        Return Value: Union[int, float] - the final sum
        """

        # Operating pandas sum function on given range
        return range_df.sum().sum()
    

    def __mean_by_range__(self, range_df: pd.DataFrame) -> float:
        """
        Function __mean_by_range__
        The Function gets two parameters only: start and stop. ignores any cells with errors. Calculates the avereage of the cells in the given range

        Parameters:
            * range_df: DataFrame - The cut dataframe to calculate
        
        Return Value: float - the final average
        """

        # Operating pandas mean function on given range
        return range_df.mean().mean()


    def __countnums__(self, range_df: pd.DataFrame) -> int:
        """
        Function __countnums__
        Function counts every numeric value in a given range of cells

        Parameters:
            * range_df: pd.DataFrame - The cut df range
        
        Return Value: int - The total number of numeral values in the range of cells
        """

        # Getting rows
        rows = range_df.index.to_list()
        cols = range_df.columns.to_list()

        count = 0

        # Counting numbers amount
        for row in rows:
            for col in cols:
                if range_df.at[row, col].isdigit():
                    count += 1

        return count


    def __count_if__(self, range_df: pd.DataFrame, condition: str) -> int:
        """
        Function __count_if__
        The function counts every cell in the given range that is equal to the condition

        Parameters:
            * range_df: pd.DataFrame - The cut df
            * condition: str - The value to compare all of the values with
        """

        rows = range_df.index.to_list()
        cols = range_df.columns.to_list()

        count = 0

        cell_token = findall(self.__CELL_NAME_RGULAR_EXPRESSION, condition)

        if len(cell_token) > 1:
            raise FunctionSyntaxError('Wrong useage in COUNTIF function. Correct one: COUNTIF(start, stop, condition), condition can be only one cell name or a value')
        elif cell_token:
            condition = self.get_value_from_cell(cell_token[0])

        # Counting based on condition
        for row in rows:
            for col in cols:
                if range_df.at[row, col] == condition:
                    count += 1
        
        return count


    def __get_df_by_range__(self, start: str, stop: str, target_cell: str, count_func=False) -> pd.DataFrame:
        """
        Function __get_df_by_range__
        An auxiliary function that cuts the existing df and returns only the needed part from it.

        Parameters:
            * start: str - starting cell id
            * stop: str - end of range cell id
        
        Return Value: DataFrame - the new cut
        """

        # Getting row and cols sperated from start and stop cells ids
        start_r, start_c = self.__cell_name_tuple__(start)
        stop_r, stop_c = self.__cell_name_tuple__(stop)

        # Getting needed df part
        range_to_sum = self.cells_df.loc[start_r:stop_r, start_c:stop_c]

        # Maintaining old version of pandas where pandas infer types
        range_to_sum = range_to_sum.replace(Cell.EMPTY_CELL, '0').infer_objects()

        rows = range_to_sum.index.to_list()
        cols = range_to_sum.columns.to_list()

        # Covering circular references
        for row in rows:
            for col in cols:
                if col+row == target_cell:
                    raise CircularReferenceError("Circular Reference at cell: target_cell")


        try:
            # Making sure all of the values in the cut df are numeric or raising matching SpreadsheetError
            if not count_func:
                range_to_sum = range_to_sum.apply(pd.to_numeric)
        except:
            raise FormulaValueError("Cannot perform mathematical operation on non numeral types. Check your cells in the wanted range.")
        
        return range_to_sum


    def reset_sheet(self, rows: int, cols: int) -> None:
        """
        Function reset_sheet
        The function restes the whole spreadsheet to a new size with empty cells, in order to set the spreadsheet for a new import of data

        Parameters:
            * rows, cols: int
        
        Return Value: None
        """

        # Temp prev vars
        self.prev_rows = self.rows
        self.prev_cols = self.cols
        self.prev_df = self.cells_df
        self.prev_cells_objs = self.cells_objs[:]


        # Size of the spreadsheet
        self.rows = rows
        self.cols = cols

        # Dataframe representation of the spreadsheet
        self.cells_df = pd.DataFrame(index=[str(i + 1) for i in range(rows)], columns=[self.__int_to_letter__(i) for i in range(1, self.cols+1)])
        self.cells_df.index.name = 'Row'
        self.cells_df.columns.name = 'Column'

        # Making sure no Nan values are in initial dataframe
        self.cells_df.fillna('', inplace=True)

        # Cells objects list
        self.cells_objs: List[List[Cell]] = [[Cell(row, col, name=(self.__int_to_letter__(col)+str(row+1))) for col in range(1, self.cols+1)] for row in range(rows)]


    def undo_reset(self) -> None:
        """
        Function undo_reset
        The function is caleed before the spreadsheet is being reset and saves the current values. Then, if its called with the activate flag, it will set
        the values of the spreadsheet to the values it had before the change.

        Parameter:
            * activate: bool, by defult False. The flag that marks to the function to undo the reset.
        
        Return Value: None
        """

        self.rows = self.prev_rows
        self.cols = self.prev_cols
        self.cells_df = self.prev_df
        self.cells_objs = self.prev_cells_objs[:]
