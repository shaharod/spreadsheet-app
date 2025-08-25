# Spreadsheet Python App

**By:** Shahar Oded - Koren

---

## Project Extensions

1. GUI
2. Exporting and importing to and from files, including formulas and colors - Using buttons in the interface
3. Text formatting - Foreground and Background for each cell - Using buttons in the interface
4. Testing using pytest - Tried to cover everything that can happen.
5. Mathematical functions - SQRT, POWER, COUNTIF, COUNTNUMS (including by cell names)

**Usage of Functions:**
SQRT(<value>) - value may be cell name or just a numeral value
POWER(<base>, <power>) - again, both values can be cell names or numeral
COUNTNUMS(<start>,<stop>) - same specifications as before
COUNTIF(<start>, <stop>, <condition>) - condition must be one cell name or any other expression. cannot calculate cell names

---

## Explanation for Extensions

### 1. GUI

The GUI extension lets the user interact with the spreadsheet in a convinet way.
The gui is designd so the user could navigate between cells using the keyboard for a more efficiect workflow.

It contains a toolbar at the top with buttons for possible actions on the file.

Also, it contains a formula bar so that the formulas are presented to the user when typing. The formula bar is readonly and cannot be edited.

In order for a text to be counted as formula, use '=' sign before writing the formula.

The cells can be edited by pressing 'Enter' (Return) key in the keyboard or by F2. They can be also be edited using mouse double click on a given cell.
In order to change cell's color just press on the cell and then press the color button - background or foreground.
Navigation between cells is dont by mouse click or by keyboard arrows and Tab (Shift-Tab to go backwards).

---

### 2. Importing and Exporting Files

**Importing**
My program supports importing from json files and excel files. colors are not supported for excel type import, but for json only.
Excel Import Warnings: Not all excel functions are supported so information may not load as expected.

If json file is chosen, the import can be done in two ways from which the user choses:

* Informed: That means the file of json also stores information about the cells, like colors or formulas. In that case, it is possible to import a json file
  if it is formatted correctly: {'rows': int, 'cols': int, 'cells': \[List of dictionries with information for each cell]}
  Detailed format:
  {'rows': int, 'cols': int, 'cells': \[{'name': cell\_name, 'index': \[row, col], 'formula': cell\_formula, 'value': cell\_value, 'bg color': background\_color, 'fg color': foreground\_color}, ...]}

* Values only: That means no additional info is being imported, only the values for each cell.
  The format is as follows:
  {col\_name: {row: value, row: value,...}, col\_name: {row: value, row: value,...}, ...}
  CAPITAL letters only for col\_name, rows and values are strings. row lowest value is 1.

If excel file is chosen. the import will be done with all formulas

**Exporting**
These are the available files for exporting, the menu in the app is clear.
The json file can be imported later if exported informed, so does the excel.

| File Type | Values Export | Info Export | Notes                                    |
| --------- | ------------- | ----------- | ---------------------------------------- |
|   JSON    |       X       |      X      |                                          |
|   YAML    |       X       |      X      |                                          |
|   Excel   |       X       |      X      | Not all functions are supported by Excel |
|   CSV     |       X       |             |                                          |
|   PDF     |       X       |             |                                          |

---

## Usage

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the app:

```bash
python3 main.py
```

Default: opens a 50×26 spreadsheet.

Custom size:

```bash
python3 main.py -s ROWS COLS
```

Example:

```bash
python3 main.py -s 20 10
```

---

## Testing

Run the tests with:

```bash
pytest
```

## Demo Video

You can watch a walkthrough of the project here:  
[Project Explanation Video](https://www.youtube.com/your-link-here)
