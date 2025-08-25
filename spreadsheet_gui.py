from libraries import *
from tkinter import messagebox
import tkinter.colorchooser as colorchooser
from tkinter import filedialog
from spreadsheet import Spreadsheet
from spreadsheet_types import *
from spreadsheet_errors import *
from file_manager import FileManager


class SpreadsheetGui:
    def __init__(self, spreadsheet: Spreadsheet):

        # Setting class needed objects
        self.spreadsheet: Spreadsheet = spreadsheet
        self.rows_names = self.spreadsheet.cells_df.index.to_list()
        self.cols_names = self.spreadsheet.cells_df.columns.to_list()

        self.file_manager: FileManager = FileManager(self.spreadsheet)


        # Basic settings
        self.root = tk.Tk()
        self.root.geometry("950x600")
        self.root.title("Spreadsheet Application")

        self.cells_entries: Dict[str, tk.Entry] = {}

        self.create_widgets()
        self.bind_events()


    def create_widgets(self):
        """
        Function create_widgets
        The function takes care of the general structure of the scree
        """

        self.create_toolbar()
        self.create_formula_bar()

        # Creating canvas
        self.canvas = tk.Canvas(self.root, height=10)
        self.canvas.pack(side="left", fill="both", expand=True)


        # Creating vertical scrollbar
        self.scrollbar_vertical = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollbar_vertical.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar_vertical.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Horizontal scrollbar
        self.scrollbar_horizontal = tk.Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        self.scrollbar_horizontal.pack(side="bottom", fill="x", anchor="w")
        self.canvas.configure(xscrollcommand=self.scrollbar_horizontal.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.create_grid()


    def create_formula_bar(self):
        """
        Function create_formula_bar
        The function creates the formula bar as readonly
        """
    
        # Create a frame for the formula bar
        self.formula_frame = tk.Frame(self.root)
        self.formula_frame.pack(fill="x")

        # Create the formula label
        self.formula_label = tk.Label(self.formula_frame, text="Formula:", font=font.Font(weight="bold"))
        self.formula_label.pack(side="left")

        # Create the formula entry
        self.formula_entry = tk.Entry(self.formula_frame)
        self.formula_entry.pack(side="left", fill="x", expand=True)
        self.formula_entry.config(state='readonly', readonlybackground=self.formula_entry.cget("bg"))


    def create_grid(self):
        """
        Function create_grid
        The function creates the grid of cells

        """

        # Create the spreadsheet display
        self.grid_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.grid_frame, anchor="nw")

        # Corner
        self.blank_lable = tk.Label(self.grid_frame, text='|||', width=3, background='white')
        self.blank_lable.grid(row=0, column=0)


        self.fill_grid()


    def fill_grid(self):
        """
        Function fill_grid
        The function creates and fills the entries in th grid with values from the spreadsheet.
        """
        # Create column lables
        for col in range(len(self.cols_names)):
            lable = tk.Label(self.grid_frame, text=self.cols_names[col], width=10, background='white')
            lable.grid(row=0, column=col+1)

        # Create rows lables
        for r in self.rows_names:
            lable = tk.Label(self.grid_frame, text=str(r), width=3, background='white')
            lable.grid(row = r, column=0)


        # Flag for showing color error only once, only to notify
        fg_error_showed = False
        bg_error_showed = False

        # Creating empty cell entries
        for i in self.rows_names:
            for j in range(len(self.cols_names)):

                cell_id = self.cols_names[j] + str(i)

                #if cell_id not in self.cells_entries:
                curr_cell_info = self.spreadsheet.get_cells_info(cell_id)
                self.cells_entries[cell_id] = tk.Entry(self.grid_frame, width=10)

                try:
                    self.cells_entries[cell_id].config(foreground=curr_cell_info['fg color'])
                except:
                    if not fg_error_showed:
                        messagebox.showerror('Foreground Color Error', f"Text color for one or more cells is illegal, setting defult color for each")
                        fg_error_showed = True
                try:
                    self.cells_entries[cell_id].config(background=curr_cell_info['bg color'])
                except:
                    if not bg_error_showed:
                        messagebox.showerror('Foreground Color Error', f"Background color for one or more cells is illegal, setting defult color for eacg")
                        bg_error_showed = True

                self.cells_entries[cell_id].grid(row=i  , column=j+1)

                # Adding value
                self.cells_entries[cell_id].delete(0, tk.END)
                self.cells_entries[cell_id].insert(0, curr_cell_info['value'])
                self.cells_entries[cell_id].config(state="readonly", readonlybackground= self.cells_entries[cell_id].cget("bg"))


    def reset_grid(self):
        """
        Function reset_grid
        The function resets the grid by erasing all of the entries
        """

        cells_ids = list(self.cells_entries.keys())

        # Erasing entries and emptying the entries dictionary
        for cell_id in cells_ids:
            self.cells_entries[cell_id].destroy()
            self.cells_entries.pop(cell_id)


    def create_toolbar(self):
        """
        Function create_toolbar
        The function creates the toolbar design
        """

        # Create a frame and lable for the toolbar
        self.toolbar_frame = tk.Frame(self.root)
        self.toolbar_frame.pack(side="top", fill="x")

        self.toolbar_title_lable = tk.Label(self.toolbar_frame, text="Toolbar", font=font.Font(weight="bold"))
        self.toolbar_title_lable.pack(side="top")

        # Setting toolbar area
        self.set_files_handling_area()
        self.text_formatting_area()


    def set_files_handling_area(self):
        self.files_export_frame = tk.Frame(self.toolbar_frame)
        self.files_export_frame.pack(side="left")

        self.files_export_title = tk.Label(self.files_export_frame, text="Export File", font=font.Font(weight="bold"))
        self.files_export_title.pack(side="top")

        self.export_formula_menu = tk.Menu(self.root, tearoff=0)

        self.export_button_formula = tk.Button(self.files_export_frame, text="Informed", command=self.show_export_formula)
        self.export_button_formula.pack(side="right")


        self.export_formula_menu.add_command(label="Export as JSON", command=lambda: self.export(FileType.JSON, ExportType.INCLUDE_INFO))
        self.export_formula_menu.add_command(label="Export as YAML", command=lambda: self.export(FileType.YAML, ExportType.INCLUDE_INFO))
        self.export_formula_menu.add_command(label="Export as EXCEL", command=lambda: self.export(FileType.EXCEL, ExportType.INCLUDE_INFO))


        self.export_button_values = tk.Button(self.files_export_frame, text="Values Only", command=self.show_export_values)
        self.export_button_values.pack(side="left")

        self.export_values_menu = tk.Menu(self.root, tearoff=0)

        self.export_values_menu.add_command(label="Export as JSON", command=lambda: self.export(FileType.JSON, ExportType.VALUES_ONLY))
        self.export_values_menu.add_command(label="Export as EXCEL", command=lambda: self.export(FileType.EXCEL, ExportType.VALUES_ONLY))
        self.export_values_menu.add_command(label="Export as CSV", command=lambda: self.export(FileType.CSV, ExportType.VALUES_ONLY))
        self.export_values_menu.add_command(label="Export as PDF", command=lambda: self.export(FileType.PDF, ExportType.VALUES_ONLY))

        
        # IMPORT SECTIPN #
        self.files_import_frame = tk.Frame(self.toolbar_frame)
        self.files_import_frame.pack(side='left')

        self.import_label = tk.Label(self.files_import_frame, text='Import File', font=font.Font(weight="bold"))
        self.import_label.pack(side='top')

        self.json_options_menu = tk.Menu(self.root, tearoff=0)

        self.import_button_json = tk.Button(self.files_import_frame, text='Json File', command=self.show_json_options)
        self.import_button_json.pack(side='right')

        self.json_options_menu.add_command(label='Informed', command=lambda: self.import_file(FileType.JSON, ImportType.FORMULAS))
        self.json_options_menu.add_command(label='Values Only', command=lambda: self.import_file(FileType.JSON, ImportType.VALUES_ONLY))

        self.import_excel_button = tk.Button(self.files_import_frame, text='Excel File', command=lambda: self.import_file(FileType.EXCEL, ImportType.FORMULAS))
        self.import_excel_button.pack(side="left")


    def show_json_options(self):
        self.json_options_menu.post(self.import_button_json.winfo_rootx(), self.import_button_json.winfo_rooty() + self.import_button_json.winfo_height())


    def show_export_formula(self):
        self.export_formula_menu.post(self.export_button_formula.winfo_rootx(), self.export_button_formula.winfo_rooty() + self.export_button_formula.winfo_height())


    def show_export_values(self):
        self.export_values_menu.post(self.export_button_values.winfo_rootx(), self.export_button_values.winfo_rooty() + self.export_button_values.winfo_height())


    def import_file(self, file_type: FileType, mode: ImportType) -> None:
        """
        Function import_file
        The function takes car of file importing. it gets the wanted file from the user.

        Parameters:
            * file_type: FileType
            * mode: ImportType
        
        Return Value: None
        """

        # Warning regarding current file
        messagebox.showwarning(message='You are about to import a file, If you wish to save the current file information, Make sure to export it first.\nRecommended: Json Informed Export.')

        # Notifying excel functions
        if file_type is FileType.EXCEL:
            messagebox.showwarning('Excel Files', 'You are trying to import an excel file, please be aware that not all excel functions are supported, Info may not load as excpected if any use in those functions is done.')

        # opening file
        try:
            filename = filedialog.askopenfilename(defaultextension=file_type.value, title='Spreadsheet App - Choose a file to import')
        except Exception as e:
            messagebox.showerror(str(e))


        try:
            if filename:
                
                # Handling import
                self.file_manager.import_file(filename, file_type, mode)
                
                # Creating new grid
                self.reset_grid()
                
                self.rows_names = self.spreadsheet.cells_df.index.to_list()
                self.cols_names = self.spreadsheet.cells_df.columns.to_list()
                
                self.fill_grid()
                self.bind_events()
            else:
                messagebox.showinfo(message='No file was chosen')

        # Handling Errors
        except SpreadsheetError as e:
            messagebox.showerror(f'{e}', f"An Error has occured: {e}")

        except FileError as e:
            messagebox.showerror(f'{e}', f"An Error has occured: {e}")
 

    def export(self, file_type: FileType, export_type: ExportType) -> None:
        """
        Function export
        The function exports type based on the user's choice of file type and export.

        Parameters:
            * file_type: FileType
            * export_type: ExportType
        
        Return Value: None
        """

        if file_type is FileType.EXCEL:
            messagebox.showwarning('Excel Files', 'You are trying to export an excel type.\nPlease be aware that some formulas may not be supported by excel (SUM_BY_RANGE, AVERAGE_BY_RANGE)')

        # Getting path to save the file in
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=file_type.value, title="Spreadsheet App - Choose exporting location")
        except Exception as e:
            messagebox.showerror(str(e))


        # Exporting and handling errros
        try:
            if file_path:
                self.file_manager.export_to_file(file_path, file_type, export_type)
            else:
                messagebox.showinfo(message='No file was exported')
        except FileError as e:
            messagebox.showerror("An error has occured During file saving", str(e))


    def text_formatting_area(self):
        """
        Function text_formatting_area
        The function creates the area of the file formatting including cell foreground color and background color
        """

        
        # Creating frame and label
        self.text_format_frame = tk.Frame(self.toolbar_frame)
        self.text_format_frame.pack(side='left')

        self.txt_color_label = tk.Label(self.text_format_frame, text="Cell Formatting", font=font.Font(weight="bold"))
        self.txt_color_label.pack(side='top')

        # Creating buttons
        self.txt_color_button = tk.Button(self.text_format_frame, text="Text Color", command=self.choose_text_color)
        self.txt_color_button.pack(side="right")

        self.bg_color_button = tk.Button(self.text_format_frame, text="Background Color", command=self.choose_bg_color)
        self.bg_color_button.pack(side='left')


    def choose_text_color(self):
        """
        Function choose_text_color
        The Function shows the menu for colors for foreground text color
        """

        color = colorchooser.askcolor(title= "Choose Text Color")
        
        # Applyinh color
        if color:
            self.apply_text_color(color[1])


    def choose_bg_color(self):
        """
        Function choose_bg_color
        The function shows for colors for background color
        """

        color = colorchooser.askcolor(title= "Choose Background Color")

        # Applying color
        if color:
            self.apply_bg_color(color[1])


    def apply_text_color(self, color: str):
        """
        Function appply_text_color
        The function sets the text color to the focused entry

        parameter: color
        """

        # Getting needed focused cell
        focused_cell: tk.Entry = self.root.focus_get()

        # Making sure to color only entries within the grid
        for cell_entry in self.cells_entries.values():
            if focused_cell == cell_entry:
                focused_cell.config(fg=color)
                self.save_change_in_color(focused_cell, fg=color)
                break


    def apply_bg_color(self, color: str):
        """
        Function apply_by_color
        The function sets the text color to the focused entry

        parameter: color
        """

        # Getting needed focused cell        
        focused_cell: tk.Entry = self.root.focus_get()
       
        # Making sure to color only entries within the grid
        for cell_entry in self.cells_entries.values():
            if focused_cell == cell_entry:
                focused_cell.config(bg=color)
                self.save_change_in_color(focused_cell, bg=color)
                break

    
    def save_change_in_color(self, focused_cell: tk.Entry, fg=None, bg=None):
        """
        Function save_change_in_color
        The function updates the actual cell object of the cell with the new color
        """
        # Getting cell id
        for cell_id, entry in self.cells_entries.items():
            if entry == focused_cell:
                
                # Getting cell
                cell = self.spreadsheet.__get_cell_by_name__(cell_id)
                
                # Updating color
                if fg:
                    cell.set_design(fg=fg)
                if bg:
                    cell.set_design(bg=bg)

                break


    def bind_events(self):
        """
        Function bind_events
        The function takes care of binding keys and events to the relevant functions per cell
        """

        # Bind events for cell updates
        for i in self.rows_names:
            for j in range(len(self.cols_names)):
                cell_id = self.cols_names[j]+str(i)
                cell_entry = self.cells_entries[cell_id]
                cell_entry.bind("<FocusIn>", self.update_formula_entry)
                cell_entry.bind("<FocusOut>", self.save_cell_value)
                cell_entry.bind("<Left>", self.navigate)
                cell_entry.bind("<Right>", self.navigate)
                cell_entry.bind("<Up>", self.navigate)
                cell_entry.bind("<Down>", self.navigate)
                cell_entry.bind("<Shift-Tab>", self.navigate)
                cell_entry.bind("<Tab>", self.navigate)
                cell_entry.bind("<F2>", lambda event, i=i, j=j: self.start_editing(i, j))
                cell_entry.bind("<Double-Button-1>", lambda event, i=i, j=j: self.start_editing(i, j))
                cell_entry.bind("<Return>", lambda event, i=i, j=j: self.start_editing(i, j))


    def start_editing(self, i, j):
        """
        Function start_editing
        The function takes the cell out of readonly mode and sets it to editing mode
        """

        # Getting entry by index and updating its state
        cell_entry = self.cells_entries[self.cols_names[j]+i]
        cell_entry.config(state="normal")


    def update_formula_entry(self, event: tk.Event):
        """
        Function update_formula_entry
        The function takes care of updating the formula bar value based on the focused cell

        parameters:
            * event: tk.Event - The event that led to the function
        """

        # Getting matching cell
        for cell_id, entry in self.cells_entries.items():
            if entry == event.widget:

                # Getting cell's formula
                formula = self.spreadsheet.get_cell_formula(cell_id)

                # Making sure to actually show the formula and not the current cell value
                if self.formula_entry.get() != formula:
                    
                    # Editing the formula entry
                    self.formula_entry.config(state='normal')
                    self.formula_entry.delete(0, tk.END)
                    self.formula_entry.insert(0, formula)
                    self.formula_entry.config(state='readonly', readonlybackground=self.formula_entry.cget("bg"))


    def navigate(self, event: tk.Event):
        """
        Function navigate
        The function updates the focused cell based on the move the user chose
        """

        # Getting needed entry for the move
        curr_entry = event.widget
        found_entry = False

        # Getting index
        for row in self.rows_names:
            for col, col_name in enumerate(self.cols_names):
                if self.cells_entries[col_name+str(row)] == curr_entry:
                    found_entry = True
                    break
            if found_entry:
                break

        # Only moving if focused widget is an actual entry
        if not found_entry:
            return

        row = int(row)
        rows = len(self.rows_names)

        # Update row and column indices based on the pressed key
        if event.keysym == 'Right' or event.keysym == 'Tab':
            col = col + 1 if col < len(self.cols_names)-1 else col

        elif event.keysym == 'Left' or event.keysym == 'Shift-Tab':
            col = col - 1 if col > 0 else col

        elif event.keysym == 'Up':
            row = row - 1 if row > 1 else row

        elif event.keysym == 'Down':
            row = row + 1 if row < rows else rows
        else:
            print("Ilegal move")
        
        # Setting the new focus
        new_entry = self.cells_entries[self.cols_names[col]+str(row)]
        new_entry.focus()


    def save_cell_value(self, event: tk.Event):
        """
        Function save_cell_value
        The function updates the value of the cell in the spreadsheet to the new values.

        Parameters:
            * event: tk.Event
        """

        # Getting cell
        for cell_id, entry in self.cells_entries.items():
            if entry == event.widget:
                new_value = entry.get()
                old_value = self.spreadsheet.get_value_from_cell(cell_id)
                try:
                    # Updating only if the value has changed
                    if new_value != old_value:
                        self.spreadsheet.edit_cell(cell_id, new_value)
                # Taking care of errors
                except SpreadsheetError as e:
                    messagebox.showerror(message=e)

                entry.delete(0, tk.END)
                entry.insert(0, self.spreadsheet.get_value_from_cell(cell_id))
                self.populate_entries()
                entry.config(state="readonly", readonlybackground=entry.cget("bg"))
                break


    def populate_entries(self):
        """
        Function populate_entries
        The function fills the entries with the values from the spreadsheet based on formula and colors
        """
        for i in self.rows_names:
            for j in range(len(self.cols_names)):
                
                # Getting cell and its values
                cell_id = self.cols_names[j] + str(i)
                cell_value = self.spreadsheet.get_value_from_cell(cell_id)
                formula = self.spreadsheet.get_cell_formula(cell_id)

                # Insert cell' value
                self.cells_entries[cell_id].config(state='normal')
                self.cells_entries[cell_id].delete(0, tk.END)
                self.cells_entries[cell_id].insert(0, cell_value)
                self.cells_entries[cell_id].config(state="readonly", readonlybackground= self.cells_entries[cell_id].cget("bg"))
                
                # inserting formula
                self.formula_entry.delete(0,tk.END)
                self.formula_entry.insert(0,formula)


    def run(self):
        """
        Function run
        Gets called only to run the program
        """
        self.root.mainloop()

