import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from MCO import MCO
from functools import partial

class MainWindow:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title('MCO')
        window_width = 900
        window_height = 500
        self.window.geometry(f'{window_width}x{window_height}')
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f'+{x}+{y}')

        self.create_gui()

        self.weight_wind = None
        self.data = pd.DataFrame
        self.criteria_num = 0
        self.treeview = None
        self.mco = None
        self.weights = {}
        self.directions = {}
        self.gc_direction = 'max'

        self.lower_restrictions = {}
        self.upper_restrictions = {}

        self.additive_best = None
        self.mc_best = None
        self.br_best = None

        self.is_normalized = False
        self.is_additived = False
        self.is_mc = False
        self.is_binary_relations = False
    
    def show(self):
        self.window.mainloop()

    def open_file(self):
        filetypes = (('Excel Files', '*.xlsx;*.xls'), ('All Files', '*.*')) 
        initialdir = os.path.dirname(os.path.abspath(__file__))

        self.file_path = filedialog.askopenfilename(initialdir=initialdir, filetypes=filetypes)
        self.data = pd.read_excel(self.file_path)
        self.criteria_num = len(self.data.columns) - 1
        self.mco = MCO(self.data, self.treeview)
        self.create_table()

        self.btn_normalize.configure(state='disabled')
        self.btn_binary_relations.configure(state='disabled')
        self.btn_additive.configure(state='disabled')
        self.btn_main_criterion.configure(state='disabled')
        self.weights.clear()
        self.directions.clear()
        self.btn_add_weight_direction.configure(state='normal')
        self.is_normalized = False
        self.is_additived = False
        self.is_mc = False
        self.is_binary_relations = False
        self.additive_best = None
        self.br_best = None
        self.mc_best = None

    def create_gui(self):
        label_text = "Multi-criteria Optimization"
        label = tk.Label(self.window, text=label_text, font=('Arial', 23, 'bold'), fg='gray', anchor='center')
        label.pack()

        label_text = "Open File"
        label = tk.Label(self.window, text=label_text, font=('Arial', 70, 'bold'), fg='lightgray', anchor='center')
        label.place(x=130, y=150)

        self.create_buttons()

        label_text = "Generalized Criterion"
        label = tk.Label(self.window, text=label_text, font=('Arial', 20, 'bold'), fg='gray', anchor='center')
        label.place(x=590, y=380)

        btn_gc_min_max = tk.Button(self.window, text='max', width=10, height=2, font=('Arial', 9, 'bold'), bg='lightgreen', fg='white')
        btn_gc_min_max.configure(command=partial(self.gc_direction, btn_gc_min_max))
        btn_gc_min_max.place(x=690, y=420)
    
    def gc_direction(self, btn):
        if btn.cget('text') == 'min':
            btn.configure(text='max', bg='lightgreen')
            self.gc_direction = 'max'
        else:
            btn.configure(text='min', bg='pink')
            self.gc_direction = 'min'

        if self.is_normalized:
            self.normalize()

        if self.is_additived:
            self.additive()

        if self.is_mc:
            self.on_restriction_window_close()

    def create_table(self):
        for child in self.window.winfo_children():
            if isinstance(child, ttk.Treeview):
                child.destroy()

        frame = ttk.Frame(self.window)
        frame.place(x=20, y=70, width=600, height=300)

        self.treeview = ttk.Treeview(frame, show='headings')
        self.treeview.grid(row=0, column=0, sticky='nsew')

        scrollbar_y = ttk.Scrollbar(frame, orient='vertical', command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=scrollbar_y.set)
        scrollbar_y.grid(row=0, column=1, sticky='ns')

        scrollbar_x = ttk.Scrollbar(frame, orient='horizontal', command=self.treeview.xview)
        self.treeview.configure(xscrollcommand=scrollbar_x.set)
        scrollbar_x.grid(row=1, column=0, sticky='ew')

        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        columns = tuple(self.data.columns)
        self.treeview['columns'] = columns

        for column in columns:
            self.treeview.heading(column, text=column, anchor='center')
            self.treeview.column(column, width=50, stretch=False, anchor='center')

        for row in self.data.itertuples(index=False):
            self.treeview.insert('', 'end', values=row)

    def create_buttons(self):
        width = 20
        height = 2
        bg = 'blue'
        fg = 'white'
        font = ('Arial', 9, 'bold')
        x = 150 
        dx = 170
        y = 400

        btn_open_file = tk.Button(self.window, text='Open File', command=self.open_file)
        btn_open_file.configure(width=width, height=height, bg=bg, fg=fg, font=font)
        btn_open_file.place(x=x, y=y)

        btn_save = tk.Button(self.window, text='Save', command=self.save)
        btn_save.configure(width=width, height=height, bg='green', fg=fg, font=font)
        btn_save.place(x=x+dx, y=y)

        width = 25
        height = 2
        bg = 'purple'
        fg = 'white'
        font = ('Arial', 9, 'bold')
        x = 670
        y = 70
        dy = 60

        self.btn_main_criterion = tk.Button(self.window, text='Main Criterion', command=self.main_criterion)
        self.btn_main_criterion.configure(width=width, height=height, bg=bg, fg=fg, font=font, state='disabled')
        self.btn_main_criterion.place(x=x, y=y+4*dy)

        self.btn_additive = tk.Button(self.window, text='Additive', command=self.additive)
        self.btn_additive.configure(width=width, height=height, bg=bg, fg=fg, font=font, state='disabled')
        self.btn_additive.place(x=x, y=y+2*dy)

        self.btn_binary_relations = tk.Button(self.window, text='Binary Relations', command=self.do_binary_relations)
        self.btn_binary_relations.configure(width=width, height=height, bg=bg, fg=fg, font=font, state='disabled')
        self.btn_binary_relations.place(x=x, y=y+3*dy)

        self.btn_add_weight_direction = tk.Button(self.window, text='Add weight coefficients\nand specify the direction', command=self.add_weight_direction)
        self.btn_add_weight_direction.configure(width=width, height=height, bg='blue', fg=fg, font=font, state='disabled')
        self.btn_add_weight_direction.place(x=x, y=y)

        self.btn_normalize = tk.Button(self.window, text='Normalize', command=self.normalize)
        self.btn_normalize.configure(width=width, height=height, bg='green', fg=fg, font=font, state='disabled')
        self.btn_normalize.place(x=x, y=y+dy)

    def save(self):
        if self.additive_best is not None:
            self.data.at[0, 'add_best'] = self.additive_best
        if self.mc_best is not None:
            self.data.at[0, 'mc_best'] = self.mc_best
        if self.br_best is not None:
            self.data.at[0, 'br_best'] = self.br_best 

        filename, extension = self.file_path.split('.')
        new_path = filename + '_final' + '.' + extension
        self.data.to_excel(new_path, index=False)
        messagebox.showinfo('Data Saved', 'Data was saved successfully.\nPath: ' + new_path)
        self.window.lift()

        if self.additive_best is not None:
            self.data.drop(columns=['add_best'], inplace=True)
        if self.mc_best is not None:
            self.data.drop(columns=['mc_best'], inplace=True)
        if self.br_best is not None:
            self.data.drop(columns=['br_best'], inplace=True)

    def add_weight_direction(self):
        self.weight_wind = WieghtsAndDirectionsWindow(self.data, self.criteria_num, self.weights, self.directions)
        self.weight_wind.show(self.on_weight_direction_window_close)

    def on_weight_direction_window_close(self):
        self.weights = self.weight_wind.weights
        self.directions = self.weight_wind.directions

        self.btn_normalize.configure(state='normal')
        self.btn_binary_relations.configure(state='normal')

        self.mco.do_binary_relations(self.directions)

        if self.is_normalized:
            self.normalize()

        if self.is_additived:
            self.additive()

        if self.is_mc:
            self.on_restriction_window_close()
        
        if self.is_binary_relations:
            self.mco.add_br_column()
            self.create_table()

    def normalize(self):
        self.mco.normalize(self.directions, self.gc_direction)

        self.create_table()

        self.btn_additive.configure(state='normal')
        self.btn_main_criterion.configure(state='normal')

        self.is_normalized = True

    def main_criterion(self):
        self.restrictions_wind = RestrictionsWindow(self.data, self.criteria_num, self.lower_restrictions, self.upper_restrictions)
        self.restrictions_wind.show(self.on_restriction_window_close)
    
    def on_restriction_window_close(self):
        self.lower_restrictions = self.restrictions_wind.lower_restrictions
        self.upper_restrictions = self.restrictions_wind.upper_restrictions

        self.mc_best = self.mco.main_criterion(self.weights, self.lower_restrictions, self.upper_restrictions, self.gc_direction)
        self.create_table()

        self.is_mc = True

    def additive(self):
        self.additive_best = self.mco.additive(self.weights, self.gc_direction)
        self.create_table()

        self.is_additived = True
    
    def do_binary_relations(self):
        self.br_best = self.mco.add_br_column()
        self.create_table()
        self.is_binary_relations = True
        BinaryRelationsWindow(self.mco.binary_relations).show()


class WieghtsAndDirectionsWindow:
    def __init__(self, data, criteria_num, init_weights, init_directions):
        self.window = tk.Tk()
        self.window.title('Weight coefficients and directions')
        window_width = 450
        window_height = 250
        self.window.geometry(f'{window_width}x{window_height}')
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f'+{x}+{y}')

        self.init_weights = init_weights
        self.init_directions = init_directions
        self.criteria_num = criteria_num
        self.data = data
        self.text_boxes = {}
        self.direction_buttons = {}
        self.weights = {}
        self.directions = {}

    def create_gui(self):
        frame = ttk.Frame(self.window)
        frame.place(x=20, y=20, width=180, height=170)

        canvas = tk.Canvas(frame)
        canvas.grid(row=1, column=0, columnspan=3, sticky='nsew')

        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        scrollbar.grid(row=1, column=3, sticky='ns')
        canvas.configure(yscrollcommand=scrollbar.set)

        inner_frame = ttk.Frame(canvas)
        inner_frame.grid(row=1, column=0, columnspan=3, sticky='nsew')

        canvas.create_window((0, 0), window=inner_frame, anchor='nw')

        columns = self.data.columns
        for i, column in enumerate(columns[1:self.criteria_num + 1], start=1):
            label = ttk.Label(inner_frame, text=column)
            label.grid(row=i, column=0, padx=3, pady=3, sticky='e')

            text = 'max'
            bg = 'lightgreen'
            if len(self.init_directions) > 0:
                text = self.init_directions[column]
                if text == 'min':
                    bg = 'pink'
            
            btn_direction = tk.Button(inner_frame, text=text)
            btn_direction.configure(command=partial(self.min_max_dir, btn_direction), width=5, bg=bg, fg='white')
            btn_direction.grid(row=i, column=1, padx=3, pady=3, sticky='w')

            self.direction_buttons[column] = btn_direction

            textbox = ttk.Entry(inner_frame)
            textbox.configure(width=7)
            textbox.grid(row=i, column=2, padx=3, pady=3, sticky='w')

            if len(self.init_weights) > 0:
                textbox.insert('end', self.init_weights[column])

            self.text_boxes[column] = textbox

        inner_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox('all'))

        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=0)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_columnconfigure(3, weight=0)
        frame.grid_rowconfigure(1, weight=1)

        text = ttk.Label(self.window, text='NOTE:\nThe sum of\nall coefficients\nmust be equal\nto 1.', foreground='gray', font=('Arial', 21, 'bold'))
        text.place(x=220, y=20)

        btn_save = tk.Button(self.window, text='Save', command=self.save)
        btn_save.configure(width=20, bg='green', fg='white')
        btn_save.place(x=20, y=200)

    def save(self):
        self.directions.clear()
        self.weights.clear()
        for column, button in self.direction_buttons.items():
            self.directions[column] = button.cget('text')
        
        try:
            for column, text_box in self.text_boxes.items():
                self.weights[column] = float(text_box.get())
        except ValueError:
            messagebox.showwarning('Invalid Input', 'Please enter valid numeric weights.')
            self.window.lift()
            return

        total_sum = sum(float(weight) for weight in self.weights.values())

        if total_sum == 1:
            self.window.destroy()
            self.on_close_callback()
        else:
            messagebox.showwarning('Invalid coefficients', 'The sum should equal to 1.')
            self.window.lift()

    def min_max_dir(self, btn):
        if btn.cget('text') == 'min':
            btn.configure(text='max', bg='lightgreen')
        else:
            btn.configure(text='min', bg='pink')

    def show(self, on_close_callback):
        self.create_gui()
        self.on_close_callback = on_close_callback
        self.window.mainloop()

class RestrictionsWindow:
    def __init__(self, data, criteria_num, init_lower_restrictions, init_upper_restrictions):
        self.window = tk.Tk()
        self.window.title('Restrictions')
        window_width = 220
        window_height = 250
        self.window.geometry(f'{window_width}x{window_height}')
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f'+{x}+{y}')

        self.data = data
        self.criteria_num = criteria_num
        self.init_lower_restrictions = init_lower_restrictions
        self.init_upper_restrictions = init_upper_restrictions

        self.lower_restrictions = {}
        self.upper_restrictions = {}

        self.text_boxes_lower = {}
        self.text_boxes_upper = {}

    def create_gui(self):
        frame = ttk.Frame(self.window)
        frame.place(x=20, y=20, width=180, height=170)

        canvas = tk.Canvas(frame)
        canvas.grid(row=1, column=0, columnspan=3, sticky='nsew')

        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        scrollbar.grid(row=1, column=3, sticky='ns')
        canvas.configure(yscrollcommand=scrollbar.set)

        inner_frame = ttk.Frame(canvas)
        inner_frame.grid(row=1, column=0, columnspan=3, sticky='nsew')

        canvas.create_window((0, 0), window=inner_frame, anchor='nw')

        columns = self.data.columns
        for i, column in enumerate(columns[1:self.criteria_num + 1], start=1):
            label = ttk.Label(inner_frame, text='< ' + column + ' <')
            label.grid(row=i, column=1, padx=3, pady=3, sticky='w')

            textbox = ttk.Entry(inner_frame)
            textbox.configure(width=7)
            textbox.grid(row=i, column=2, padx=3, pady=3, sticky='w')

            if len(self.init_upper_restrictions) > 0:
                textbox.insert('end', self.init_upper_restrictions[column])
            else:
                textbox.insert('end', 0.9)

            self.text_boxes_upper[column] = textbox

            textbox = ttk.Entry(inner_frame)
            textbox.configure(width=7)
            textbox.grid(row=i, column=0, padx=3, pady=3, sticky='e')

            if len(self.init_lower_restrictions) > 0:
                textbox.insert('end', self.init_lower_restrictions[column])
            else:
                textbox.insert('end', 0)

            self.text_boxes_lower[column] = textbox

        inner_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox('all'))

        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=0)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_columnconfigure(3, weight=0)
        frame.grid_rowconfigure(1, weight=1)

        btn_save = tk.Button(self.window, text='Save', command=self.save)
        btn_save.configure(width=23, bg='green', fg='white')
        btn_save.place(x=20, y=200)
        
    def save(self):
        self.upper_restrictions.clear()
        self.lower_restrictions.clear()
        
        try:
            for column, text_box in self.text_boxes_lower.items():
                self.lower_restrictions[column] = float(text_box.get())

            for column, text_box in self.text_boxes_upper.items():
                self.upper_restrictions[column] = float(text_box.get())
        except ValueError:
            messagebox.showwarning('Invalid Input', 'Please enter valid numeric restrictions.')
            self.window.lift()
            return
        
        check = True
        for column in self.data.columns[1:self.criteria_num + 1]:
            if self.lower_restrictions[column] >= self.upper_restrictions[column]:
                check = False
                break

        if check:
            self.window.destroy()
            self.on_close_callback()
        else:
            messagebox.showwarning('Invalid coefficients', 'There\'s a lower bound greater than or equal to the upper bound.')
            self.window.lift()
 
    def show(self, on_close_callback):
        self.create_gui()
        self.on_close_callback = on_close_callback
        self.window.mainloop()

class BinaryRelationsWindow:
    def __init__(self, binary_relations):
        self.window = tk.Tk()
        self.window.title('Binary Relations')
        window_width = 450
        window_height = 450
        self.window.geometry(f'{window_width}x{window_height}')
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f'+{x}+{y}')

        self.binary_relations = binary_relations

    def create_gui(self):
        for child in self.window.winfo_children():
            if isinstance(child, ttk.Treeview):
                child.destroy()

        frame = ttk.Frame(self.window)
        frame.place(x=10, y=5, width=440, height=440)

        self.treeview = ttk.Treeview(frame, show='headings')
        self.treeview.grid(row=0, column=0, sticky='nsew')

        scrollbar_y = ttk.Scrollbar(frame, orient='vertical', command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=scrollbar_y.set)
        scrollbar_y.grid(row=0, column=1, sticky='ns')

        scrollbar_x = ttk.Scrollbar(frame, orient='horizontal', command=self.treeview.xview)
        self.treeview.configure(xscrollcommand=scrollbar_x.set)
        scrollbar_x.grid(row=1, column=0, sticky='ew')

        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        columns = ['N']
        for i in range(len(self.binary_relations)):
            columns.append(i + 1)

        columns.append('SUM')
        self.treeview['columns'] = columns

        for column in columns:
            self.treeview.heading(column, text=column, anchor='center')
            self.treeview.column(column, width=50, stretch=False, anchor='center')

        sums = []
        for i, row in enumerate(self.binary_relations):
            sum = 0
            for j in range(len(row)):
                if j != i:
                    betters, worses = row[j].split(':')
                    sum += float(betters)
            
            sums.append(sum)

        new_rows = [[i + 1] + row + [sums[i]] for i, row in enumerate(self.binary_relations)]
        for row in new_rows:
            self.treeview.insert('', 'end', values=row)

    def show(self):
        self.create_gui()
        self.window.mainloop()