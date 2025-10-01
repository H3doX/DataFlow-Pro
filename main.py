import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import pandas as pd
import pyautogui
import pyperclip
import time
import threading
from datetime import datetime
import os
from translations import translations, get_text
from ui_improvements import (
    setup_styles, create_tooltip, create_icon_button,
    create_status_bar, show_notification, setup_keyboard_shortcuts,
    ICONS, THEMES
)

class AutomationGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)

        # Theme and language settings
        self.current_language = self.load_language_preference()
        self.current_theme = self.load_theme_preference()

        # Apply theme styling
        self.style, self.text_config = setup_styles(self.root, self.current_theme)
        # Store theme reference for notifications
        self.root._current_theme = self.current_theme

        # Set window title with custom icon
        self.root.title(f"🌊 {get_text('app_title', self.current_language)} - Smart Data Automation")

        # Data storage
        self.excel_data = None
        self.excel_columns = []
        self.excel_sheets = []
        self.current_file_path = None
        self.automation_steps = []
        self.current_preset = None
        self.presets_folder = "presets"

        # Ensure presets folder exists
        if not os.path.exists(self.presets_folder):
            os.makedirs(self.presets_folder)

        self.create_gui()

        # Create status bar
        self.status_bar = create_status_bar(self.root, self.current_theme)
        self.update_status("Ready")

        # Setup keyboard shortcuts
        self.setup_shortcuts()

    def create_gui(self):
        # Main container frame with padding
        main_container = ttk.Frame(self.root, style="Card.TFrame")
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Main notebook for tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Tab 1: Excel Configuration
        self.excel_frame = ttk.Frame(self.notebook, style="Card.TFrame")
        self.notebook.add(self.excel_frame, text=f"{ICONS['excel']} {get_text('excel_data', self.current_language)}")
        self.create_excel_tab()

        # Tab 2: Automation Steps
        self.automation_frame = ttk.Frame(self.notebook, style="Card.TFrame")
        self.notebook.add(self.automation_frame, text=f"{ICONS['settings']} {get_text('automation_steps', self.current_language)}")
        self.create_automation_tab()

        # Tab 3: Execution
        self.execution_frame = ttk.Frame(self.notebook, style="Card.TFrame")
        self.notebook.add(self.execution_frame, text=f"{ICONS['play']} {get_text('execute', self.current_language)}")
        self.create_execution_tab()

        # Menu bar
        self.create_menu()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu with icons
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=f"{ICONS['file']} {get_text('file', self.current_language)}", menu=file_menu)
        file_menu.add_command(label=f"{ICONS['save']} {get_text('save_preset', self.current_language)}",
                            command=self.save_preset, accelerator="Ctrl+S")
        file_menu.add_command(label=f"{ICONS['load']} {get_text('load_preset', self.current_language)}",
                            command=self.load_preset, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label=f"{ICONS['error']} {get_text('exit', self.current_language)}",
                            command=self.root.quit, accelerator="Alt+F4")

        # Language menu
        language_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=f"{ICONS['language']} {get_text('language', self.current_language)}", menu=language_menu)
        language_menu.add_command(label="🇬🇧 English", command=lambda: self.change_language("en"))
        language_menu.add_command(label="🇮🇹 Italiano", command=lambda: self.change_language("it"))
        language_menu.add_command(label="🇷🇺 Русский", command=lambda: self.change_language("ru"))
        language_menu.add_command(label="🇫🇷 Français", command=lambda: self.change_language("fr"))
        language_menu.add_command(label="🇪🇸 Español", command=lambda: self.change_language("es"))
        language_menu.add_command(label="🇩🇪 Deutsch", command=lambda: self.change_language("de"))
        language_menu.add_command(label="🇨🇳 中文", command=lambda: self.change_language("zh"))

        # Theme menu
        theme_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=f"{ICONS['theme']} Theme", menu=theme_menu)
        theme_menu.add_command(label="☀️ Light", command=lambda: self.change_theme("light"))
        theme_menu.add_command(label="🌙 Dark", command=lambda: self.change_theme("dark"))

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=f"{ICONS['info']} Help", menu=help_menu)
        help_menu.add_command(label="📖 Documentation", command=self.show_help, accelerator="F1")
        help_menu.add_command(label="⌨️ Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="ℹ️ About", command=self.show_about)

    def create_excel_tab(self):
        lang = self.current_language
        # Excel file selection with modern style
        excel_section = ttk.LabelFrame(self.excel_frame, text=f"{ICONS['excel']} {get_text('excel_file', lang)}", style="TLabelframe")
        excel_section.pack(fill=tk.X, padx=10, pady=5)

        # File selection button with icon and primary style
        load_btn = create_icon_button(excel_section, "file", get_text("load_excel_file", lang),
                                     command=self.load_excel_file, style="Primary.TButton")
        load_btn.pack(side=tk.LEFT, padx=5, pady=5)
        create_tooltip(load_btn, "Click to browse and select an Excel file (Ctrl+O)")

        self.excel_file_label = ttk.Label(excel_section, text=get_text("no_file_selected", lang))
        self.excel_file_label.pack(side=tk.LEFT, padx=10, pady=5)

        # Sheet selection
        sheet_frame = ttk.Frame(excel_section)
        sheet_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(sheet_frame, text=get_text("sheet", lang)).pack(side=tk.LEFT, padx=5)
        self.sheet_combo = ttk.Combobox(sheet_frame, state="readonly", width=20)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_selected)

        load_sheet_btn = ttk.Button(sheet_frame, text=f"{ICONS['load']} {get_text('load_sheet', lang)}",
                                   command=self.load_selected_sheet, style="Secondary.TButton")
        load_sheet_btn.pack(side=tk.LEFT, padx=10)
        create_tooltip(load_sheet_btn, "Load the selected sheet's data")

        # Data preview
        self.data_preview = ttk.Label(excel_section, text=get_text("no_data_loaded", lang), foreground="gray")
        self.data_preview.pack(pady=5)

        # Column mapping
        mapping_section = ttk.LabelFrame(self.excel_frame, text=f"{ICONS['mapping']} {get_text('column_mapping', lang)}", style="TLabelframe")
        mapping_section.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create treeview for column mapping
        columns = (get_text("variable", lang), get_text("excel_column", lang), get_text("sample_data", lang))
        self.mapping_tree = ttk.Treeview(mapping_section, columns=columns, show='headings', style="Treeview")

        for col in columns:
            self.mapping_tree.heading(col, text=col)
            self.mapping_tree.column(col, width=200)

        scrollbar_mapping = ttk.Scrollbar(mapping_section, orient=tk.VERTICAL, command=self.mapping_tree.yview)
        self.mapping_tree.configure(yscrollcommand=scrollbar_mapping.set)

        self.mapping_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_mapping.pack(side=tk.RIGHT, fill=tk.Y)

        # Mapping controls
        mapping_controls = ttk.Frame(mapping_section)
        mapping_controls.pack(fill=tk.X, padx=5, pady=5)

        add_map_btn = create_icon_button(mapping_controls, "add", get_text("add_mapping", lang),
                                        command=self.add_column_mapping, style="Secondary.TButton")
        add_map_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(add_map_btn, "Map an Excel column to a variable for automation")

        remove_map_btn = create_icon_button(mapping_controls, "remove", get_text("remove_mapping", lang),
                                           command=self.remove_column_mapping, style="Secondary.TButton")
        remove_map_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(remove_map_btn, "Remove selected column mapping")

    def create_automation_tab(self):
        lang = self.current_language
        # Action types
        action_section = ttk.LabelFrame(self.automation_frame, text=f"{ICONS['add']} {get_text('add_action', lang)}", style="TLabelframe")
        action_section.pack(fill=tk.X, padx=10, pady=5)

        # Action type selection
        ttk.Label(action_section, text=get_text("action_type", lang)).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.action_type = ttk.Combobox(action_section, values=[
            get_text("click", lang), get_text("double_click", lang), get_text("right_click", lang), get_text("type_text", lang),
            get_text("key_press", lang), get_text("wait", lang), get_text("move_mouse", lang)
        ], state="readonly")
        self.action_type.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.action_type.bind('<<ComboboxSelected>>', self.on_action_type_change)

        # Dynamic parameters frame
        self.params_frame = ttk.Frame(action_section)
        self.params_frame.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)

        add_action_btn = create_icon_button(action_section, "add", get_text("add_action", lang),
                                           command=self.add_automation_step, style="Primary.TButton")
        add_action_btn.grid(row=2, column=0, padx=5, pady=10)
        create_tooltip(add_action_btn, "Add this action to the automation sequence")

        # Steps list
        steps_section = ttk.LabelFrame(self.automation_frame, text=f"{ICONS['settings']} {get_text('automation_steps_label', lang)}", style="TLabelframe")
        steps_section.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create treeview for steps
        step_columns = (get_text("step", lang), get_text("action", lang), get_text("parameters", lang), get_text("delay_s", lang))
        self.steps_tree = ttk.Treeview(steps_section, columns=step_columns, show='headings', style="Treeview")

        for col in step_columns:
            self.steps_tree.heading(col, text=col)

        self.steps_tree.column(step_columns[0], width=50)
        self.steps_tree.column(step_columns[1], width=120)
        self.steps_tree.column(step_columns[2], width=300)
        self.steps_tree.column(step_columns[3], width=80)

        scrollbar_steps = ttk.Scrollbar(steps_section, orient=tk.VERTICAL, command=self.steps_tree.yview)
        self.steps_tree.configure(yscrollcommand=scrollbar_steps.set)

        self.steps_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_steps.pack(side=tk.RIGHT, fill=tk.Y)

        # Step controls
        step_controls = ttk.Frame(steps_section)
        step_controls.pack(fill=tk.X, padx=5, pady=5)

        up_btn = create_icon_button(step_controls, "up", get_text("move_up", lang),
                                   command=self.move_step_up, style="Secondary.TButton")
        up_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(up_btn, "Move selected step up")

        down_btn = create_icon_button(step_controls, "down", get_text("move_down", lang),
                                     command=self.move_step_down, style="Secondary.TButton")
        down_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(down_btn, "Move selected step down")

        remove_btn = create_icon_button(step_controls, "remove", get_text("remove_step", lang),
                                       command=self.remove_step, style="Secondary.TButton")
        remove_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(remove_btn, "Remove selected step")


    def create_execution_tab(self):
        lang = self.current_language
        exec_section = ttk.LabelFrame(self.execution_frame, text=f"{ICONS['play']} {get_text('execution_control', lang)}", style="TLabelframe")
        exec_section.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(exec_section, text=get_text("rows_to_process", lang)).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        row_frame = ttk.Frame(exec_section)
        row_frame.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        self.process_all = tk.BooleanVar(value=True)
        ttk.Checkbutton(row_frame, text=get_text("all_rows", lang), variable=self.process_all,
                       command=self.toggle_row_selection).pack(side=tk.LEFT)

        ttk.Label(row_frame, text=get_text("from", lang)).pack(side=tk.LEFT, padx=(10, 0))
        self.from_row = tk.IntVar(value=1)
        ttk.Entry(row_frame, textvariable=self.from_row, width=5).pack(side=tk.LEFT, padx=2)

        ttk.Label(row_frame, text=get_text("to", lang)).pack(side=tk.LEFT, padx=(5, 0))
        self.to_row = tk.IntVar(value=10)
        ttk.Entry(row_frame, textvariable=self.to_row, width=5).pack(side=tk.LEFT, padx=2)

        # Execution buttons
        button_frame = ttk.Frame(exec_section)
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)

        start_btn = create_icon_button(button_frame, "play", get_text("start_automation", lang),
                                      command=self.start_automation, style="Success.TButton")
        start_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(start_btn, "Start running the automation (F5)")

        stop_btn = create_icon_button(button_frame, "stop", get_text("stop_automation", lang),
                                     command=self.stop_automation, style="Danger.TButton")
        stop_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(stop_btn, "Stop the running automation (Esc)")

        test_btn = create_icon_button(button_frame, "test", get_text("test_single_step", lang),
                                     command=self.test_single_step, style="Secondary.TButton")
        test_btn.pack(side=tk.LEFT, padx=5)
        create_tooltip(test_btn, "Test the selected step with sample data")

        # Progress and log
        progress_section = ttk.LabelFrame(self.execution_frame, text=f"{ICONS['info']} {get_text('progress_log', lang)}", style="TLabelframe")
        progress_section.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.progress = ttk.Progressbar(progress_section, mode='determinate')
        self.progress.pack(fill=tk.X, padx=5, pady=5)

        self.log_text = tk.Text(progress_section, height=15)
        if hasattr(self, 'text_config'):
            self.log_text.configure(**self.text_config)
        log_scrollbar = ttk.Scrollbar(progress_section, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

        self.running = False

    def load_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if file_path:
            try:
                # Load Excel file and get sheet names
                excel_file = pd.ExcelFile(file_path)
                self.excel_sheets = excel_file.sheet_names
                self.current_file_path = file_path

                # Update UI
                self.excel_file_label.config(text=get_text("file_loaded", self.current_language).format(os.path.basename(file_path)))
                self.sheet_combo['values'] = self.excel_sheets
                if self.excel_sheets:
                    self.sheet_combo.set(self.excel_sheets[0])  # Select first sheet by default

                # Reset preview
                self.data_preview.config(text=get_text("select_sheet_msg", self.current_language), foreground="orange")

                self.log(get_text("sheets_found", self.current_language).format(len(self.excel_sheets)))
                self.update_status(f"Excel file loaded: {len(self.excel_sheets)} sheets found")

                # Load the first sheet by default
                self.load_selected_sheet()

            except Exception as e:
                messagebox.showerror(get_text("error", self.current_language), get_text("error_loading_file", self.current_language).format(str(e)))
                self.log(get_text("error_loading_file", self.current_language).format(str(e)))

    def on_sheet_selected(self, event):
        # This method is called when a sheet is selected from dropdown
        pass

    def load_selected_sheet(self):
        lang = self.current_language
        if not self.current_file_path or not self.sheet_combo.get():
            messagebox.showwarning(get_text("warning", lang), get_text("select_file_sheet", lang))
            return

        try:
            sheet_name = self.sheet_combo.get()
            self.excel_data = pd.read_excel(self.current_file_path, sheet_name=sheet_name)
            self.excel_columns = list(self.excel_data.columns)

            # Clear existing mappings
            self.mapping_tree.delete(*self.mapping_tree.get_children())

            # Update preview
            cols_text = ', '.join(self.excel_columns[:5]) + (' ...' if len(self.excel_columns) > 5 else '')
            self.data_preview.config(
                text=get_text("sheet_info", self.current_language).format(sheet_name, len(self.excel_data), cols_text),
                foreground="green"
            )

            self.log(get_text("sheet_loaded", self.current_language).format(sheet_name, len(self.excel_data), len(self.excel_columns)))
            self.update_status(f"Sheet '{sheet_name}' loaded successfully")

            # Update row selection max values
            if hasattr(self, 'to_row'):
                self.to_row.set(len(self.excel_data))

        except Exception as e:
            self.data_preview.config(text=get_text("error_loading_sheet", self.current_language).format(str(e)), foreground="red")
            messagebox.showerror(get_text("error", self.current_language), get_text("error_loading_sheet", self.current_language).format(str(e)))
            self.log(get_text("error_loading_sheet", self.current_language).format(str(e)))

    def add_column_mapping(self):
        lang = self.current_language
        if not self.excel_columns:
            messagebox.showwarning(get_text("warning", lang), get_text("load_excel_first", lang))
            return

        dialog = ColumnMappingDialog(self.root, self.excel_columns, self.excel_data, lang)
        self.root.wait_window(dialog.dialog)  # Wait for dialog to close
        if dialog.result:
            var_name, excel_col = dialog.result
            sample_data = str(self.excel_data[excel_col].iloc[0]) if not self.excel_data.empty else "N/A"
            self.mapping_tree.insert("", tk.END, values=(var_name, excel_col, sample_data))
            self.log(get_text("added_mapping", lang).format(var_name, excel_col))

    def remove_column_mapping(self):
        selected = self.mapping_tree.selection()
        if selected:
            self.mapping_tree.delete(selected[0])

    def on_action_type_change(self, event):
        # Clear previous parameters
        for widget in self.params_frame.winfo_children():
            widget.destroy()

        lang = self.current_language
        action_type = self.action_type.get()
        self.current_params = {}

        # Map translated action names back to English for internal use
        action_map = {
            get_text("click", lang): "Click",
            get_text("double_click", lang): "Double Click",
            get_text("right_click", lang): "Right Click",
            get_text("type_text", lang): "Type Text",
            get_text("key_press", lang): "Key Press",
            get_text("wait", lang): "Wait",
            get_text("move_mouse", lang): "Move Mouse"
        }
        action_type_en = action_map.get(action_type, action_type)

        row = 0
        if action_type_en in ["Click", "Double Click", "Right Click"]:
            ttk.Label(self.params_frame, text="X:").grid(row=row, column=0, padx=5, sticky=tk.W)
            self.current_params['x'] = tk.IntVar()
            ttk.Entry(self.params_frame, textvariable=self.current_params['x'], width=10).grid(row=row, column=1, padx=5)

            ttk.Label(self.params_frame, text="Y:").grid(row=row, column=2, padx=5, sticky=tk.W)
            self.current_params['y'] = tk.IntVar()
            ttk.Entry(self.params_frame, textvariable=self.current_params['y'], width=10).grid(row=row, column=3, padx=5)

            # Coordinate capture button with icon
            capture_btn = ttk.Button(self.params_frame, text=f"{ICONS['capture']} {get_text('capture', lang)}",
                                    command=self.capture_coords_for_action, style="Secondary.TButton")
            capture_btn.grid(row=row, column=4, padx=5)
            create_tooltip(capture_btn, "Click to capture mouse coordinates (3-second countdown)")

        elif action_type_en == "Type Text":
            # Source type selection
            ttk.Label(self.params_frame, text=get_text("source", lang)).grid(row=row, column=0, padx=5, sticky=tk.W)
            self.current_params['text_source'] = ttk.Combobox(self.params_frame, values=[get_text("fixed_text", lang), get_text("excel_data_source", lang)],
                                                             state="readonly", width=12)
            self.current_params['text_source'].grid(row=row, column=1, padx=5)
            self.current_params['text_source'].bind('<<ComboboxSelected>>', self.on_text_source_change)

            # Text input (will be updated based on source selection)
            self.current_params['text'] = tk.StringVar()
            self.text_input_widget = ttk.Entry(self.params_frame, textvariable=self.current_params['text'], width=25)
            self.text_input_widget.grid(row=row, column=2, columnspan=2, padx=5)

            # Set default
            self.current_params['text_source'].set(get_text("fixed_text", lang))

        elif action_type_en == "Key Press":
            ttk.Label(self.params_frame, text=get_text("key", lang)).grid(row=row, column=0, padx=5, sticky=tk.W)
            self.current_params['key'] = ttk.Combobox(self.params_frame, values=[
                'enter', 'tab', 'esc', 'space', 'ctrl+a', 'ctrl+c', 'ctrl+v', 'ctrl+s', 'delete', 'backspace'
            ], width=15)
            self.current_params['key'].grid(row=row, column=1, padx=5)

        elif action_type_en == "Wait":
            ttk.Label(self.params_frame, text=get_text("seconds", lang)).grid(row=row, column=0, padx=5, sticky=tk.W)
            self.current_params['seconds'] = tk.DoubleVar(value=1.0)
            ttk.Entry(self.params_frame, textvariable=self.current_params['seconds'], width=10).grid(row=row, column=1, padx=5)

        elif action_type_en == "Move Mouse":
            ttk.Label(self.params_frame, text="X:").grid(row=row, column=0, padx=5, sticky=tk.W)
            self.current_params['x'] = tk.IntVar()
            ttk.Entry(self.params_frame, textvariable=self.current_params['x'], width=10).grid(row=row, column=1, padx=5)

            ttk.Label(self.params_frame, text="Y:").grid(row=row, column=2, padx=5, sticky=tk.W)
            self.current_params['y'] = tk.IntVar()
            ttk.Entry(self.params_frame, textvariable=self.current_params['y'], width=10).grid(row=row, column=3, padx=5)

            # Coordinate capture button with icon
            capture_btn = ttk.Button(self.params_frame, text=f"{ICONS['capture']} {get_text('capture', lang)}",
                                    command=self.capture_coords_for_action, style="Secondary.TButton")
            capture_btn.grid(row=row, column=4, padx=5)
            create_tooltip(capture_btn, "Click to capture mouse coordinates (3-second countdown)")


        # Delay parameter (common to all actions)
        row += 1
        ttk.Label(self.params_frame, text=get_text("delay_after", lang)).grid(row=row, column=0, padx=5, sticky=tk.W)
        self.current_params['delay'] = tk.DoubleVar(value=0.5)
        ttk.Entry(self.params_frame, textvariable=self.current_params['delay'], width=10).grid(row=row, column=1, padx=5)

    def on_text_source_change(self, event):
        lang = self.current_language
        source = self.current_params['text_source'].get()

        # Remove current widget
        self.text_input_widget.destroy()

        if source == get_text("fixed_text", lang):
            # Create entry widget for fixed text
            self.current_params['text'] = tk.StringVar()
            self.text_input_widget = ttk.Entry(self.params_frame, textvariable=self.current_params['text'], width=25)
            self.text_input_widget.grid(row=0, column=2, columnspan=2, padx=5)
        else:  # Excel Data
            # Create combobox for Excel variables
            variables = [item[0] for item in [self.mapping_tree.item(child)['values'] for child in self.mapping_tree.get_children()]]
            self.current_params['text'] = ttk.Combobox(self.params_frame, values=variables, width=22)
            self.text_input_widget = self.current_params['text']
            self.text_input_widget.grid(row=0, column=2, columnspan=2, padx=5)

    def add_automation_step(self):
        lang = self.current_language
        action_type = self.action_type.get()
        if not action_type:
            messagebox.showwarning(get_text("warning", lang), get_text("select_action_type", lang))
            return

        # Map translated action names back to English for internal use
        action_map = {
            get_text("click", lang): "Click",
            get_text("double_click", lang): "Double Click",
            get_text("right_click", lang): "Right Click",
            get_text("type_text", lang): "Type Text",
            get_text("key_press", lang): "Key Press",
            get_text("wait", lang): "Wait",
            get_text("move_mouse", lang): "Move Mouse"
        }
        action_type_en = action_map.get(action_type, action_type)

        params = {}
        params_text = []

        for key, var in self.current_params.items():
            if key == 'delay':
                continue
            if isinstance(var, tk.StringVar):
                value = var.get()
            elif isinstance(var, (tk.IntVar, tk.DoubleVar)):
                value = var.get()
            elif hasattr(var, 'get'):  # Combobox
                value = var.get()
            else:
                value = str(var)

            params[key] = value

            # Special formatting for display
            if key == 'text_source' and value == get_text("excel_data_source", lang):
                text_value = self.current_params.get('text', tk.StringVar()).get()
                params_text.append(f"Excel:{text_value}")
            elif key == 'text' and params.get('text_source') != "Excel Data":
                params_text.append(f"{key}={value}")
            elif key not in ['text', 'text_source']:
                params_text.append(f"{key}={value}")

        delay = self.current_params.get('delay', tk.DoubleVar(value=0.5)).get()

        step_data = {
            'action': action_type_en,  # Store English version internally
            'params': params,
            'delay': delay
        }

        self.automation_steps.append(step_data)

        step_num = len(self.automation_steps)
        params_display = ", ".join(params_text)
        self.steps_tree.insert("", tk.END, values=(step_num, action_type_en, params_display, delay))

    def move_step_up(self):
        selected = self.steps_tree.selection()
        if selected:
            item = selected[0]
            index = self.steps_tree.index(item)
            if index > 0:
                self.steps_tree.move(item, '', index - 1)
                self.automation_steps[index], self.automation_steps[index - 1] = \
                    self.automation_steps[index - 1], self.automation_steps[index]
                self.refresh_steps_tree()

    def move_step_down(self):
        selected = self.steps_tree.selection()
        if selected:
            item = selected[0]
            index = self.steps_tree.index(item)
            if index < len(self.automation_steps) - 1:
                self.steps_tree.move(item, '', index + 1)
                self.automation_steps[index], self.automation_steps[index + 1] = \
                    self.automation_steps[index + 1], self.automation_steps[index]
                self.refresh_steps_tree()

    def remove_step(self):
        selected = self.steps_tree.selection()
        if selected:
            item = selected[0]
            index = self.steps_tree.index(item)
            self.steps_tree.delete(item)
            del self.automation_steps[index]
            self.refresh_steps_tree()

    def refresh_steps_tree(self):
        for i, item in enumerate(self.steps_tree.get_children()):
            values = list(self.steps_tree.item(item)['values'])
            values[0] = i + 1
            self.steps_tree.item(item, values=values)

    def capture_coords_for_action(self):
        lang = self.current_language
        # Create countdown capture window
        capture_window = tk.Toplevel(self.root)
        capture_window.title(get_text("coordinate_capture", lang))
        capture_window.geometry("400x200")
        capture_window.transient(self.root)
        capture_window.grab_set()
        capture_window.attributes('-topmost', True)

        # Center the window
        capture_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))

        # Instructions
        ttk.Label(capture_window, text=get_text("position_mouse", lang),
                 font=("Arial", 12, "bold")).pack(pady=10)

        # Countdown display
        countdown_label = ttk.Label(capture_window, text="3",
                                   font=("Arial", 36, "bold"), foreground="red")
        countdown_label.pack(pady=20)

        # Current coordinates display
        coord_display = ttk.Label(capture_window, text="X: -, Y: -",
                                 font=("Arial", 11))
        coord_display.pack(pady=5)

        # Cancel button
        ttk.Button(capture_window, text=get_text("cancel", lang),
                  command=capture_window.destroy).pack(pady=10)

        countdown = [3]  # Use list to modify in nested function

        def update_display():
            x, y = pyautogui.position()
            coord_display.config(text=get_text("current_position", lang).format(x, y))

            if countdown[0] > 0:
                countdown_label.config(text=str(countdown[0]))
                countdown[0] -= 1
                capture_window.after(1000, update_display)
            else:
                # Capture coordinates
                final_x, final_y = pyautogui.position()

                # Update the action parameters
                if hasattr(self, 'current_params') and 'x' in self.current_params:
                    self.current_params['x'].set(final_x)
                    self.current_params['y'].set(final_y)

                self.log(get_text("auto_captured", lang).format(final_x, final_y))

                # Show success message briefly
                countdown_label.config(text="✓", foreground="green")
                coord_display.config(text=get_text("captured_coords", lang).format(final_x, final_y))

                capture_window.after(1000, capture_window.destroy)

        def start_fast_update():
            x, y = pyautogui.position()
            coord_display.config(text=get_text("current_position", lang).format(x, y))
            if countdown[0] > 0:
                capture_window.after(50, start_fast_update)
            else:
                update_display()

        start_fast_update()
        capture_window.after(1000, update_display)


    def toggle_row_selection(self):
        # Enable/disable row range inputs based on "All rows" checkbox
        pass

    def save_preset(self):
        preset_data = {
            'automation_steps': self.automation_steps,
            'column_mappings': [self.mapping_tree.item(child)['values'] for child in self.mapping_tree.get_children()]
        }

        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=self.presets_folder
        )

        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(preset_data, f, indent=2)
                self.log(get_text("preset_saved", self.current_language).format(filename))
                messagebox.showinfo(get_text("success", self.current_language), get_text("preset_save_success", self.current_language))
            except Exception as e:
                messagebox.showerror(get_text("error", self.current_language), get_text("preset_save_error", self.current_language).format(str(e)))

    def load_preset(self):
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=self.presets_folder
        )

        if filename:
            try:
                with open(filename, 'r') as f:
                    preset_data = json.load(f)

                # Load automation steps
                self.automation_steps = preset_data.get('automation_steps', [])
                self.steps_tree.delete(*self.steps_tree.get_children())
                for i, step in enumerate(self.automation_steps):
                    params_text = ", ".join([f"{k}={v}" for k, v in step['params'].items()])
                    self.steps_tree.insert("", tk.END, values=(i + 1, step['action'], params_text, step['delay']))

                # Load column mappings
                self.mapping_tree.delete(*self.mapping_tree.get_children())
                for mapping in preset_data.get('column_mappings', []):
                    self.mapping_tree.insert("", tk.END, values=mapping)

                self.log(get_text("preset_loaded", self.current_language).format(filename))
                messagebox.showinfo(get_text("success", self.current_language), get_text("preset_load_success", self.current_language))
            except Exception as e:
                messagebox.showerror(get_text("error", self.current_language), get_text("preset_load_error", self.current_language).format(str(e)))

    def start_automation(self):
        lang = self.current_language
        if not self.automation_steps:
            show_notification(self.root, get_text("no_automation_steps", lang), "warning")
            return

        if self.excel_data is None:
            show_notification(self.root, get_text("no_excel_data", lang), "warning")
            return

        self.running = True
        self.update_status("Automation running...")
        show_notification(self.root, "Automation started", "success")
        thread = threading.Thread(target=self.run_automation)
        thread.daemon = True
        thread.start()

    def stop_automation(self):
        self.running = False
        self.log(get_text("automation_stopped", self.current_language))
        self.update_status("Automation stopped")
        show_notification(self.root, "Automation stopped", "info")

    def test_single_step(self):
        lang = self.current_language
        selected = self.steps_tree.selection()
        if not selected:
            messagebox.showwarning(get_text("warning", lang), get_text("select_step_test", lang))
            return

        index = self.steps_tree.index(selected[0])
        step = self.automation_steps[index]

        try:
            self.execute_step(step, 0)  # Test with first row data
            self.log(get_text("test_completed", lang).format(step['action']))
        except Exception as e:
            messagebox.showerror(get_text("error", lang), get_text("test_failed", lang).format(str(e)))
            self.log(get_text("test_failed", lang).format(str(e)))

    def run_automation(self):
        start_row = 0 if self.process_all.get() else self.from_row.get() - 1
        end_row = len(self.excel_data) if self.process_all.get() else self.to_row.get()

        total_rows = end_row - start_row
        self.progress.config(maximum=total_rows)

        for row_idx in range(start_row, min(end_row, len(self.excel_data))):
            if not self.running:
                break

            self.log(get_text("processing_row", self.current_language).format(row_idx + 1))

            try:
                for step in self.automation_steps:
                    if not self.running:
                        break
                    self.execute_step(step, row_idx)
                    time.sleep(step['delay'])

                self.progress['value'] = row_idx - start_row + 1
                self.root.update_idletasks()

            except Exception as e:
                self.log(get_text("error_in_row", self.current_language).format(row_idx + 1, str(e)))
                if messagebox.askyesno(get_text("error", self.current_language), get_text("continue_next_row", self.current_language).format(row_idx + 1, str(e))):
                    continue
                else:
                    break

        self.running = False
        self.log(get_text("automation_completed", self.current_language))
        self.update_status("Automation completed successfully")
        show_notification(self.root, "Automation completed!", "success")

    def execute_step(self, step, row_idx):
        action = step['action']
        params = step['params']

        if action == "Click":
            pyautogui.click(params['x'], params['y'])

        elif action == "Double Click":
            pyautogui.doubleClick(params['x'], params['y'])

        elif action == "Right Click":
            pyautogui.rightClick(params['x'], params['y'])

        elif action == "Type Text":
            text_source = params.get('text_source', 'Fixed Text')

            if text_source == "Excel Data":
                # Find the Excel data for this variable
                variable_name = params['text']
                for child in self.mapping_tree.get_children():
                    values = self.mapping_tree.item(child)['values']
                    if values[0] == variable_name:
                        excel_col = values[1]
                        text_to_type = str(self.excel_data.iloc[row_idx][excel_col])
                        pyautogui.write(text_to_type)
                        break
            else:
                # Fixed text
                pyautogui.write(params['text'])

        elif action == "Key Press":
            if '+' in params['key']:
                keys = params['key'].split('+')
                pyautogui.hotkey(*keys)
            else:
                pyautogui.press(params['key'])

        elif action == "Wait":
            time.sleep(params['seconds'])

        elif action == "Move Mouse":
            pyautogui.moveTo(params['x'], params['y'])

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def load_language_preference(self):
        """Load saved language preference from file"""
        try:
            if os.path.exists("preferences.json"):
                with open("preferences.json", "r") as f:
                    data = json.load(f)
                    return data.get("language", "en")
        except:
            pass
        return "en"

    def load_theme_preference(self):
        """Load saved theme preference from file"""
        try:
            if os.path.exists("preferences.json"):
                with open("preferences.json", "r") as f:
                    data = json.load(f)
                    return data.get("theme", "light")
        except:
            pass
        return "light"

    def save_preferences(self):
        """Save preferences to file"""
        try:
            prefs = {
                "language": self.current_language,
                "theme": self.current_theme
            }
            with open("preferences.json", "w") as f:
                json.dump(prefs, f, indent=2)
        except:
            pass

    def change_language(self, lang):
        """Change the application language"""
        self.current_language = lang
        self.save_preferences()
        # Update all UI elements
        self.update_ui_language()
        self.update_status(f"Language changed to {lang}")

        # Build multilanguage success message
        success_messages = {
            "en": "Language changed successfully!",
            "it": "Lingua cambiata con successo!",
            "ru": "Язык успешно изменен!",
            "fr": "Langue changée avec succès!",
            "es": "¡Idioma cambiado con éxito!",
            "de": "Sprache erfolgreich geändert!",
            "zh": "语言更改成功！"
        }

        # Show notification instead of messagebox
        show_notification(self.root, success_messages.get(lang, success_messages["en"]), "success")

    def update_ui_language(self):
        """Update all UI text with current language"""
        lang = self.current_language

        # Update window title
        self.root.title(f"🌊 {get_text('app_title', lang)} - Smart Data Automation")

        # Update notebook tabs with icons
        self.notebook.tab(0, text=f"{ICONS['excel']} {get_text('excel_data', lang)}")
        self.notebook.tab(1, text=f"{ICONS['settings']} {get_text('automation_steps', lang)}")
        self.notebook.tab(2, text=f"{ICONS['play']} {get_text('execute', lang)}")

        # Recreate UI elements with new language
        self.recreate_ui_elements()

    def recreate_ui_elements(self):
        """Recreate UI elements with updated language"""
        # Clear and recreate the frames content
        for widget in self.excel_frame.winfo_children():
            widget.destroy()
        for widget in self.automation_frame.winfo_children():
            widget.destroy()
        for widget in self.execution_frame.winfo_children():
            widget.destroy()

        # Recreate tabs content
        self.create_excel_tab()
        self.create_automation_tab()
        self.create_execution_tab()

        # Recreate menu
        self.create_menu()

        # Restore data if exists
        if hasattr(self, 'excel_file_label') and self.current_file_path:
            self.excel_file_label.config(text=get_text("file_loaded", self.current_language).format(os.path.basename(self.current_file_path)))

        # Restore steps and mappings if they exist
        self.restore_steps_and_mappings()

    def restore_steps_and_mappings(self):
        """Restore automation steps and column mappings after language change"""
        # Restore automation steps
        for i, step in enumerate(self.automation_steps):
            params_text = []
            for key, value in step['params'].items():
                if key == 'text_source' and value == "Excel Data":
                    text_value = step['params'].get('text', '')
                    params_text.append(f"Excel:{text_value}")
                elif key == 'text' and step['params'].get('text_source') != "Excel Data":
                    params_text.append(f"{key}={value}")
                elif key not in ['text', 'text_source']:
                    params_text.append(f"{key}={value}")

            params_display = ", ".join(params_text)
            self.steps_tree.insert("", tk.END, values=(i + 1, step['action'], params_display, step['delay']))

    def change_theme(self, theme):
        """Change application theme"""
        self.current_theme = theme
        self.save_preferences()
        self.style, self.text_config = setup_styles(self.root, theme)
        self.root._current_theme = theme  # Store for notifications
        self.recreate_ui_elements()
        # Update status bar colors
        if hasattr(self, 'status_bar'):
            colors = THEMES[theme]
            self.status_bar.configure(bg=colors["frame_bg"], fg=colors["label_fg"])
            self.status_bar.master.configure(bg=colors["frame_bg"])
        self.update_status(f"Theme changed to {theme}")
        show_notification(self.root, f"Theme changed to {theme}", "success")

    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        shortcuts = {
            '<Control-o>': self.load_excel_file,
            '<Control-s>': self.save_preset,
            '<F5>': self.start_automation,
            '<Escape>': self.stop_automation,
            '<F1>': self.show_help,
        }
        setup_keyboard_shortcuts(self.root, shortcuts)

    def update_status(self, message):
        """Update status bar message"""
        if hasattr(self, 'status_bar'):
            self.status_bar.config(text=f"{ICONS['info']} {message}")

    def show_help(self):
        """Show help dialog"""
        help_text = """DataFlow Pro - Help

Keyboard Shortcuts:
• Ctrl+O: Load Excel file
• Ctrl+S: Save preset
• F5: Start automation
• Esc: Stop automation
• F1: Show this help

How to use:
1. Load an Excel file with your data
2. Map columns to variables
3. Create automation steps
4. Run the automation

For more information, visit the documentation."""
        messagebox.showinfo("Help", help_text)

    def show_shortcuts(self):
        """Show keyboard shortcuts"""
        shortcuts_text = """Keyboard Shortcuts:

• Ctrl+O - Load Excel file
• Ctrl+S - Save preset
• F5 - Start automation
• Esc - Stop automation
• F1 - Show help
• Alt+F4 - Exit application"""
        messagebox.showinfo("Keyboard Shortcuts", shortcuts_text)

    def show_about(self):
        """Show about dialog"""
        about_text = """🌊 DataFlow Pro v2.0
“Automate Your Data Entry, Amplify Your Productivity”

The ultimate solution for automated data entry and workflow optimization.
Replace repetitive manual tasks with intelligent automation.

✨ Key Features:
• 🌐 Multi-language support (7 languages)
• 🌙 Light/Dark theme
• 🔄 Excel data integration & mapping
• 🎯 Smart coordinate capture
• ⌨️ Keyboard shortcuts
• 💾 Preset save/load system
• 🚀 Batch processing capabilities

DataFlow Pro - Where Data Flows Effortlessly™

Developed with ❤️ using Python
DataFlow Solutions"""
        messagebox.showinfo("About", about_text)


class ColumnMappingDialog:
    def __init__(self, parent, excel_columns, excel_data, lang="en"):
        self.result = None
        self.lang = lang

        self.dialog = tk.Toplevel(parent)
        self.dialog.title(get_text("add_column_mapping", lang))
        self.dialog.geometry("400x200")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Center the dialog
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))

        # Variable name
        ttk.Label(self.dialog, text=get_text("variable_name", self.lang)).grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.var_name = tk.StringVar()
        ttk.Entry(self.dialog, textvariable=self.var_name, width=20).grid(row=0, column=1, padx=10, pady=10)

        # Excel column
        ttk.Label(self.dialog, text=get_text("excel_column", self.lang)).grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        self.excel_col = ttk.Combobox(self.dialog, values=excel_columns, state="readonly", width=18)
        self.excel_col.grid(row=1, column=1, padx=10, pady=10)

        # Preview
        ttk.Label(self.dialog, text=get_text("preview", self.lang)).grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        self.preview_label = ttk.Label(self.dialog, text=get_text("select_column_preview", self.lang))
        self.preview_label.grid(row=2, column=1, padx=10, pady=10, sticky=tk.W)

        self.excel_col.bind('<<ComboboxSelected>>', lambda e: self.update_preview(excel_data))

        # Buttons
        button_frame = ttk.Frame(self.dialog)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        ttk.Button(button_frame, text=get_text("ok", self.lang), command=self.ok_clicked).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text=get_text("cancel", self.lang), command=self.cancel_clicked).pack(side=tk.LEFT, padx=5)

    def update_preview(self, excel_data):
        col = self.excel_col.get()
        if col and not excel_data.empty:
            preview = str(excel_data[col].iloc[0])[:30]
            self.preview_label.config(text=preview)

    def ok_clicked(self):
        if self.var_name.get().strip() and self.excel_col.get():
            self.result = (self.var_name.get().strip(), self.excel_col.get())
            print(f"DEBUG: Dialog result set to {self.result}")  # Debug line
            self.dialog.destroy()
        else:
            messagebox.showwarning(get_text("warning", self.lang), get_text("fill_both_fields", self.lang))

    def cancel_clicked(self):
        self.dialog.destroy()




if __name__ == "__main__":
    app = AutomationGUI()
    app.root.mainloop()