import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import time
from datetime import datetime
import os

class ExcelImporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Importer")
        self.root.geometry("1200x800")
        
        self.df = None
        self.filtered_df = None
        self.current_sort_column = None
        self.current_sort_order = "asc"
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # File selection
        ttk.Label(main_frame, text="File Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.file_path_var, width=80).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Load button
        ttk.Button(main_frame, text="Load File", command=self.load_file).grid(row=1, column=0, columnspan=3, pady=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to load file")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=3, column=0, columnspan=3, pady=5)
        
        # Treeview for data display
        self.tree_frame = ttk.Frame(main_frame)
        self.tree_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        self.tree_frame.columnconfigure(0, weight=1)
        self.tree_frame.rowconfigure(0, weight=1)
        
        self.tree = ttk.Treeview(self.tree_frame, show='headings')
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Filter controls
        filter_frame = ttk.Frame(main_frame)
        filter_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(filter_frame, text="Filter by:").grid(row=0, column=0, sticky=tk.W)
        self.filter_column_var = tk.StringVar()
        self.filter_column_combo = ttk.Combobox(filter_frame, textvariable=self.filter_column_var, state='readonly')
        self.filter_column_combo.grid(row=0, column=1, padx=5)
        
        self.filter_value_var = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=self.filter_value_var, width=20).grid(row=0, column=2, padx=5)
        
        ttk.Button(filter_frame, text="Apply Filter", command=self.apply_filter).grid(row=0, column=3, padx=5)
        ttk.Button(filter_frame, text="Clear Filter", command=self.clear_filter).grid(row=0, column=4, padx=5)
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="Sort Ascending", command=lambda: self.sort_data('asc')).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Sort Descending", command=lambda: self.sort_data('desc')).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Validate Data", command=self.validate_data).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="Export to Excel", command=self.export_data).grid(row=0, column=3, padx=5)
        ttk.Button(button_frame, text="Show Statistics", command=self.show_statistics).grid(row=0, column=4, padx=5)
        
        # Bind events
        self.tree.bind('<Button-1>', self.on_tree_click)
        self.filter_value_var.trace('w', self.on_filter_change)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
    
    def load_file(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a file first")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "File does not exist")
            return
        
        self.status_var.set("Loading file...")
        self.progress.start(10)
        
        # Run loading in separate thread
        threading.Thread(target=self._load_file_thread, daemon=True).start()
    
    def _load_file_thread(self):
        try:
            file_path = self.file_path_var.get()
            
            # Read Excel file
            self.df = pd.read_excel(file_path, engine='openpyxl')
            
            # Update UI in main thread
            self.root.after(0, self._file_loaded_successfully)
            
        except Exception as e:
            self.root.after(0, lambda: self._file_load_error(str(e)))
    
    def _file_loaded_successfully(self):
        self.progress.stop()
        self.status_var.set(f"File loaded successfully: {len(self.df)} rows, {len(self.df.columns)} columns")
        
        # Update filter column combo
        self.filter_column_combo['values'] = self.df.columns.tolist()
        if self.df.columns.any():
            self.filter_column_var.set(self.df.columns[0])
        
        # Display data
        self.display_data()
        
        messagebox.showinfo("Success", f"File loaded successfully!\nRows: {len(self.df)}\nColumns: {len(self.df.columns)}")
    
    def _file_load_error(self, error_msg):
        self.progress.stop()
        self.status_var.set("Error loading file")
        messagebox.showerror("Error", f"Failed to load file:\n{error_msg}")
    
    def display_data(self, data=None):
        if data is None:
            data = self.df
        
        # Clear existing tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree['columns'] = []
        
        if data is None or data.empty:
            return
        
        # Set up columns
        columns = data.columns.tolist()
        self.tree['columns'] = columns
        
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(c))
            self.tree.column(col, width=100, minwidth=50)
        
        # Add data (limited to 1000 rows for performance)
        display_data = data.head(1000) if len(data) > 1000 else data
        
        for _, row in display_data.iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else "" for col in columns]
            self.tree.insert("", "end", values=values)
    
    def sort_treeview(self, column):
        if self.df is None:
            return
        
        if self.current_sort_column == column:
            self.current_sort_order = 'desc' if self.current_sort_order == 'asc' else 'asc'
        else:
            self.current_sort_column = column
            self.current_sort_order = 'asc'
        
        self.sort_data(self.current_sort_order, column)
    
    def sort_data(self, order='asc', column=None):
        if self.df is None:
            return
        
        if column is None:
            # Get currently selected column from treeview
            focus = self.tree.focus()
            if not focus:
                messagebox.showinfo("Info", "Please select a column to sort by clicking its header")
                return
            
            # This is simplified - in real implementation, track sort column
            messagebox.showinfo("Info", "Click on column header to sort")
            return
        
        try:
            data_to_sort = self.filtered_df if self.filtered_df is not None else self.df
            
            if column not in data_to_sort.columns:
                return
            
            # Handle different data types for sorting
            if pd.api.types.is_numeric_dtype(data_to_sort[column]):
                sorted_df = data_to_sort.sort_values(by=column, ascending=(order == 'asc'))
            elif pd.api.types.is_datetime64_any_dtype(data_to_sort[column]):
                sorted_df = data_to_sort.sort_values(by=column, ascending=(order == 'asc'))
            else:
                sorted_df = data_to_sort.sort_values(by=column, ascending=(order == 'asc'), key=lambda x: x.astype(str).str.lower())
            
            if self.filtered_df is not None:
                self.filtered_df = sorted_df
            else:
                self.df = sorted_df
            
            self.display_data(sorted_df)
            self.status_var.set(f"Sorted by {column} ({order})")
            
        except Exception as e:
            messagebox.showerror("Error", f"Sorting error: {str(e)}")
    
    def apply_filter(self):
        if self.df is None:
            return
        
        column = self.filter_column_var.get()
        value = self.filter_value_var.get().strip()
        
        if not column or not value:
            messagebox.showinfo("Info", "Please select column and enter filter value")
            return
        
        try:
            if column not in self.df.columns:
                return
            
            # Handle different filter types
            if pd.api.types.is_numeric_dtype(self.df[column]):
                try:
                    num_value = float(value)
                    self.filtered_df = self.df[self.df[column] == num_value]
                except ValueError:
                    messagebox.showerror("Error", "Please enter a valid number for numeric column")
                    return
            
            elif pd.api.types.is_datetime64_any_dtype(self.df[column]):
                try:
                    date_value = pd.to_datetime(value)
                    self.filtered_df = self.df[self.df[column] == date_value]
                except ValueError:
                    # Try partial match for dates
                    self.filtered_df = self.df[self.df[column].astype(str).str.contains(value, case=False, na=False)]
            
            else:
                # Text filter
                self.filtered_df = self.df[self.df[column].astype(str).str.contains(value, case=False, na=False)]
            
            self.display_data(self.filtered_df)
            self.status_var.set(f"Filter applied: {column} contains '{value}' - {len(self.filtered_df)} rows found")
            
        except Exception as e:
            messagebox.showerror("Error", f"Filter error: {str(e)}")
    
    def clear_filter(self):
        self.filtered_df = None
        self.filter_value_var.set("")
        if self.df is not None:
            self.display_data(self.df)
            self.status_var.set("Filter cleared")
    
    def on_filter_change(self, *args):
        # Real-time filtering could be implemented here
        pass
    
    def validate_data(self):
        if self.df is None:
            return
        
        validation_results = []
        
        # Check for missing values
        missing_values = self.df.isnull().sum()
        for col, count in missing_values.items():
            if count > 0:
                validation_results.append(f"{col}: {count} missing values")
        
        # Check data types consistency
        for col in self.df.columns:
            unique_types = self.df[col].apply(type).nunique()
            if unique_types > 1:
                validation_results.append(f"{col}: Mixed data types detected")
        
        # Show results
        if validation_results:
            result_text = "Validation Results:\n\n" + "\n".join(validation_results)
            messagebox.showwarning("Validation Results", result_text)
        else:
            messagebox.showinfo("Validation Results", "No validation issues found!")
    
    def export_data(self):
        if self.df is None:
            return
        
        data_to_export = self.filtered_df if self.filtered_df is not None else self.df
        
        if data_to_export.empty:
            messagebox.showinfo("Info", "No data to export")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            data_to_export.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Data exported successfully to {file_path}")
            self.status_var.set(f"Data exported: {len(data_to_export)} rows")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def show_statistics(self):
        if self.df is None:
            return
        
        stats_window = tk.Toplevel(self.root)
        stats_window.title("Data Statistics")
        stats_window.geometry("600x400")
        
        text_area = ScrolledText(stats_window, wrap=tk.WORD)
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        stats_text = f"Data Statistics\n{'='*50}\n"
        stats_text += f"Total rows: {len(self.df)}\n"
        stats_text += f"Total columns: {len(self.df.columns)}\n\n"
        
        stats_text += f"Column Information:\n{'='*50}\n"
        for col in self.df.columns:
            stats_text += f"{col}: {self.df[col].dtype}\n"
            stats_text += f"  Missing values: {self.df[col].isnull().sum()}\n"
            if pd.api.types.is_numeric_dtype(self.df[col]):
                stats_text += f"  Min: {self.df[col].min():.2f}\n"
                stats_text += f"  Max: {self.df[col].max():.2f}\n"
                stats_text += f"  Mean: {self.df[col].mean():.2f}\n"
            stats_text += "\n"
        
        text_area.insert(tk.END, stats_text)
        text_area.config(state=tk.DISABLED)
    
    def on_tree_click(self, event):
        # Handle treeview clicks for sorting
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            column = self.tree.identify_column(event.x)
            col_index = int(column.replace('#', '')) - 1
            if col_index < len(self.tree['columns']):
                col_name = self.tree['columns'][col_index]
                self.sort_treeview(col_name)

def main():
    root = tk.Tk()
    app = ExcelImporterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()