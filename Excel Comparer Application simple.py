import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import sqlite3
import os

# --- Persian Language Strings ---
PERSIAN_STRINGS = {
    "app_title": "برنامه مقایسه فایل‌های اکسل",
    "tab1_title": "فایل اکسل ۱",
    "tab2_title": "فایل اکسل ۲",
    "tab3_title": "ردیف‌های مشترک",
    "tab4_title": "ردیف‌های منحصر به فرد (فایل ۱)",
    "tab5_title": "ردیف‌های منحصر به فرد (فایل ۲)",
    "open_file1_btn": "باز کردن فایل ۱",
    "open_file2_btn": "باز کردن فایل ۲",
    "create_output_file_btn": "ایجاد فایل خروجی XLS/X",
    "database_name": "مقایسه_اکسل.db",
    "db_connect_error": "خطا در اتصال به پایگاه داده: ",
    "db_table_created": "جدول '{table_name}' با موفقیت ایجاد شد.",
    "db_data_inserted": "داده‌ها با موفقیت در جدول '{table_name}' درج شدند.",
    "db_error": "خطای پایگاه داده: ",
    "select_file_prompt": "لطفاً یک فایل اکسل انتخاب کنید (xls/xlsx).",
    "file_load_error": "خطا در بارگذاری فایل: ",
    "file_not_selected": "فایلی انتخاب نشد.",
    "no_file_loaded": "لطفاً ابتدا فایل‌های اکسل را بارگذاری کنید.",
    "output_file_created": "فایل خروجی با موفقیت ایجاد شد:\n",
    "output_file_error": "خطا در ایجاد فایل خروجی: ",
    "output_path_required": "مسیر ذخیره‌سازی فایل خروجی الزامی است.",
    "comparison_ready": "مقایسه‌ها آماده هستند.",
    "no_data_for_tab": "هیچ داده‌ای برای نمایش در این بخش وجود ندارد.",
    "rtl_option": "تنظیم خروجی راست به چپ (RTL)",
    "sqlite_data_saved": "داده‌ها در SQLite ذخیره شدند.",
    "loading_file": "در حال بارگذاری فایل...",
    "performing_comparison": "در حال انجام مقایسه...",
    "comparison_error": "خطا در مقایسه",
    "no_data_for_output": "داده‌ای برای ذخیره در فایل خروجی وجود ندارد.",
}

class ExcelComparerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(PERSIAN_STRINGS["app_title"])
        self.root.geometry("1200x800")

        self.df1 = None
        self.df2 = None
        self.common_df = None
        self.unique1_df = None
        self.unique2_df = None
        
        self.df1_filepath = None # To store original file paths for reference
        self.df2_filepath = None

        self.db_conn = None

        # --- Database Connection ---
        self._connect_db()

        # --- UI Setup ---
        self._create_widgets()

    def _connect_db(self):
        """Attempts to connect to the SQLite database."""
        try:
            self.db_conn = sqlite3.connect(PERSIAN_STRINGS["database_name"])
            self.db_conn.execute("PRAGMA encoding = 'UTF-8';")  # Ensure UTF-8 for Persian
            self.db_conn.execute("PRAGMA journal_mode=WAL;")    # Better concurrency if needed
        except sqlite3.Error as e:
            messagebox.showerror(PERSIAN_STRINGS["db_connect_error"], str(e))
            self.db_conn = None

    def _create_widgets(self):
        """Sets up the main Tkinter widgets: notebook, tabs, buttons, treeviews, etc."""
        # Main frame
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # Notebook (Tabs)
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(pady=5, fill="both", expand=True)

        self.tabs = {}
        self.treeviews = {}

        # Tab 1: Excel File 1
        self.tabs["tab1"] = ttk.Frame(self.notebook)
        self.notebook.add(self.tabs["tab1"], text=PERSIAN_STRINGS["tab1_title"])
        self._create_tab_content(self.tabs["tab1"], "tab1", self._open_file1)

        # Tab 2: Excel File 2
        self.tabs["tab2"] = ttk.Frame(self.notebook)
        self.notebook.add(self.tabs["tab2"], text=PERSIAN_STRINGS["tab2_title"])
        self._create_tab_content(self.tabs["tab2"], "tab2", self._open_file2)

        # Tab 3: Common Rows
        self.tabs["tab3"] = ttk.Frame(self.notebook)
        self.notebook.add(self.tabs["tab3"], text=PERSIAN_STRINGS["tab3_title"])
        self._create_comparison_tab_content(self.tabs["tab3"], "common")

        # Tab 4: Unique Rows (File 1)
        self.tabs["tab4"] = ttk.Frame(self.notebook)
        self.notebook.add(self.tabs["tab4"], text=PERSIAN_STRINGS["tab4_title"])
        self._create_comparison_tab_content(self.tabs["tab4"], "unique1")

        # Tab 5: Unique Rows (File 2)
        self.tabs["tab5"] = ttk.Frame(self.notebook)
        self.notebook.add(self.tabs["tab5"], text=PERSIAN_STRINGS["tab5_title"])
        self._create_comparison_tab_content(self.tabs["tab5"], "unique2")
        
        # --- Global Controls ---
        control_frame = ttk.Frame(self.main_frame)
        control_frame.pack(fill="x", pady=5)

        # Create Output File Button
        self.create_file_btn = ttk.Button(control_frame, text=PERSIAN_STRINGS["create_output_file_btn"], 
                                         command=self._create_output_file, style="TButton")
        self.create_file_btn.pack(side="right", padx=5)

        # RTL Option Checkbox
        self.rtl_var = tk.BooleanVar(value=True) # Default to True for Persian context
        self.rtl_checkbox = ttk.Checkbutton(control_frame, text=PERSIAN_STRINGS["rtl_option"], variable=self.rtl_var)
        self.rtl_checkbox.pack(side="right", padx=5)

        # Status Label
        self.status_label = ttk.Label(self.main_frame, text="", anchor="e", style="Status.TLabel")
        self.status_label.pack(fill="x", padx=10, pady=5)

        # Configure styles
        style = ttk.Style()
        style.configure("Treeview", font=("Arial", 10), rowheight=25, direction="rtl")
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"), anchor="e")
        style.configure("TButton", font=("Arial", 10))
        style.configure("Status.TLabel", font=("Arial", 10), anchor="e")
        style.configure("TNotebook", tabposition="ne")  # Tabs at top-right

    def _create_tab_content(self, tab_frame, treeview_key, open_command):
        """Helper to create content for file loading tabs (tab1, tab2)."""
        # Button frame
        btn_frame = ttk.Frame(tab_frame)
        btn_frame.pack(fill="x", pady=5)

        btn = ttk.Button(btn_frame, text=PERSIAN_STRINGS[f"open_file{treeview_key[-1]}_btn"], 
                         command=open_command, style="TButton")
        btn.pack(side="right", padx=5)

        # Treeview with scrollbars
        self.treeviews[treeview_key] = self._create_treeview_with_scrollbar(tab_frame)
        self.treeviews[treeview_key].pack(expand=True, fill="both")

    def _create_comparison_tab_content(self, tab_frame, treeview_key):
        """Helper to create content for comparison result tabs (common, unique1, unique2)."""
        self.treeviews[treeview_key] = self._create_treeview_with_scrollbar(tab_frame)
        self.treeviews[treeview_key].pack(expand=True, fill="both")

    def _create_treeview_with_scrollbar(self, parent_frame):
        """Creates a Treeview with both vertical and horizontal scrollbars, adjusted for RTL."""
        frame = ttk.Frame(parent_frame)
        frame.pack(expand=True, fill="both")

        scrollbar_y = ttk.Scrollbar(frame, orient="vertical")
        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal")
        treeview = ttk.Treeview(frame, yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set, 
                                show="headings", style="Treeview")

        scrollbar_y.config(command=treeview.yview)
        scrollbar_x.config(command=treeview.xview)

        scrollbar_y.pack(side="left", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        treeview.pack(side="right", expand=True, fill="both")  # Align to right for RTL

        return treeview

    def _display_df_in_treeview(self, treeview, df, max_rows=1000000):
        """Displays DataFrame content in a Treeview, including headers, respecting max_rows."""
        # Clear existing items
        for item in treeview.get_children():
            treeview.delete(item)

        if df is None or df.empty:
            treeview["columns"] = ["پیام"]
            treeview.heading("پیام", text="پیام", anchor="e")
            treeview.column("پیام", anchor="e", width=200)
            treeview.insert("", tk.END, values=(PERSIAN_STRINGS["no_data_for_tab"],))
            return

        # Set up columns in reverse order for RTL
        columns = list(df.columns)[::-1]
        treeview["columns"] = columns
        for col in columns:
            treeview.heading(col, text=col, anchor="e")  # Right-align headers
            treeview.column(col, anchor="e", width=150)  # Right-align data, adjustable width

        # Check which columns contain only numeric data
        numeric_columns = []
        for col in df.columns:
            is_numeric = df[col].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x)).all()
            if is_numeric:
                numeric_columns.append(col)

        # Insert rows, limiting to max_rows for performance
        for i, (index, row) in enumerate(df.iterrows()):
            if i >= max_rows:
                treeview.insert("", tk.END, values=(f"... {len(df)-max_rows} ردیف دیگر (نمایش محدود به {max_rows} ردیف)",))
                break
            # Create a list to store the row values
            row_values = []
            for col, val in zip(df.columns, row):
                row_values.append(str(val))
            # Reverse the row values for RTL display
            row_values = row_values[::-1]
            treeview.insert("", tk.END, values=row_values)

    def _open_file(self, file_num):
        """Handles opening an Excel file, loading it into a DataFrame, and displaying it."""
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            title=PERSIAN_STRINGS["select_file_prompt"]
        )
        if not filepath:
            self._update_status(PERSIAN_STRINGS["file_not_selected"])
            return

        self._update_status(PERSIAN_STRINGS["loading_file"] + os.path.basename(filepath))

        try:
            df = pd.read_excel(filepath)
            # Ensure columns are in correct order as in the Excel file
            if file_num == 1:
                self.df1 = df
                self.df1_filepath = filepath
                self._display_df_in_treeview(self.treeviews["tab1"], self.df1)
                self._update_status(f"{PERSIAN_STRINGS['tab1_title']} بارگذاری شد: {os.path.basename(filepath)}")
                self._save_df_to_sqlite(self.df1, "file1_data")
            elif file_num == 2:
                self.df2 = df
                self.df2_filepath = filepath
                self._display_df_in_treeview(self.treeviews["tab2"], self.df2)
                self._update_status(f"{PERSIAN_STRINGS['tab2_title']} بارگذاری شد: {os.path.basename(filepath)}")
                self._save_df_to_sqlite(self.df2, "file2_data")

            # If both files are loaded, perform comparison
            if self.df1 is not None and self.df2 is not None:
                self._perform_comparison()

        except Exception as e:
            messagebox.showerror(PERSIAN_STRINGS["file_load_error"], f"{str(e)}\n فایل: {filepath}")
            self._update_status(f"خطا در بارگذاری فایل: {os.path.basename(filepath)}")

    def _open_file1(self):
        self._open_file(1)

    def _open_file2(self):
        self._open_file(2)

    def _perform_comparison(self):
        """Performs the comparison logic between df1 and df2 to find common and unique rows."""
        if self.df1 is None or self.df2 is None:
            self._update_status(PERSIAN_STRINGS["no_file_loaded"])
            return

        self._update_status(PERSIAN_STRINGS["performing_comparison"])

        try:
            # حفظ ترتیب ستون‌ها از فایل اول به عنوان مرجع
            # اگر فایل دوم ستون‌های متفاوتی دارد، آنها را به انتها اضافه می‌کنیم
            reference_columns = list(self.df1.columns)
            additional_columns = [col for col in self.df2.columns if col not in reference_columns]
            all_columns = reference_columns + additional_columns

            # Reindex both dataframes to have the same columns, filling missing with NaN
            df1_aligned = self.df1.reindex(columns=all_columns)
            df2_aligned = self.df2.reindex(columns=all_columns)

            # Convert all columns to string type to avoid type mismatch issues during merge
            df1_str = df1_aligned.astype(str).fillna('')
            df2_str = df2_aligned.astype(str).fillna('')

            # Common Rows (Inner Join)
            self.common_df = pd.merge(df1_str, df2_str, how='inner', on=all_columns)
            self._display_df_in_treeview(self.treeviews["common"], self.common_df)
            self._save_df_to_sqlite(self.common_df, "common_rows")

            # Unique Rows (Outer join with indicator)
            merged_outer = pd.merge(df1_str, df2_str, how='outer', on=all_columns, indicator=True)

            self.unique1_df = merged_outer[merged_outer['_merge'] == 'left_only'].drop(columns=['_merge'])
            # بازگرداندن ترتیب ستون‌ها به حالت اصلی
            self.unique1_df = self.unique1_df[all_columns]
            self._display_df_in_treeview(self.treeviews["unique1"], self.unique1_df)
            self._save_df_to_sqlite(self.unique1_df, "unique1_rows")

            self.unique2_df = merged_outer[merged_outer['_merge'] == 'right_only'].drop(columns=['_merge'])
            # بازگرداندن ترتیب ستون‌ها به حالت اصلی
            self.unique2_df = self.unique2_df[all_columns]
            self._display_df_in_treeview(self.treeviews["unique2"], self.unique2_df)
            self._save_df_to_sqlite(self.unique2_df, "unique2_rows")

            self._update_status(PERSIAN_STRINGS["comparison_ready"])

        except Exception as e:
            messagebox.showerror(PERSIAN_STRINGS["comparison_error"], f"خطا در انجام مقایسه: {str(e)}")
            self._update_status("خطا در مقایسه فایل‌ها.")
            print(f"Comparison Error: {e}") # For debugging

    def _create_output_file(self):
        """Creates an XLSX file with data from all tabs, supporting RTL."""
        # Check if any data exists to save
        if all(df is None or df.empty for df in [self.df1, self.df2, self.common_df, self.unique1_df, self.unique2_df]):
            messagebox.showwarning("داده‌ای موجود نیست", PERSIAN_STRINGS["no_data_for_output"])
            return

        output_filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="انتخاب مسیر ذخیره فایل خروجی"
        )
        if not output_filepath:
            self._update_status(PERSIAN_STRINGS["output_path_required"])
            return

        try:
            with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
                # Write each DataFrame to a separate sheet if it exists
                if self.df1 is not None and not self.df1.empty:
                    self.df1.to_excel(writer, sheet_name=PERSIAN_STRINGS["tab1_title"], index=False)
                if self.df2 is not None and not self.df2.empty:
                    self.df2.to_excel(writer, sheet_name=PERSIAN_STRINGS["tab2_title"], index=False)
                if self.common_df is not None and not self.common_df.empty:
                    self.common_df.to_excel(writer, sheet_name=PERSIAN_STRINGS["tab3_title"], index=False)
                if self.unique1_df is not None and not self.unique1_df.empty:
                    self.unique1_df.to_excel(writer, sheet_name=PERSIAN_STRINGS["tab4_title"], index=False)
                if self.unique2_df is not None and not self.unique2_df.empty:
                    self.unique2_df.to_excel(writer, sheet_name=PERSIAN_STRINGS["tab5_title"], index=False)

                # Set RTL property for all sheets in the workbook if checked
                if self.rtl_var.get():
                    workbook = writer.book
                    for sheet_name in workbook.sheetnames:
                        sheet = workbook[sheet_name]
                        sheet.sheet_view.rightToLeft = True # Correct openpyxl property for RTL

            self._update_status(PERSIAN_STRINGS["output_file_created"] + output_filepath)
            messagebox.showinfo("موفقیت", PERSIAN_STRINGS["output_file_created"] + output_filepath)

        except Exception as e:
            messagebox.showerror(PERSIAN_STRINGS["output_file_error"], str(e))
            self._update_status(PERSIAN_STRINGS["output_file_error"] + str(e))

    def _save_df_to_sqlite(self, df, table_name):
        """Saves a DataFrame to a SQLite table."""
        if self.db_conn is None:
            return
        if df is None or df.empty:
            return

        try:
            df.to_sql(table_name, self.db_conn, if_exists='replace', index=False)
            self.db_conn.commit()
            self._update_status(PERSIAN_STRINGS["db_data_inserted"].format(table_name=table_name))
        except sqlite3.Error as e:
            messagebox.showerror(PERSIAN_STRINGS["db_error"], f"جدول: {table_name}\nخطا: {e}")
            self._update_status(f"خطا در ذخیره داده در {table_name} در DB.")
            print(f"SQLite Error: {e}") # For debugging

    def _update_status(self, message):
        """Updates the status bar label with a given message."""
        self.status_label.config(text=message)
        self.root.update_idletasks() # Ensure UI updates immediately

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparerApp(root)
    root.mainloop()