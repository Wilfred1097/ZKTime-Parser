import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import json
import webbrowser
import platform
import subprocess
import base64
from datetime import datetime, time

# Try importing pymysql for MySQL/XAMPP database support
try:
    import pymysql
    MYSQL_AVAILABLE = True
except ImportError:
    MYSQL_AVAILABLE = False

# Global Database Configuration & Persistence File with Encryption
CONFIG_FILE = "db_config.json"
ENCRYPTION_KEY = 0x5A  # Simple byte key for lightweight built-in encryption

DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": ""
}

def encrypt_text(text):
    """Encrypts plain text into an obfuscated encrypted string."""
    try:
        byte_arr = text.encode('utf-8')
        xored = bytes([b ^ ENCRYPTION_KEY for b in byte_arr])
        return base64.b64encode(xored).decode('utf-8')
    except Exception:
        return text

def decrypt_text(encrypted_text):
    """Decrypts the encrypted string back to plain text."""
    try:
        decoded_bytes = base64.b64decode(encrypted_text.encode('utf-8'))
        xored = bytes([b ^ ENCRYPTION_KEY for b in decoded_bytes])
        return xored.decode('utf-8')
    except Exception:
        return encrypted_text

def load_db_config():
    """Loads and decrypts saved database credentials from local file if it exists."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                content = f.read().strip()
                if content:
                    decrypted_json = decrypt_text(content)
                    data = json.loads(decrypted_json)
                    for key in DB_CONFIG:
                        if key in data:
                            DB_CONFIG[key] = data[key]
        except Exception as e:
            print(f"Failed to load config file: {e}")

# Load configuration at startup
load_db_config()

def save_db_config():
    """Encrypts and saves current database credentials to local file."""
    try:
        json_str = json.dumps(DB_CONFIG)
        encrypted_content = encrypt_text(json_str)
        with open(CONFIG_FILE, "w") as f:
            f.write(encrypted_content)
    except Exception as e:
        print(f"Failed to save config file: {e}")

def open_file(file_path):
    """Cross-platform helper to open a saved file with its default application."""
    try:
        if platform.system() == 'Darwin':       # macOS
            subprocess.call(('open', file_path))
        elif platform.system() == 'Windows':    # Windows
            os.startfile(file_path)
        else:                                   # Linux variants
            subprocess.call(('xdg-open', file_path))
    except Exception as e:
        print(f"Could not open file: {e}")

def fetch_employee_metadata_from_mysql():
    """Connects to XAMPP MySQL and returns a DataFrame of employee details."""
    if not MYSQL_AVAILABLE:
        return pd.DataFrame()
    try:
        connection = pymysql.connect(
            host=DB_CONFIG["host"],
            port=int(DB_CONFIG["port"]),
            user=DB_CONFIG["user"],
            password=DB_CONFIG["password"],
            database=DB_CONFIG["database"],
            cursorclass=pymysql.cursors.DictCursor
        )
        
        with connection.cursor() as cursor:
            query = """
                SELECT
                    e.user_id,
                    e.employee_id,
                    e.employee_name,
                    s.section_name,
                    p.position_name
                FROM
                    `employees` e
                JOIN sections s ON s.id = e.section_id
                JOIN positions p ON p.id = e.position_id
            """
            cursor.execute(query)
            result = cursor.fetchall()
            
        connection.close()
        return pd.DataFrame(result)
    except Exception as e:
        print(f"Database connection failed: {e}")
        return pd.DataFrame()

def load_attendance_data(file_path):
    # 1. Try reading as standard Excel (.xls via xlrd)
    try:
        df = pd.read_excel(file_path, engine='xlrd')
        if not df.empty and len(df.columns) >= 4:
            return df
    except Exception:
        pass
    
    # 2. Try reading as modern Excel (.xlsx via openpyxl)
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        if not df.empty and len(df.columns) >= 4:
            return df
    except Exception:
        pass

    # 3. Try reading as HTML table (Disguised .xls exports from biometric devices)
    try:
        dfs = pd.read_html(file_path)
        for d in dfs:
            if len(d.columns) >= 4:
                return d
    except Exception:
        pass

    # 4. Fallback to Tab-Separated or CSV text file parsing
    try:
        for sep in ['\t', ',', ';']:
            df = pd.read_csv(file_path, sep=sep, header=None, engine='python')
            if len(df.columns) >= 4:
                df = df.dropna(how='all')
                first_row_str = str(df.iloc[0].values).lower()
                if 'name' in first_row_str or 'date' in first_row_str:
                    df.columns = df.iloc[0]
                    df = df.drop(0).reset_index(drop=True)
                else:
                    df.columns = ['Name', 'No.', 'Date/Time', 'Status'][:len(df.columns)]
                return df
    except Exception:
        pass

    raise ValueError("Could not parse file. Ensure it is a valid attendance export file.")

def generate_processed_dataframe(input_path):
    df = load_attendance_data(input_path)
    
    df.columns = [str(c).strip() for c in df.columns]
    
    name_col = next((c for c in df.columns if 'name' in c.lower()), df.columns[0])
    no_col = next((c for c in df.columns if 'no' in c.lower() or 'id' in c.lower()), df.columns[1])
    dt_col = next((c for c in df.columns if 'date' in c.lower() or 'time' in c.lower()), df.columns[2])
    status_col = next((c for c in df.columns if 'status' in c.lower() or 'in/out' in c.lower()), df.columns[3])

    df['CleanDateTime'] = pd.to_datetime(df[dt_col].astype(str).str.strip(), format='%m/%d/%Y %I:%M:%S %p', errors='coerce')
    
    mask_nat = df['CleanDateTime'].isna()
    if mask_nat.any():
        df.loc[mask_nat, 'CleanDateTime'] = pd.to_datetime(df.loc[mask_nat, dt_col], errors='coerce')

    df = df.dropna(subset=['CleanDateTime']).sort_values('CleanDateTime').reset_index(drop=True)
    
    if df.empty:
        raise ValueError("No valid date/time records found in the file.")

    min_date = df['CleanDateTime'].dt.date.min()
    max_date = df['CleanDateTime'].dt.date.max()
    full_date_range = pd.date_range(start=min_date, end=max_date, freq='D')
    
    processed_records = []
    grouped = list(df.groupby([name_col, no_col]))
    
    for idx, ((name, emp_no), group) in enumerate(grouped):
        group = group.sort_values('CleanDateTime').reset_index(drop=True)
        
        shifts = {}
        i = 0
        n = len(group)
        
        while i < n:
            curr = group.iloc[i]
            status_val = str(curr[status_col]).strip()
            
            is_in = 'in' in status_val.lower()
            
            if is_in:
                in_time = curr['CleanDateTime']
                out_time = None
                
                if i + 1 < n:
                    next_status = str(group.iloc[i + 1][status_col]).strip()
                    if 'out' in next_status.lower():
                        out_time = group.iloc[i + 1]['CleanDateTime']
                        i += 2
                    else:
                        i += 1
                else:
                    i += 1
                
                expected_in = in_time.replace(hour=8, minute=0, second=0, microsecond=0)
                if in_time > expected_in:
                    late_diff = int((in_time - expected_in).total_seconds() // 60)
                    if late_diff > 60:
                        late_str = "-"
                    else:
                        late_str = str(late_diff)
                else:
                    late_str = "0"
                
                worked_str = "-"
                if pd.notnull(out_time):
                    diff_sec = (out_time - in_time).total_seconds()
                    hrs = int(diff_sec // 3600)
                    mins = int((diff_sec % 3600) // 60)
                    
                    if hrs >= 5:
                        hrs -= 1  # 1-hour lunch break deduction
                    worked_str = f"{hrs}:{mins:02d}"
                
                shift_date = in_time.date()
                shifts[shift_date] = [
                    str(name), 
                    str(emp_no), 
                    in_time.strftime('%Y-%m-%d (%a)'),
                    in_time.strftime('%I:%M:%S %p'),
                    out_time.strftime('%I:%M:%S %p') if pd.notnull(out_time) else 'MISSING OUT',
                    worked_str,
                    late_str
                ]
            else:
                shift_date = curr['CleanDateTime'].date()
                shifts[shift_date] = [
                    str(name), 
                    str(emp_no), 
                    curr['CleanDateTime'].strftime('%Y-%m-%d (%a)'),
                    'MISSING IN', 
                    curr['CleanDateTime'].strftime('%I:%M:%S %p'), 
                    '-',
                    '-'
                ]
                i += 1

        for single_date in full_date_range:
            d = single_date.date()
            if d in shifts:
                processed_records.append(shifts[d])
            else:
                processed_records.append([
                    str(name),
                    str(emp_no),
                    single_date.strftime('%Y-%m-%d (%a)'),
                    'NO DUTY',
                    'NO DUTY',
                    '-',
                    '-'
                ])
        
        if idx < len(grouped) - 1:
            processed_records.append(["", "", "", "", "", "", ""])
            
    headers = ["Name", "Emp ID", "Shift Date", "Time In", "Time Out", "Worked Hours", "Late (Mins)"]
    df_output = pd.DataFrame(processed_records, columns=headers)
    
    # Clean and format 'Emp ID' to safely remove any trailing '.0' from excel parsing
    df_output['Emp ID'] = df_output['Emp ID'].astype(str).str.strip()
    df_output['Emp ID'] = df_output['Emp ID'].apply(lambda x: x[:-2] if x.endswith('.0') else x)

    # Merge MySQL Employee Metadata matching explicitly against 'user_id'
    df_emp_meta = fetch_employee_metadata_from_mysql()
    if not df_emp_meta.empty:
        df_emp_meta.columns = [c.lower() for c in df_emp_meta.columns]
        
        if 'user_id' in df_emp_meta.columns:
            df_emp_meta['user_id'] = df_emp_meta['user_id'].astype(str).str.strip()
            df_emp_meta['user_id'] = df_emp_meta['user_id'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
            
            df_output = df_output.merge(
                df_emp_meta, 
                left_on='Emp ID', 
                right_on='user_id', 
                how='left'
            )
            
            # Override 'Name' with database 'employee_name' if available
            if 'employee_name' in df_output.columns:
                mask = df_output['employee_name'].notnull() & (df_output['Emp ID'] != "") & (df_output['Emp ID'] != "nan")
                df_output.loc[mask, 'Name'] = df_output.loc[mask, 'employee_name']
                df_output = df_output.drop(columns=['employee_name'])
                
    # Combine section_name and position_name into "Section - Position"
    def combine_sec_pos(row):
        sec = str(row.get('section_name', '')).strip()
        pos = str(row.get('position_name', '')).strip()
        if sec and pos and sec.lower() != 'nan' and pos.lower() != 'nan':
            return f"{sec} - {pos}"
        elif sec and sec.lower() != 'nan':
            return sec
        elif pos and pos.lower() != 'nan':
            return pos
        return "-"

    df_output['Section - Position'] = df_output.apply(combine_sec_pos, axis=1)

    # Keep only the requested columns and drop Emp ID, section_name, position_name, user_id
    final_headers = ["Name", "Section - Position", "Shift Date", "Time In", "Time Out", "Worked Hours", "Late (Mins)"]
    for h in final_headers:
        if h not in df_output.columns:
            df_output[h] = "-"
            
    df_output = df_output[final_headers]
                
    return df_output

def process_attendance_file_to_excel(input_path, output_path):
    df_output = generate_processed_dataframe(input_path)
    df_output.to_excel(output_path, index=False, engine='openpyxl')
    return f"Successfully processed and saved to:\n{output_path}"

def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select Attendance File",
        filetypes=[("Excel & HTML Files", "*.xls *.xlsx *.html *.htm"), ("All Files", "*.*")]
    )
    if file_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, file_path)
        lbl_status.config(text="File loaded successfully.", fg="#2e7d32")

def center_window(window, width, height):
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

def open_db_settings():
    if not MYSQL_AVAILABLE:
        messagebox.showerror("Library Missing", "pymysql library is not installed.\nRun: pip install pymysql")
        return

    settings_win = tk.Toplevel(root)
    settings_win.title("MySQL Database Settings")
    center_window(settings_win, 420, 320)
    settings_win.configure(bg=BG_COLOR)
    settings_win.transient(root)
    settings_win.grab_set()
    
    try:
        settings_win.iconbitmap("icon/calendar.ico")
    except Exception:
        pass

    frame = tk.Frame(settings_win, bg=BG_COLOR, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    fields = [
        ("Host:", "host"),
        ("Port:", "port"),
        ("Database Name:", "database"),
        ("Username:", "user"),
        ("Password:", "password")
    ]

    entries = {}
    for i, (label_text, key) in enumerate(fields):
        lbl = tk.Label(frame, text=label_text, font=("Segoe UI", 9, "bold"), bg=BG_COLOR, fg=TEXT_DARK)
        lbl.grid(row=i, column=0, sticky="w", pady=5)
        
        # Mask password and database credentials/sensitive info with asterisks if desired, or password field specifically
        entry = tk.Entry(frame, font=("Segoe UI", 10), relief="solid", bd=1)
        if key in ("password", "user", "database"):  # Masking sensitive configuration fields with asterisks
            entry.config(show="*")
        entry.insert(0, str(DB_CONFIG[key]))
        entry.grid(row=i, column=1, sticky="ew", pady=5, padx=(10, 0))
        entries[key] = entry

    frame.columnconfigure(1, weight=1)

    # Set cursor focus to Host field upon pop up
    entries["host"].focus_set()

    def save_settings():
        DB_CONFIG["host"] = entries["host"].get().strip()
        try:
            DB_CONFIG["port"] = int(entries["port"].get().strip())
        except ValueError:
            DB_CONFIG["port"] = 3306
        DB_CONFIG["database"] = entries["database"].get().strip()
        DB_CONFIG["user"] = entries["user"].get().strip()
        DB_CONFIG["password"] = entries["password"].get()
        
        try:
            conn = pymysql.connect(
                host=DB_CONFIG["host"],
                port=DB_CONFIG["port"],
                user=DB_CONFIG["user"],
                password=DB_CONFIG["password"],
                database=DB_CONFIG["database"]
            )
            conn.close()
            
            # Save and encrypt settings persistently locally
            save_db_config()
            
            messagebox.showinfo("Success", "Database connection successful and credentials saved securely!")
            lbl_status.config(text="Database connected and saved securely.", fg="#2e7d32")
            settings_win.destroy()
        except Exception as ex:
            messagebox.showerror("Connection Error", f"Could not connect to database:\n{str(ex)}")

    btn_save = tk.Button(frame, text="Test & Save Settings", font=("Segoe UI", 9, "bold"), bg=PRIMARY_COLOR, fg="white", relief="flat", pady=6, cursor="hand2", command=save_settings)
    btn_save.grid(row=len(fields), column=0, columnspan=2, sticky="ew", pady=(15, 0))

def on_preview():
    input_file = entry_path.get()
    if not input_file:
        messagebox.showerror("Error", "Please select an attendance file first.")
        return
        
    try:
        root.config(cursor="watch")
        root.update()
        df = generate_processed_dataframe(input_file)
        root.config(cursor="")
        
        preview_win = tk.Toplevel(root)
        preview_win.title("Attendance Data Preview")
        center_window(preview_win, 950, 520)
        preview_win.configure(bg=BG_COLOR)
        preview_win.transient(root)
        
        try:
            preview_win.iconbitmap("calendar.ico")
        except Exception:
            pass
        
        search_frame = tk.Frame(preview_win, bg=BG_COLOR)
        search_frame.pack(fill=tk.X, padx=16, pady=(16, 8))

        lbl_search = tk.Label(search_frame, text="Search:", font=("Segoe UI", 9, "bold"), bg=BG_COLOR, fg=TEXT_DARK)
        lbl_search.pack(side=tk.LEFT, padx=(0, 8))

        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Segoe UI", 10), relief="solid", bd=1)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)

        table_frame = tk.Frame(preview_win, bg=BG_COLOR)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        columns = list(df.columns)
        tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        # Scrollbars (Vertical & Horizontal)
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Layout Treeview and Scrollbars using Grid
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=130, anchor="center")

        def populate_tree(data_frame):
            for item in tree.get_children():
                tree.delete(item)
            for _, row in data_frame.iterrows():
                vals = [str(val) if pd.notnull(val) else "" for val in row.values]
                tree.insert("", "end", values=vals)

        populate_tree(df)

        def filter_data(*args):
            query = search_var.get().lower().strip()
            if not query:
                filtered_df = df
            else:
                mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)
                filtered_df = df[mask]
            populate_tree(filtered_df)
            lbl_preview_info.config(text=f"Total Rows: {len(filtered_df)}")

        search_var.trace_add("write", filter_data)

        bottom_frame = tk.Frame(preview_win, bg=BG_COLOR)
        bottom_frame.pack(fill=tk.X, padx=16, pady=(0, 16))

        def save_from_preview():
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
                title="Save Processed Attendance As"
            )
            if output_file:
                try:
                    df.to_excel(output_file, index=False, engine='openpyxl')
                    preview_win.destroy()
                    
                    # Prompt user with an Open Now option
                    if messagebox.askyesno("Success", f"Successfully saved to:\n{output_file}\n\nDo you want to open the file now?"):
                        open_file(output_file)
                except Exception as ex:
                    messagebox.showerror("Error", f"Failed to save file:\n{str(ex)}")

        btn_save_preview = tk.Button(bottom_frame, text="Save to Excel", font=("Segoe UI", 9, "bold"), bg=PRIMARY_COLOR, fg="white", relief="flat", padx=16, pady=6, cursor="hand2", command=save_from_preview)
        btn_save_preview.pack(side=tk.RIGHT)

        lbl_preview_info = tk.Label(bottom_frame, text=f"Total Rows: {len(df)}", font=("Segoe UI", 9), bg=BG_COLOR, fg="#6c757d")
        lbl_preview_info.pack(side=tk.LEFT, pady=6)

        lbl_status.config(text="Preview generated successfully.", fg="#2e7d32")

    except Exception as e:
        root.config(cursor="")
        lbl_status.config(text="Preview generation failed.", fg="#c62828")
        messagebox.showerror("Preview Error", f"An error occurred:\n{str(e)}")

def on_process():
    input_file = entry_path.get()
    if not input_file:
        messagebox.showerror("Error", "Please select an attendance file first.")
        return
    
    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
        title="Save Processed Attendance As"
    )
    
    if not output_file:
        return 
        
    try:
        root.config(cursor="watch")
        root.update()
        result_msg = process_attendance_file_to_excel(input_file, output_file)
        root.config(cursor="")
        lbl_status.config(text="Processing completed successfully!", fg="#2e7d32")
        
        # Prompt user with an Open Now option
        if messagebox.askyesno("Success", f"{result_msg}\n\nDo you want to open the file now?"):
            open_file(output_file)
            
    except Exception as e:
        root.config(cursor="")
        lbl_status.config(text="Processing failed.", fg="#c62828")
        messagebox.showerror("Processing Error", f"An error occurred:\n{str(e)}")

# Build Modern GUI Window
root = tk.Tk()
root.title("Biometric Attendance Processor")
center_window(root, 560, 310)
root.resizable(False, False)

try:
    root.iconbitmap("calendar.ico")
except Exception:
    pass

BG_COLOR = "#f8f9fa"
PRIMARY_COLOR = "#1976d2"
PRIMARY_HOVER = "#115293"
SECONDARY_COLOR = "#6c757d"
SECONDARY_HOVER = "#5a6268"
TEXT_DARK = "#212529"
BORDER_COLOR = "#ced4da"

root.configure(bg=BG_COLOR)

style = ttk.Style()
style.theme_use("clam")

main_frame = tk.Frame(root, bg=BG_COLOR, padx=24, pady=20)
main_frame.pack(fill=tk.BOTH, expand=True)

# Header Title & DB Settings Button Frame
header_frame = tk.Frame(main_frame, bg=BG_COLOR)
header_frame.pack(fill=tk.X, pady=(0, 2))

lbl_title = tk.Label(header_frame, text="Attendance Log Converter", font=("Segoe UI", 14, "bold"), bg=BG_COLOR, fg=TEXT_DARK)
lbl_title.pack(side=tk.LEFT)

btn_db_settings = tk.Button(header_frame, text="⚙️ DB Settings", font=("Segoe UI", 8, "bold"), bg="#e9ecef", fg=TEXT_DARK, relief="flat", padx=8, pady=2, cursor="hand2", command=open_db_settings)
btn_db_settings.pack(side=tk.RIGHT)

lbl_subtitle = tk.Label(main_frame, text="Convert raw biometric machine exports and merge with XAMPP MySQL employee data.", font=("Segoe UI", 9), bg=BG_COLOR, fg="#6c757d")
lbl_subtitle.pack(anchor="w", pady=(0, 16))

lbl_file = tk.Label(main_frame, text="Source Attendance File:", font=("Segoe UI", 9, "bold"), bg=BG_COLOR, fg=TEXT_DARK)
lbl_file.pack(anchor="w", pady=(0, 4))

path_frame = tk.Frame(main_frame, bg=BG_COLOR)
path_frame.pack(fill=tk.X, pady=(0, 12))

entry_path = tk.Entry(path_frame, font=("Segoe UI", 10), relief="solid", bd=1)
entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 8))

btn_browse = tk.Button(path_frame, text="Browse...", font=("Segoe UI", 9, "bold"), bg="#e9ecef", fg=TEXT_DARK, relief="flat", padx=12, pady=4, cursor="hand2", command=browse_file)
btn_browse.pack(side=tk.RIGHT)

action_frame = tk.Frame(main_frame, bg=BG_COLOR)
action_frame.pack(fill=tk.X, pady=(4, 8))

btn_preview = tk.Button(action_frame, text="Preview Data", font=("Segoe UI", 10, "bold"), bg=SECONDARY_COLOR, fg="white", relief="flat", pady=8, cursor="hand2", command=on_preview)
btn_preview.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

btn_process = tk.Button(action_frame, text="Process & Save", font=("Segoe UI", 10, "bold"), bg=PRIMARY_COLOR, fg="white", relief="flat", pady=8, cursor="hand2", command=on_process)
btn_process.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(4, 0))

btn_preview.bind("<Enter>", lambda e: btn_preview.config(bg=SECONDARY_HOVER))
btn_preview.bind("<Leave>", lambda e: btn_preview.config(bg=SECONDARY_COLOR))
btn_process.bind("<Enter>", lambda e: btn_process.config(bg=PRIMARY_HOVER))
btn_process.bind("<Leave>", lambda e: btn_process.config(bg=PRIMARY_COLOR))

bottom_status_frame = tk.Frame(main_frame, bg=BG_COLOR)
bottom_status_frame.pack(fill=tk.X, pady=(2, 0))

lbl_status = tk.Label(bottom_status_frame, text="Ready", font=("Segoe UI", 9), bg=BG_COLOR, fg="#6c757d")
lbl_status.pack(side=tk.LEFT)

def open_facebook(event):
    webbrowser.open("https://www.facebook.com/wil.freeed")

lbl_credit = tk.Label(bottom_status_frame, text="Created By: Wil Fred", font=("Segoe UI", 8), bg=BG_COLOR, fg=PRIMARY_COLOR, cursor="hand2")
lbl_credit.pack(side=tk.RIGHT)
lbl_credit.bind("<Button-1>", open_facebook)

root.mainloop()