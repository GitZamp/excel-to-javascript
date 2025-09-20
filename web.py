import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import threading

# === Configuration ===
string_columns = (
    'lname', 'fname', 'nfather', 'nmother',
    'id', 'sin', 'zipcode', 'phone_number', 'address'
)

df = None  # Placeholder for loaded DataFrame

# === Functions ===
def threaded(fn):
    """Decorator to run a function in a background thread."""
    def wrapper(*args, **kwargs):
      threading.Thread(target=fn, args=args, kwargs=kwargs, daemon=True).start()
    return wrapper

@threaded
def select_file():
    global df
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],
        title="Select Excel file"
    )
    if not file_path:
        return

    root.after(0, lambda: file_label.config(text=file_path))

    try:
        sheets = pd.ExcelFile(file_path).sheet_names
        root.after(0, lambda: sheet_combo.configure(values=sheets))
        if sheets:
            root.after(0, lambda: sheet_combo.current(0))
            load_sheet(file_path, sheets[0])  # Autoload first sheet
    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"Failed to read Excel file:\n{e}"))

@threaded
def load_sheet(file_path=None, sheet_name=None):
    global df
    if file_path is None:
        file_path = file_label.cget("text")
    if sheet_name is None:
        sheet_name = sheet_combo.get()

    if not file_path or not sheet_name:
        return

    try:
        # Read Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={col: str for col in string_columns})

        # Date handling
        if 'bdate' not in df.columns:
            df['bdate'] = ''
        df['bdate'] = pd.to_datetime(df['bdate'], errors='coerce')
        df['bdate'] = df['bdate'].dt.strftime('%d/%m/%Y')
        df['bdate'] = df['bdate'].fillna('')

        # Normalize column names
        df.columns = df.columns.str.strip().str.lower().str.replace('\ufeff', '', regex=False)

        # Ensure expected columns exist
        for col in string_columns:
            if col not in df.columns:
                df[col] = ''

        root.after(0, populate_js)  # Safe gui update
    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"Failed to load sheet:\n{e}"))

def generate_js(row):
    return (
        f"document.getElementById('**PUT HERE ID FROM THE TEXT FIELD YOU WANT TO AUTOFILL**').value = '{row['**PUT HERE THE STRING COLUMN YOU WANT**']}';\n"
        f"document.getElementById('**PUT HERE ID FROM THE TEXT FIELD YOU WANT TO AUTOFILL**').dispatchEvent(new Event('input', {{ bubbles: true }}));\n" # You need this line for the above one to apply, same goes for other fields two, you need both of them

        f"let el = document.getElementById('PUT HERE ID FROM THE TEXT FIELD YOU WANT TO AUTOFILL');\n" # This codeblock is made for the date, don't need to change much
        f"el.value = '{row['bdate']}';\n"
        f"el.dispatchEvent(new Event('input', {{ bubbles: true }}));\n"
        f"el.dispatchEvent(new Event('change', {{ bubbles: true }}));\n"
        f"el.dispatchEvent(new Event('blur', {{ bubbles: true }}));\n"

        # -------------
        f"document.getElementById('PUT HERE ID FROM THE TEXT FIELD YOU WANT TO AUTOFILL').value = '**PUT HERE THE STATIC VALUE YOU WANT**';\n" # !!! These line differ from the first ones, here you put *static values*, which do not need a variable
        f"document.getElementById('PUT HERE ID FROM THE TEXT FIELD YOU WANT TO AUTOFILL').dispatchEvent(new Event('input', {{ bubbles: true }}));\n" 
    )

def copy_to_clipboard(js_code):
    root.clipboard_clear()
    root.clipboard_append(js_code)
    root.update()

def populate_js():
    # Clear prev info
    for widget in frame.winfo_children():
        widget.destroy()

    # Populate new info
    for index, row in df.iterrows():
        js_code = generate_js(row)
        person_frame = ttk.LabelFrame(frame, text=f"Person {index+1}")
        person_frame.pack(fill="x", padx=10, pady=5)

        info = (
            f"Fullname: {row['fname']} {row['lname']}\n"
            f"Father's Name: {row['nfather']} | Mother's Name: {row['nmother']}\n"
            f"Date: {row['bdate']}\n"
            f"ID Number: {row['id']} | Social Insurance Number: {row['sin']}\n"
            f"Phone Number: {row['phone_number']} | TK: {row['zipcode']}\n"
            f"Address: {row['address']}\n"
        )
        label = tk.Label(person_frame, text=info, justify="left", anchor="w")
        label.pack(side="left", padx=5)

        copy_button = tk.Button(person_frame, text="Copy JS", command=lambda code=js_code: copy_to_clipboard(code))
        copy_button.pack(side="right", padx=5)

    canvas.configure(scrollregion=canvas.bbox("all"))

# === GUI ===
root = tk.Tk()
root.title("Excel to JavaScript GUI")

# Window size and center
window_width, window_height = 900, 700
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Top controls
top_frame = ttk.Frame(root)
top_frame.pack(pady=10)

tk.Button(top_frame, text="Select Excel File", command=select_file).grid(row=0, column=0, padx=5)
file_label = tk.Label(top_frame, text="No file selected", wraplength=400)
file_label.grid(row=0, column=1, padx=5)

tk.Label(top_frame, text="Select Sheet:").grid(row=1, column=0, pady=5)
sheet_combo = ttk.Combobox(top_frame, state="readonly", width=40)
sheet_combo.grid(row=1, column=1, pady=5)
tk.Button(top_frame, text="Load Sheet", command=load_sheet).grid(row=1, column=2, padx=5)

# Scrollable frame for JS outputs
container = ttk.Frame(root)
container.pack(fill="both", expand=True, padx=10, pady=10)

canvas = tk.Canvas(container)
scroll_y = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
frame = ttk.Frame(canvas)
canvas_frame = canvas.create_window((0, 0), window=frame, anchor="nw")
canvas.configure(yscrollcommand=scroll_y.set)

frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# Mousewheel support
def _on_mousewheel(event):
    if event.num == 5 or event.delta == -120:
        canvas.yview_scroll(1, "units")
    if event.num == 4 or event.delta == 120:
        canvas.yview_scroll(-1, "units")

canvas.bind_all("<MouseWheel>", _on_mousewheel)
canvas.bind_all("<Button-4>", _on_mousewheel)
canvas.bind_all("<Button-5>", _on_mousewheel)

canvas.pack(fill="both", expand=True, side="left")
scroll_y.pack(fill="y", side="right")

root.mainloop()