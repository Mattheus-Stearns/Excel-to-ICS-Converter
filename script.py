import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import subprocess
import pandas as pd
from datetime import datetime, timedelta
from icalendar import Calendar, Event
import warnings
warnings.simplefilter("ignore", UserWarning)

# ---- Constants ----
DAY_MAP = {"M":0, "T":1, "W":2, "TH":3, "F":4, "S":5, "SU":6}

# ---- Conversion Logic ----
def convert_xlsx_to_ics(xlsx_path, output_directory):
    print(f"\n--- Converting file: {xlsx_path} ---")
    df = pd.read_excel(xlsx_path, header=2)
    df.columns = df.columns.str.strip()
    print("Columns after stripping:", df.columns.tolist())

    # Map headers
    header_map = {}
    for col in df.columns:
        col_lower = col.lower()
        if "course" in col_lower:
            header_map[col] = "Course Listing"
        elif "pattern" in col_lower or "meeting" in col_lower:
            header_map[col] = "Meeting Patterns"
        elif "start" in col_lower:
            header_map[col] = "Start Date"
        elif "end" in col_lower:
            header_map[col] = "End Date"
    df.rename(columns=header_map, inplace=True)
    print("Columns after renaming:", df.columns.tolist())

    required_cols = ["Course Listing", "Meeting Patterns", "Start Date", "End Date"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {missing}")

    df['Course Listing'] = df['Course Listing'].ffill()
    cal = Calendar()

    for _, row in df.iterrows():
        course = str(row["Course Listing"])
        pattern_text = str(row["Meeting Patterns"])
        start_date = row["Start Date"]
        end_date = row["End Date"]

        if pd.isna(start_date) or pd.isna(end_date) or pd.isna(pattern_text):
            continue

        try:
            start_dt = pd.to_datetime(start_date).date()
            end_dt = pd.to_datetime(end_date).date()
        except (ValueError, TypeError):
            continue

        # Loop through each week until end date
        week_start = start_dt
        while week_start <= end_dt:
            events = create_events_from_pattern(course, pattern_text, week_start, end_dt)
            for ev in events:
                cal.add_component(ev)
            week_start += timedelta(days=7)

    # Output ICS file
    base_name = os.path.splitext(os.path.basename(xlsx_path))[0]
    output_path = os.path.join(output_directory, f"{base_name}.ics")
    with open(output_path, "wb") as f:
        f.write(cal.to_ical())

    print(f"Conversion successful: {output_path}")
    return output_path

def create_events_from_pattern(course, pattern_text, week_start_date, semester_end_date):
    events = []
    for line in pattern_text.splitlines():
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 2:
            continue
        days, times = parts[0], parts[1]
        location = parts[2] if len(parts) > 2 else ""

        start_time_str, end_time_str = [t.strip() for t in times.split("-")]

        temp_days = days.replace("TH","H")
        for day_char in temp_days:
            day_char = "TH" if day_char=="H" else day_char
            weekday = DAY_MAP.get(day_char.upper())
            if weekday is None:
                continue

            event_date = week_start_date + timedelta(days=(weekday - week_start_date.weekday()) % 7)
            # Skip any event beyond semester end
            if event_date > semester_end_date:
                continue

            start_dt = datetime.combine(event_date, datetime.strptime(start_time_str, "%I:%M %p").time())
            end_dt = datetime.combine(event_date, datetime.strptime(end_time_str, "%I:%M %p").time())

            event = Event()
            event.add('summary', course)
            event.add('dtstart', start_dt)
            event.add('dtend', end_dt)
            event.add('description', f"Meeting Pattern: {line} | {location}")
            events.append(event)
    return events

# ---- GUI Functions ----
def refresh_file_list():
    file_listbox.delete(0, tk.END)
    files = sorted(os.listdir(output_dir))
    if not files:
        file_listbox.insert(tk.END, "(no files yet)")
    else:
        for f in files:
            file_listbox.insert(tk.END, f)

def import_files():
    file_paths = filedialog.askopenfilenames(
        title="Select Excel Files to Convert",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_paths:
        converted = []
        for path in file_paths:
            try:
                out_path = convert_xlsx_to_ics(path, output_dir)
                converted.append(os.path.basename(out_path))
            except ValueError as ve:
                messagebox.showerror("Missing Column", str(ve))
            except (pd.errors.ParserError, TypeError, ValueError) as e:
                messagebox.showerror("Error", f"Failed to convert {os.path.basename(path)}:\n{e}")

        if converted:
            messagebox.showinfo("Success", f"Converted {len(converted)} file(s).")
            refresh_file_list()

def open_selected_file(_):
    selection = file_listbox.curselection()
    if not selection:
        return
    filename = file_listbox.get(selection[0])
    if filename == "(no files yet)":
        return
    filepath = os.path.join(output_dir, filename)
    if not os.path.exists(filepath):
        messagebox.showerror("File Not Found", f"File does not exist:\n{filepath}")
        return

    try:
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin":
            subprocess.run(["open", filepath])
        else:
            subprocess.run(["xdg-open", filepath])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file:\n{e}")

# ---- Setup ----
root = tk.Tk()
root.title("Excel → ICS Converter")

output_dir = os.path.join(os.path.expanduser("~"), "Desktop")
os.makedirs(output_dir, exist_ok=True)

tk.Button(root, text="Import Excel File(s)", command=import_files, width=30).grid(
    row=0, column=0, columnspan=2, pady=10
)

tk.Label(root, text="Generated Files (Desktop):").grid(row=1, column=0, columnspan=2, sticky="w", padx=10)
file_listbox = tk.Listbox(root, width=60, height=12)
file_listbox.grid(row=2, column=0, columnspan=2, padx=10, pady=5)
file_listbox.bind("<Double-Button-1>", open_selected_file)

refresh_file_list()
root.mainloop()