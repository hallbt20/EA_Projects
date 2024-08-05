import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

selected_reports = []


def find_file(folder_path, file_start):
    for file_name in os.listdir(folder_path):
        if file_name.startswith(file_start):
            return os.path.join(folder_path, file_name)
    return None


def browse_for_file(missing_files):
    for key, value in missing_files.items():
        if not value:
            file_path = filedialog.askopenfilename(title=f"Locate {key}")
            if file_path:
                missing_files[key] = file_path
            else:
                messagebox.showwarning("File not found", f"{key} file not found. Please locate the file.")
    return missing_files


def startup_window():
    def on_browse():
        if advisory_var.get():
            selected_reports.append('Advisory')
        if outsourcing_var.get():
            selected_reports.append('Outsourcing')
        if not selected_reports:
            messagebox.showwarning("Selection Error", "Please select at least one report type.")
            return

        root.destroy()

    def on_closing():
        root.destroy()
        exit(0)

    root = tk.Tk()
    root.title("Advisory Pipeline Automatic Report Generator")

    label = tk.Label(
        root,
        text="This is the Advisory Pipeline Automatic Report Generator.",
        font=('Helvetica', 14, 'bold')
    )
    label.pack(pady=(20, 10))

    instruction = tk.Label(root, text="Please select a folder with the raw data files.", font=('Helvetica', 12))
    instruction.pack(pady=(0, 20))

    advisory_var = tk.BooleanVar()
    outsourcing_var = tk.BooleanVar()

    advisory_check = tk.Checkbutton(root, text="Advisory Report", variable=advisory_var, font=('Helvetica', 12))
    advisory_check.pack(pady=(0, 10))

    outsourcing_check = tk.Checkbutton(root, text="Outsourcing Report", variable=outsourcing_var, font=('Helvetica', 12))
    outsourcing_check.pack(pady=(0, 10))

    browse_button = tk.Button(root, text="Browse", command=on_browse, font=('Helvetica', 12))
    browse_button.pack(pady=(20, 20))

    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()

    if 'Outsourcing' in selected_reports:
        return start_window_out()
    else:
        return start_window_adv()


def start_window_adv():
    folder_path = filedialog.askdirectory(title="Select a folder")
    if not folder_path:
        messagebox.showerror("Error", "No folder selected. The program will now close")
        exit(0)

    files_to_find = {
        "sf_active_file": "Salesforce Active",
        "sf_closed_file": "Salesforce Closed",
        "ns_active_file": "Netsuite Active",
        "ns_closed_file": "Netsuite Closed",
        "great_lakes_file": "EAG GL",
        'originators_list': 'Originators List',
    }

    found_files = {key: find_file(folder_path, start) for key, start in files_to_find.items()}

    missing_files = {key: value for key, value in found_files.items() if value is None}

    if missing_files:
        messagebox.showinfo("Missing Files", "Some files were not found. Please locate the missing files.")
        found_files.update(browse_for_file(missing_files))

    file_path_vars = [
        found_files['sf_active_file'],
        found_files['sf_closed_file'],
        found_files['ns_active_file'],
        found_files['ns_closed_file'],
        found_files['great_lakes_file'],
        found_files['originators_list'],
    ]

    return file_path_vars


def start_window_out():
    folder_path = filedialog.askdirectory(title="Select a folder")
    if not folder_path:
        messagebox.showerror("Error", "No folder selected.")
        return None

    files_to_find = {
        "sf_active_file": "Salesforce Active",
        "sf_closed_file": "Salesforce Closed",
        "ns_active_file": "Netsuite Active",
        "ns_closed_file": "Netsuite Closed",
        "great_lakes_file": "EAG GL",
        'originators_list': 'Originators List',
        "hubspot_file": "hubspot",
        "eag_gc_oit_file": "EAG GC OIT",
        "legacy_oit_file": "Legacy OIT",
        "triangle_file": "EAG Triangle"
    }

    found_files = {key: find_file(folder_path, start) for key, start in files_to_find.items()}

    missing_files = {key: value for key, value in found_files.items() if value is None}

    if missing_files:
        messagebox.showinfo("Missing Files", "Some files were not found. Please locate the missing files.")
        found_files.update(browse_for_file(missing_files))

    file_path_vars = [
        found_files['sf_active_file'],
        found_files['sf_closed_file'],
        found_files['ns_active_file'],
        found_files['ns_closed_file'],
        found_files['great_lakes_file'],
        found_files['originators_list'],
        found_files['hubspot_file'],
        found_files['eag_gc_oit_file'],
        found_files['legacy_oit_file'],
        found_files['triangle_file']
    ]

    return file_path_vars


def prompt_adv_values(unique_originators):
    root = tk.Tk()
    root.title("Fill in Advisory Values")

    entries = []

    def check_values():
        all_selected = all(entry[1].get() != "Select" for entry in entries)
        submit_button.config(state=tk.NORMAL if all_selected else tk.DISABLED)

    def on_submit():
        for entry in entries:
            unique_originators[entry[0]]['Department (Advisory)'] = entry[1].get()
        root.quit()
        root.destroy()

    # Create labels and dropdowns for each unique unmatched originator
    for i, originator in enumerate(unique_originators):
        label = tk.Label(root, text=f"{originator}")
        label.grid(row=i, column=0)
        var = tk.StringVar(root)
        var.set("Select")
        dropdown = ttk.Combobox(
            root,
            textvariable=var,
            values=['Advisory', 'Other']
        )
        dropdown.grid(row=i, column=1)
        dropdown.bind("<<ComboboxSelected>>", lambda event: check_values())
        entries.append((originator, var))

    # Create a submit button, initially disabled
    submit_button = tk.Button(root, text="Submit", command=on_submit, state=tk.DISABLED)
    submit_button.grid(row=len(unique_originators), column=0, columnspan=2)

    # Run the Tkinter main loop
    root.mainloop()

    return unique_originators


def prompt_out_values(unique_originators):
    root = tk.Tk()
    root.title("Fill in Outsource Values")

    entries = []

    def check_values():
        all_selected = all(entry[1].get() != "Select" for entry in entries)
        submit_button.config(state=tk.NORMAL if all_selected else tk.DISABLED)

    def on_submit():
        for entry in entries:
            unique_originators[entry[0]]['Department (Outsourced)'] = entry[1].get()
        root.quit()
        root.destroy()

    # Create labels and dropdowns for each unique unmatched originator
    for i, originator in enumerate(unique_originators):
        label = tk.Label(root, text=f"{originator}")
        label.grid(row=i, column=0)
        var = tk.StringVar(root)
        var.set("Select")
        dropdown = ttk.Combobox(
            root,
            textvariable=var,
            values=['Advisory', 'Other', 'Outsourced', 'Assurance', 'Business Development', 'Tax']
        )
        dropdown.grid(row=i, column=1)
        dropdown.bind("<<ComboboxSelected>>", lambda event: check_values())
        entries.append((originator, var))

    # Create a submit button, initially disabled
    submit_button = tk.Button(root, text="Submit", command=on_submit, state=tk.DISABLED)
    submit_button.grid(row=len(unique_originators), column=0, columnspan=2)

    # Run the Tkinter main loop
    root.mainloop()

    return unique_originators


def show_report_generated_message(file_name, cwd):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    messagebox.showinfo(
        "Report Generated",
        f"Your advisory report has been generated. It is called {file_name}, and it has been saved to {cwd}."
    )
    root.destroy()
