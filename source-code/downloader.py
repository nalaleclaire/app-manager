import os
import sys
import ctypes
import subprocess
import json
import tkinter as tk
from tkinter import messagebox, ttk

# Function to check if running with admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        messagebox.showerror("Admin Check Failed", f"Failed to check admin status: {e}")
        return False

# Function to run the script with admin privileges
def run_as_admin():
    if not is_admin():
        try:
            params = " ".join([f'"{param}"' for param in sys.argv])
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, params, None, 1)
            sys.exit()
        except Exception as e:
            messagebox.showerror("Admin Rights Required", f"This program requires admin rights to run: {e}")
            sys.exit()

# Ensure the script runs with admin privileges
run_as_admin()

# Determine the directory path of the current script
current_directory = os.path.dirname(os.path.abspath(__file__))

# Construct the path to applications.json relative to the current directory
file_path = os.path.join(current_directory, 'applications.json')

try:
    with open(file_path, encoding='utf-8') as f:
        applications = json.load(f)
except FileNotFoundError:
    messagebox.showerror("File Not Found", f"Failed to load {file_path}.")
    sys.exit(1)
except json.JSONDecodeError:
    messagebox.showerror("JSON Error", f"{file_path} is not valid JSON.")
    sys.exit(1)
except Exception as e:
    messagebox.showerror("File Error", f"An error occurred while loading {file_path}: {e}")
    sys.exit(1)

# Function to install or update programs
def install_or_update_programs(action):
    selected_programs = [app for app, var in check_vars.items() if var.get()]
    
    if not selected_programs:
        messagebox.showwarning("No Programs Selected", "Please select programs to perform the action.")
        return
    
    progress_bar['maximum'] = len(selected_programs)
    progress_bar['value'] = 0
    progress_label.config(text=f"{action.capitalize()} in progress...")

    for app_name in selected_programs:
        if action == 'install':
            if 'winget' in applications[app_name]:
                command = f"winget install {applications[app_name]['winget']} -h"
            else:
                continue  # Handle other package managers if needed
        elif action == 'uninstall':
            if 'winget' in applications[app_name]:
                confirm = messagebox.askyesno(
                    "Confirm Uninstall",
                    f"Are you sure you want to uninstall {applications[app_name].get('content', app_name)}?"
                )
                if not confirm:
                    continue
                command = f"winget uninstall --id {applications[app_name]['winget']} -h"
            else:
                continue
        elif action == 'update':
            # Add update logic if needed
            continue
        
        try:
            result = subprocess.run(command, shell=True, capture_output=True, text=True, check=True)
            output = result.stdout.strip()

            if action == 'uninstall' and "No installed package found matching input criteria." in output:
                messagebox.showinfo("Action Complete", f"No installed package found matching input criteria for {app_name}.")
            else:
                messagebox.showinfo("Action Complete", f"Successfully {action}ed {applications[app_name].get('content', app_name)}.")

            # Run the program after installation
            if action == 'install' and 'executable' in applications[app_name]:
                try:
                    subprocess.Popen(applications[app_name]['executable'])
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to run {app_name}: {e}")

        except subprocess.CalledProcessError as e:
            if e.returncode == 1603:
                messagebox.showerror(f"{action.capitalize()} Failed", f"{action.capitalize()} failed with exit code: 1603. Maybe try running as administrator.")
            else:
                messagebox.showerror(f"{action.capitalize()} Error", f"{action.capitalize()} failed with error code: {e.returncode}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during {action}: {e}")

        progress_bar['value'] += 1
        root.update_idletasks()

    progress_label.config(text=f"{action.capitalize()} complete.")
    messagebox.showinfo("Action Complete", f"Selected programs have been {action}ed.")

# Function to show the description of a program
def show_description(app_name):
    description_text.set(applications.get(app_name, {}).get('description', "Description not found."))

# Function to check installed programs
def check_installed_programs():
    try:
        installed_apps = subprocess.run('winget list', capture_output=True, text=True, check=True).stdout
    except subprocess.CalledProcessError:
        messagebox.showerror("Error", "Failed to retrieve installed programs.")
        return
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while checking installed programs: {e}")
        return

    installed_apps = installed_apps.splitlines()
    
    for app_name, app_info in applications.items():
        if app_info.get('winget') and any(app_info['winget'] in line for line in installed_apps):
            check_vars[app_name].set(True)

# Function to unselect all programs
def unselect_all_programs():
    for var in check_vars.values():
        var.set(False)

# GUI setup
root = tk.Tk()
root.title("Program Manager")
root.geometry("1200x800")
root.config(bg="#FFFFFF")

# Theme settings
theme = {
    "font": ("Arial", 10),
    "heading_font": ("Consolas", 14),
    "bg": "#FFFFFF",
    "fg": "#333333",
    "button_bg": "#F5F5F5",
    "button_fg": "#000000",
    "listbox_bg": "#FFFFFF",
    "listbox_fg": "#000000"
}

# Header
header = tk.Label(root, text="Program Manager", font=theme["heading_font"], bg=theme["bg"], fg=theme["fg"])
header.pack(pady=10)

# Main frame
main_frame = tk.Frame(root, bg=theme["bg"])
main_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

# Frame for categories and programs with scrollbar
category_frame = tk.Frame(main_frame, bg=theme["bg"])
category_frame.pack(side=tk.LEFT, fill=tk.Y)

program_frame = tk.Frame(main_frame, bg=theme["bg"])
program_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Create a canvas and scrollbar for programs
canvas = tk.Canvas(program_frame, bg=theme["bg"])
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(program_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.configure(yscrollcommand=scrollbar.set)

# Frame to hold the checkboxes for programs
frame = tk.Frame(canvas, bg=theme["bg"])
canvas.create_window((0, 0), window=frame, anchor='nw')

# Dictionary to hold checkboxes for programs
check_vars = {}

# Populate categories and checkboxes
categories = {}

# Initialize row for grid layout
row = 0

for app_name, app_info in applications.items():
    category = app_info.get('category', 'Uncategorized')
    
    if category not in categories:
        categories[category] = []
    
    categories[category].append((app_name, app_info))

# Display categories with section headings
for category, apps in categories.items():
    category_label = tk.Label(frame, text=category, font=("Arial", 12, 'bold'), bg=theme["bg"], fg=theme["fg"])
    category_label.grid(row=row, column=0, columnspan=2, sticky='w', padx=5, pady=(10, 5))
    row += 1
    
    for app_name, app_info in apps:
        var = tk.BooleanVar()
        check = tk.Checkbutton(frame, text=app_info.get('content', app_name), variable=var, font=theme["font"], bg=theme["bg"], fg=theme["fg"], anchor='w', justify=tk.LEFT)
        check_vars[app_name] = var
        check.grid(row=row, column=0, sticky='w', padx=10, pady=5)
        
        # Add question mark button
        question_button = tk.Button(frame, text="[?]", font=theme["font"], bg="#FF5733", fg="#FFFFFF", command=lambda name=app_name: show_description(name))
        question_button.grid(row=row, column=1, sticky='e', padx=10, pady=5)
        
        row += 1

# Update canvas scroll region
frame.update_idletasks()
canvas.config(scrollregion=canvas.bbox("all"))

# Right frame for controls
right_frame = tk.Frame(main_frame, bg=theme["bg"])
right_frame.pack(side=tk.RIGHT, fill=tk.Y)

# Button to check installed programs
check_installed_button = tk.Button(right_frame, text="Check Installed Programs", font=theme["font"], bg=theme["button_bg"], fg=theme["button_fg"], command=check_installed_programs)
check_installed_button.pack(pady=10, padx=10, fill=tk.X)

# Button to install selected programs
install_button = tk.Button(right_frame, text="Install Selected", font=theme["font"], bg=theme["button_bg"], fg=theme["button_fg"], command=lambda: install_or_update_programs('install'))
install_button.pack(pady=10, padx=10, fill=tk.X)

# Button to uninstall selected programs
uninstall_button = tk.Button(right_frame, text="Uninstall Selected", font=theme["font"], bg=theme["button_bg"], fg=theme["button_fg"], command=lambda: install_or_update_programs('uninstall'))
uninstall_button.pack(pady=10, padx=10, fill=tk.X)

# Button to unselect all programs
unselect_all_button = tk.Button(right_frame, text="Unselect All", font=theme["font"], bg=theme["button_bg"], fg=theme["button_fg"], command=unselect_all_programs)
unselect_all_button.pack(pady=10, padx=10, fill=tk.X)

# StringVar for program description
description_text = tk.StringVar()
description_label = tk.Label(right_frame, textvariable=description_text, font=theme["font"], wraplength=300, justify=tk.LEFT)
description_label.pack(pady=10, padx=10, fill=tk.X)

# Progress bar
progress_bar = ttk.Progressbar(right_frame, orient="horizontal", mode="determinate")
progress_bar.pack(pady=10, padx=10, fill=tk.X)

# Label for progress status
progress_label = tk.Label(right_frame, text="", font=theme["font"], bg=theme["bg"], fg=theme["fg"])
progress_label.pack(pady=10, padx=10, fill=tk.X)

# Run the GUI main loop
if __name__ == "__main__":
    root.mainloop()
