import win32com.client
import time
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
from tkinter import filedialog, messagebox
import openpyxl
from tkinter import ttk
import sys
import shutil
import os


# Function to replace table content in the header of the Word document
def replace_table_cell_content_in_header(doc, replacements):
    # Access all types of headers in the first section
    section = doc.Sections(1)
    primary_header = section.Headers(1)  # Primary header
    first_page_header = section.Headers(2)  # First-page header
    even_page_header = section.Headers(3)  # Even-page header

    # Check for tables in each header's range
    for header in [primary_header, first_page_header, even_page_header]:
        if header.Range.Tables.Count > 0:  # Access tables via the Range property
            table = header.Range.Tables(1)  # Access the first table in the header
            for row in table.Rows:
                for cell in row.Cells:
                    text = cell.Range.Text.strip()  # Get the text in the cell, strip trailing '\r\x07'
                    for placeholder, value in replacements.items():
                        if placeholder in text:
                            cell.Range.Text = text.replace(placeholder, value)
            break  # Exit the loop once the desired header is found
    else:
        print("No tables found in any headers.")

def output_file_generator(month,  year,centre, room, file):
    file = f"{file}\\{centre}\\{centre}{room}\\{month.upper()} {year}\\"
    
    if not os.path.exists(file):
        os.makedirs(file)  
    
    return file

def generate_document(df):
    # Open Microsoft Word application once
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Set to True for debugging purposes

    # Hardcoded template file path (you can modify as needed)
    if getattr(sys, 'frozen', False):
        template_file = os.path.join(sys._MEIPASS, "template.docx")
    else:
        template_file = r"C:\Users\agraw\OneDrive\Desktop\CODE\Work\template.docx"

    output_file = select_output_directory()
    output_file_copy = output_file
    progress_counter = 0

    for index, row in df.iterrows():
        # Copy the template file to avoid modifying the original
        temp_template_path = os.path.join(output_file, "temp_template.docx")
        temp_template_path = os.path.abspath(os.path.normpath(temp_template_path))  # Normalize path

        shutil.copy(template_file, temp_template_path)

        # Verify if the file exists before opening
        if not os.path.exists(temp_template_path):
            print(f"Error: Temporary template file not found at {temp_template_path}")
            messagebox.showerror("File Error", f"Temporary template file not found at {temp_template_path}")
            continue  # Skip this row and proceed to the next one

        # Open the copied template document for each student
        output_file = output_file_copy
        try:
            doc = word.Documents.Open(temp_template_path)
        except Exception as e:
            print(f"Error opening file: {e}")
            messagebox.showerror("Error", f"Error opening template file: {e}")
            continue  # Skip to the next row if the document fails to open

        # Prepare the replacements for the current row
        name = f"{row['Student']}"
        id = str(row['ID'])
        date = f"{row['Month']} {row['Day']:02d}, {row['Year']}"
        course = f"{row['Course_Name']} {row['Course_Code']} {row['Course_Section']}"

        if not name or not id or not date:
            messagebox.showwarning("Input Error", "Please fill in all fields.")
            continue  # Skip to the next row

        output_file = output_file_generator(row["Month"], row["Year"], row["Centre"], row["Room"], output_file)
        
        # Data for replacement in the document
        replacements = {"{{Name}}": name, "{{ID}}": id, "{{Date}}": date, "{{Course}}": course}
        
        # Replace content in the table cells of the header
        replace_table_cell_content_in_header(doc, replacements)
        
        # Ensure unique filename by appending the row index or timestamp
        file_name = f"{str(row['ID'])[-6:]}.{row['Month']}.{row['Day']:02d}.{row['Year']}.{row['Course_Name']}.{row['Course_Code']}.{row['Course_Section']}.docx"
        output_file_with_timestamp = os.path.join(output_file, file_name)  # Ensure correct path

        # Check if the file already exists
        counter = 1
        original_output_file = output_file_with_timestamp
        while os.path.exists(output_file_with_timestamp):
            output_file_with_timestamp = f"{original_output_file[:-5]}_{counter}.docx"
            counter += 1

        total_rows = len(df)
        progress_counter += 1
        progress = (progress_counter) / total_rows * 100

        print(f"Progress: {int(progress)}% ({progress_counter}/{total_rows})")
        try:
            # Save the document with the new data for each student
            doc.SaveAs(output_file_with_timestamp, FileFormat=16)
            print(f"Document saved as: {file_name}")
        except Exception as e:
            print(f"Error saving file: {e}")
            messagebox.showerror("Error", f"Error saving file: {e}")
        finally:
            # Close the document after saving
            doc.Close()

        # Delete the temporary template file
        if os.path.exists(temp_template_path):
            os.remove(temp_template_path)

    # Quit Word application after processing all documents
    word.Quit()
def select_output_directory():
    """Opens a dialog for the user to select a directory for storing files."""
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    directory = filedialog.askdirectory(title="Select Individual Rooms")  

    if not directory:
        print("No directory selected. Please try again.")
        return None
    
    print(f"Selected Directory: {directory}")
    return directory  # Return the selected directory path



def locate_master_sheet():
    """Opens a file dialog for the user to select the Master Sheet."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    print("Please locate the Master Sheet for the prep date.")  # Instruction text

    # Open the file dialog
    file_path = filedialog.askopenfilename(
        title="Select Master Sheet",
        filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
    )

    if not file_path:
        print("No file selected. Please try again.")
        return None

    print(f"Selected File: {file_path}")
    
    # Read the file based on extension
    if file_path.endswith(".xlsx") or file_path.endswith(".xls"):
        df = pd.read_excel(file_path,header=1)
    elif file_path.endswith(".csv"):
        df = pd.read_csv(file_path,header=1)
    else:
        print("Unsupported file format. Please select an Excel or CSV file.")
        return None
    
    return df

def get_selected_date(df):  
    """Retrieve and print the selected day, month, and year."""
    day = day_var.get()
    month = month_var.get()
    year = year_var.get()
    
    if day and month and year:
        print(f"Selected Date: {year}-{month:02d}-{day:02d}")  # Print instead of setting label
    else:
        print("❌ Please select a valid date.")

    df['Day'] = day
    df['Month'] = month
    df['Year'] = year 
    month_map = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }
    df['Month'] = df['Month'].map(month_map)

    generate_document(df)


def initialize_df(df):
    """Initialize dataframe by cleaning and extracting necessary columns."""
    df.drop(df.columns[[0, 1, 3,6,7, 5,8,10,11]], axis=1, inplace=True)
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    df.dropna(inplace=True)
    df["Word"] = df.iloc[:, -2]
    df["ID"] = df.iloc[:, -2]
    df.drop(df.columns[3], axis=1, inplace=True)
    df.drop(df.columns[3], axis=1, inplace=True)
    df["Course_Name"] = df["Course"].str.split(" ", expand=True)[0]
    df["Course_Code"] = df["Course"].str.split(" ", expand=True)[1]
    df["Course_Section"] = df["Course"].str.split(" ", expand=True)[3]
    first_names = []
    last_names = []
    df["Room Booking"] = df["Room Booking"].str.replace(r'\s+', ' ', regex=True)  # Fix extra spaces

    for name in df['Student']:
        name_parts = name.split()
        first_name = name_parts[0]
        last_name = name_parts[-1]
        first_names.append(first_name)
        last_names.append(last_name)

    df['First_Name'] = first_names
    df['Last_Name'] = last_names
    
 
    df.drop(columns=["Word"], inplace=True)

    df['ID'] = df['ID'].astype(int)
    buildings = []
    rooms = []

    for booking in df["Room Booking"]:
        parts = booking.split()
        building = ''.join(filter(str.isdigit, parts[-2])) if len(parts) > 1 else ""
        room = parts[-1] if len(parts) > 0 else ""
        buildings.append(building)
        rooms.append(room)

    df['Centre'] = buildings
    df['Room'] = rooms
   
    df.drop(df.columns[2], axis=1, inplace=True)
    df.drop(df.columns[0], axis=1, inplace=True)
    
    return df

root = tk.Tk()

df = locate_master_sheet()
if df is not None:
    print("Master Sheet Loaded Successfully!")
    print(df.head()) 

df=initialize_df(df)
# Initialize the tkinter variables for day, month, and year
day_var = tk.IntVar(value=1)  # Default value for day
month_var = tk.IntVar(value=1)  # Default value for month
year_var = tk.IntVar(value=2025)  # Default value for year

# Day, Month, Year lists
days = list(range(1, 32))  # Days from 1 to 31
months = list(range(1, 13))  # Months from 1 to 12
years = list(range(2000, 2031))  # Years from 2000 to 2031



root.title("Select a Date")
root.geometry("300x250")  # Adjusted the height to accommodate labels and menus

# Add labels for each dropdown
tk.Label(root, text="Day").pack(pady=5)
day_menu = ttk.Combobox(root, textvariable=day_var, values=days, width=5, state="readonly")
day_menu.pack(pady=5)

tk.Label(root, text="Month").pack(pady=5)
month_menu = ttk.Combobox(root, textvariable=month_var, values=months, width=5, state="readonly")
month_menu.pack(pady=5)

tk.Label(root, text="Year").pack(pady=5)
year_menu = ttk.Combobox(root, textvariable=year_var, values=years, width=7, state="readonly")
year_menu.pack(pady=5)

# Button to confirm selection
tk.Button(root, text="Submit", command=lambda: get_selected_date(df)).pack(pady=10)



# Start the Tkinter event loop
root.mainloop()
