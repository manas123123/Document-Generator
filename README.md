# Document-Generator
Python-based Word document generator using pandas and pywin32 to populate templates from DataFrame data. Ensures unique filenames, prevents template modification, and automates text replacement via COM.


Word Document Generator

Overview

This Python-based software automates the generation of Microsoft Word documents using data from a pandas DataFrame. It utilizes pywin32 to manipulate Word templates, replacing placeholders with relevant data while ensuring unique filenames and preventing template modifications. Designed for efficiency, the program streamlines document creation at the workplace.

Features

Reads data from a DataFrame.

Uses a predefined Word template.

Replaces placeholders dynamically.

Generates uniquely named files.

Prevents accidental template modifications.

Provides a progress tracker during document generation.

Technologies Used

Python

pandas

pywin32

tkinter (for UI interactions)

Installation

Ensure Python is installed.

Install dependencies:

pip install pandas pywin32

Clone the repository:

git clone <repository_url>

Run the script:

python generate_docs.py

Usage

Select the input data file.

Choose the output directory.

Run the program to generate documents.

License

This project is intended for internal workplace use. Modify as needed for your requirements.
