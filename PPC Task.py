from flask import Flask, request, render_template
import os
import shutil
import pandas as pd
from datetime import datetime
from pathlib import Path
import re

app = Flask(__name__)


# Function to move files with long filenames
def move_long_filename(source, destination):
    try:
        # On Windows, prefix the path with \\?\ to handle long paths
        if os.name == 'nt':
            source = f"\\\\?\\{source}"
            destination = f"\\\\?\\{destination}"

        # Convert to pathlib Path objects
        src_path = Path(source)
        dest_path = Path(destination)

        # Ensure the destination directory exists
        if not dest_path.parent.exists():
            dest_path.parent.mkdir(parents=True, exist_ok=True)

        # Move the file
        shutil.move(str(src_path), str(dest_path))

        print(f"File moved successfully from {src_path} to {dest_path}")
    except Exception as e:
        print(f"Error: {e}")
        if not src_path.exists():
            print(f"The source path {src_path} does not exist.")
        if not dest_path.parent.exists():
            print(f"The destination directory {dest_path.parent} does not exist.")


# Function to check if all keywords are in the filename
def check_keywords_in_filename(filename, keywords):
    return all(keyword.lower() in filename.lower() for keyword in keywords)


# Function to read order priority file
def read_order_priority(order_priority_file):
    df_order_priority = pd.read_excel(order_priority_file, sheet_name='Days')
    return df_order_priority.set_index('Ref. Order')


# Function to read master file (keywords.xlsx in the project folder)
def read_master_sheet():
    master_file = 'keywords.xlsx'
    df_master = pd.read_excel(master_file)
    keyword_machine_mapping = {}
    for index, row in df_master.iterrows():
        keywords = [kw.strip() for kw in row['keyword'].split(',')]
        machine = row['machine_name']
        keyword_machine_mapping[tuple(keywords)] = machine
    return keyword_machine_mapping


# Function to identify machines based on file names
def identify_machines(file_names, keyword_machine_mapping, order_priority_df, today_date):
    results = []

    for filename in file_names:
        machine_name = None
        priority = None

        match = re.search(r'-(.*?)_', filename)
        if match:
            extracted_word = match.group(1)
        else:
            extracted_word = "Pattern not found"

        if extracted_word in order_priority_df.index:
            priority = order_priority_df.loc[extracted_word, 'Priority']
            print(f"Priority for {filename}: {priority}")
        else:
            print(f"Priority not found for {filename}")

        for keywords, machine in keyword_machine_mapping.items():
            if check_keywords_in_filename(filename, keywords):
                machine_name = machine
                break  # Once the machine is identified, stop checking

        if machine_name is None:
            print(f"No match found for: {filename}")

        results.append((filename, machine_name, priority))

    return results


# Function to move files to respective machine folders
def move_files(file_mappings, source_folder, master_folder):
    today_date = datetime.now().strftime('%d-%m-%Y')
    for filename, machine_name, priority in file_mappings:
        if machine_name:
            if priority == 'Today':
                dest_folder = os.path.join(master_folder, today_date, machine_name, 'Today')
            elif priority == 'Tomorrow':
                dest_folder = os.path.join(master_folder, today_date, machine_name, 'Tomorrow')
            else:
                dest_folder = os.path.join(master_folder, today_date, machine_name)

            os.makedirs(dest_folder, exist_ok=True)
            src_file = os.path.join(source_folder, filename)
            dest_file = os.path.join(dest_folder, filename)
            move_long_filename(src_file, dest_file)
        else:
            print(f"No machine name for file: {filename}")


# Main function to process files
def process_files_main(order_priority_file, source_folder, master_folder):
    keyword_machine_mapping = read_master_sheet()
    print(f"Keyword to machine mapping: {keyword_machine_mapping}")
    order_priority_df = read_order_priority(order_priority_file)
    entries = os.listdir(source_folder)
    file_names = [file for file in entries if file.endswith(".pdf") or file.endswith(".tiff") or file.endswith(".tif")]
    results = identify_machines(file_names, keyword_machine_mapping, order_priority_df,
                                datetime.now().strftime('%d-%m-%Y'))

    for result in results:  # Debug statement
        print(result)

    move_files(results, source_folder, master_folder)

    return f"<h1>Files processed successfully to {master_folder}</h1>"


@app.route('/')
def index():
    return render_template('Process_files.html')


@app.route('/process_files', methods=['POST'])
def process_files():
    try:
        # Get inputs from the form
        order_priority_file = request.form['order_priority_file']
        source_folder = request.form['source_folder']
        master_folder = request.form['master_folder']

        # Execute the main function
        result_message = process_files_main(order_priority_file, source_folder, master_folder)

        return result_message

    except FileNotFoundError:
        return f"<h1>File or directory not found. Check your paths.</h1>"
    except PermissionError:
        return f"<h1>Permission denied to access file or directory.</h1>"
    except Exception as e:
        return f"<h1>An error occurred: {e}</h1>"


if __name__ == '__main__':
    app.run(debug=True)