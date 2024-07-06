from flask import Flask, request, render_template
import os
import datetime
import openpyxl

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('Generate_file_list.html')

@app.route('/generate_file_list', methods=['POST'])
def generate_file_list():
    try:
        # Get inputs from the form
        master_directory = request.form['master_directory']
        x_days = int(request.form['x_days'])
        output_file = request.form['output_file']

        # Function to get last x days folders
        def get_last_x_days_folders(master_directory, x):
            today = datetime.date.today()
            folders = []

            for i in range(x):
                day = today - datetime.timedelta(days=i)
                month_year = day.strftime("%b").upper() + " " + day.strftime("%Y")
                day_month_year = day.strftime("%d-%m-%Y")
                folder_path = os.path.join(master_directory, month_year, day_month_year)
                folders.append(folder_path)

            return folders

        # Function to list files in folders
        def list_files_in_folders(folder_list):
            main_list = []

            for folder in folder_list:
                if os.path.exists(folder):
                    for root, dirs, files in os.walk(folder):
                        for file in files:
                            if file.lower().endswith(('.pdf', '.tiff', '.tif')):
                                main_list.append(os.path.join(root, file))
                else:
                    print(f"Folder not found: {folder}")

            return main_list

        # Function to write to Excel
        def write_to_excel(file_paths, output_file):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "File List"

            sheet.append(["File Path"])
            for file_path in file_paths:
                sheet.append([file_path])

            workbook.save(output_file)

        # Get folders to check
        folders_to_check = get_last_x_days_folders(master_directory, x_days)

        # Get list of files in folders
        files_list = list_files_in_folders(folders_to_check)

        # Write file list to Excel
        write_to_excel(files_list, output_file)

        return f"<h1>File list written to {output_file}</h1>"
    except Exception as e:
        return f"<h1>Error:</h1><pre>{str(e)}</pre>"

if __name__ == '__main__':
    app.run(debug=True)
