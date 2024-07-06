from flask import Flask, request, render_template
from datetime import datetime, timedelta
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('Create_folder.html')

@app.route('/create_folders', methods=['POST'])
def create_folders():
    try:
        # Read the Excel file
        Machine_list = pd.read_excel('Machine List.xlsx')

        # Get the master directory path from the form
        master_directory = request.form['path']

        # Calculate tomorrow's date
        nextday = datetime.now() + timedelta(1)
        folder_name = nextday.strftime("%d-%m-%Y")

        # Create the main directory for tomorrow's date
        os.makedirs(os.path.join(master_directory, folder_name), exist_ok=True)

        # Create subdirectories for each machine
        for index, row in Machine_list.iterrows():
            machine_dir = os.path.join(master_directory, folder_name, row['Machine List'])
            os.makedirs(machine_dir, exist_ok=True)
            os.makedirs(os.path.join(machine_dir, "Today"), exist_ok=True)
            os.makedirs(os.path.join(machine_dir, "Tomorrow"), exist_ok=True)

        return f"<h1>Folders Created Successfully!</h1><p>Main Directory: {os.path.join(master_directory, folder_name)}</p>"
    except Exception as e:
        return f"<h1>Error:</h1><pre>{str(e)}</pre>"

if __name__ == '__main__':
    app.run(debug=True)