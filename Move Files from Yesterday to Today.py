from flask import Flask, request, render_template
import os
from datetime import datetime, timedelta
import shutil
from pathlib import Path

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('Move_files.html')


@app.route('/move_files', methods=['POST'])
def move_files():
    try:
        # Get input from the form
        master_directory = request.form['master_directory']

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

        # Get today's and next day's date in "%d-%m-%Y" format
        today = datetime.now().strftime("%d-%m-%Y")
        nextday = (datetime.now() + timedelta(1)).strftime("%d-%m-%Y")

        # Construct read path and destination path
        read_path = os.path.join(master_directory, today)
        destination_path = os.path.join(master_directory, nextday)

        extensions = (".pdf", ".tiff")

        # List all files in read path (excluding directories containing 'ok')
        all_files = []
        for root, dirs, files in os.walk(read_path):
            # Normalize the root path
            root = os.path.normpath(root)

            # Exclude directories containing 'ok' in any case
            dirs[:] = [d for d in dirs if 'ok' not in d.lower()]

            for file in files:
                if file.lower().endswith(extensions):
                    # Normalize the file path
                    file_path = os.path.normpath(os.path.join(root, file))
                    all_files.append(file_path)

        print(f"Files in '{read_path}' (excluding 'Ok' folders):")
        for file in all_files:
            new_file = file.replace("\\", "/")
            path_backend = new_file.split(today)[1]
            destination = os.path.join(destination_path, path_backend)
            print(new_file)
            print(destination)
            move_long_filename(new_file, destination)

        return f"<h1>Files moved successfully to {destination_path}</h1>"

    except FileNotFoundError:
        return f"<h1>The directory '{read_path}' does not exist.</h1>"
    except PermissionError:
        return f"<h1>Permission denied to access '{read_path}'.</h1>"
    except Exception as e:
        return f"<h1>An error occurred: {e}</h1>"


if __name__ == '__main__':
    app.run(debug=True)