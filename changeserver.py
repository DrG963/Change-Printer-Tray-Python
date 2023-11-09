import win32print
# import win32gui
from flask import Flask, request, jsonify
# import pyscanner as scanner
import os
from flask_cors import CORS, cross_origin
import subprocess
import json
# import pyuac
from PIL import Image

app = Flask(__name__)

cors = CORS(app, resources={r"*": {"origins": "*"}})

# define printer name exactly as listed here...
device_name = "Brother MFC-J6955DW Printer"

ps1_path = './initscan.ps1'

network_path = '\\\\tpcserver\\jobFiles\\'

method = "c"
# can be c for c# or p for powershell


def resize_scan(img_path):
    # Path to the original image
    original_image_path = img_path

    # Open the original image
    image = Image.open(original_image_path)

    # Define the desired width and height for letter size paper
    desired_width = 8.5 * 300  # 8.5 inches at 300 DPI
    desired_height = 11 * 300  # 11 inches at 300 DPI

    # Resize the image while maintaining aspect ratio
    image.thumbnail((desired_width, desired_height))

    # Get the directory and filename of the original image
    image_directory, image_filename = os.path.split(original_image_path)

    # Generate the path for the resized image with the same name
    resized_image_path = os.path.join(image_directory, "resized_" + image_filename)

    # Save the resized image with the same name as the original
    image.save(resized_image_path)

    # Delete the original image
    os.remove(original_image_path)

    # Rename the resized image to have the same name as the original
    os.rename(resized_image_path, original_image_path)

    print(f"Resized and renamed: {original_image_path}")

@app.route('/changeprintertray')
@cross_origin()
def change_printer_tray():
    tray_number = request.args.get('tray', type=int)

    if tray_number is None and tray_number != 1 and tray_number != 2:
        print("Invalid Tray Number")
        return "Invalid tray number", 400

    # Get a handle for the default printer
    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    handle = win32print.OpenPrinter(device_name, PRINTER_DEFAULTS)

    # Get the default properties for the printer
    properties = win32print.GetPrinter(handle, 2)
    devmode = properties['pDevMode']

    # Change the default paper source based on the provided tray number
    devmode.DefaultSource = tray_number

    # Write these changes back to the printer
    properties["pDevMode"] = devmode
    win32print.SetPrinter(handle, 2, properties, 0)

    # Confirm the changes were updated
    print(f"Changed printer tray to {tray_number}")
    return f"Changed printer tray to {tray_number}"

@app.route('/scan_and_save', methods=['POST'])
@cross_origin()
def scan_and_save():
    # Get data from the JSON request
    data = request.get_json()
    filename = data.get('filename')
    tpcID = data.get('jobID')
    customerID = data.get("customerID")
    contactID = data.get("contactID")

    network_path = '\\\\tpcserver\\jobFiles\\'
    # Generate the file path based on the provided IDs
    file_directory = f"{network_path}\\C.{customerID}\\P.{contactID}\\J.{tpcID}\\"
    os.makedirs(file_directory, exist_ok=True)
    file_name = f"{filename}.jpg"
    full_file_path = file_directory + file_name
    network_web_path = f"//tpcserver/TPC/jobFiles/C.{customerID}/P.{contactID}/J.{tpcID}/{file_name}"

    print(data)

    # Check if required data is missing
    if not filename or not customerID or not contactID or not tpcID:
        print("Missing Data!")
        return jsonify({'error': 'Missing data'}), 400
    
    if method == "p":
        try:
            # open PS1 file, change file path text
            with open(ps1_path, "r+") as file:
                content = file.read()
                content = content.replace("superpythonpointer", full_file_path)
                file.seek(0)
                file.write(content)
                file.truncate()

            # run PS1 file
            # Run the PowerShell script
            result = subprocess.run(['powershell', '-File', ps1_path], capture_output=True, text=True)

            # open PS1 file, change file path text back for next cycle
            with open(ps1_path, "r+") as file:
                content = file.read()
                content = content.replace(full_file_path, "superpythonpointer")
                file.seek(0)
                file.write(content)
                file.truncate()

            # Check if the script ran successfully
            if result.returncode == 0:
                print("PowerShell script ran successfully.")
                print("Script output:")
                print(result.stdout)
            else:
                print("PowerShell script encountered an error.")
                print("Error output:")
                print(result.stderr)
                return jsonify({'error': 'PS1 File execution FAILED'}), 400

            print("FILE SAVED AND WHATNOT")

            return jsonify({'message': 'Scanned document saved successfully', 'file_path': full_file_path, "filename": file_name, "web_path": network_web_path}), 200

        except Exception as error:
            return jsonify({'error': 'there was a problem in scan_and_save'}), 400
        
    if method == "c":
        config = {}
        try:
            # Modify the save directory
            config['saveDirectory'] = file_directory
            config['filename'] = file_name
            # Save the updated configuration
            with open('scannerapp\\bin\Debug\\net8.0\\savepath.json', 'w') as config_file:
                json.dump(config, config_file)

            process = subprocess.Popen(['scannerapp\\bin\\Debug\\net8.0\\scannerapp.exe'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = process.communicate()
            print("Standard Output:")
            csharp_output = stdout.decode('utf-8')
            print(csharp_output)
            if "ERROR" in csharp_output.upper():
                return jsonify({'error': 'there was a problem in scan_and_save c version - error was in output text'}), 400
            if stderr:
                print("Standard Error:")
                print(stderr.decode('utf-8'))
                return jsonify({'error': 'there was a problem in scan_and_save c version - stderr output type detected'}), 400
            
            # Wait for the process to finish.
            process.wait()

            resize_scan(full_file_path)
            return jsonify({'message': 'Scanned document saved successfully', 'file_path': full_file_path, "filename": file_name, "web_path": network_web_path}), 200
        except Exception as error:
            print(error)
            return jsonify({'error': 'there was a problem in scan_and_save c version'}), 400

if __name__ == "__main__":
    # Check if the network path exists
    if os.path.exists(network_path):
        app.run(debug=True, port=3030)
    else:
        raise Exception(f"Network path '{network_path}' does not exist.")