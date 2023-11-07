import win32print
# import win32gui
from flask import Flask, request, jsonify
# import pyscanner as scanner
import os
from flask_cors import CORS, cross_origin
import subprocess
import json

app = Flask(__name__)

cors = CORS(app, resources={r"*": {"origins": "http://192.168.1.98:5000"}})

# define printer name exactly as listed here...
device_name = "Brother MFC-J6955DW Printer"

ps1_path = './initscan.ps1'

network_path = '\\\\tpcserver\\jobFiles\\'

method = "c"
# can be c for c# or p for powershell

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
            with open('./savepath.json', 'w') as config_file:
                json.dump(config, config_file)

            process = subprocess.Popen(['./scannerapp.exe'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = process.communicate()
            print("Standard Output:")
            print(stdout.decode('utf-8'))
            if stderr:
                print("Standard Error:")
                print(stderr.decode('utf-8'))

            # Wait for the process to finish.
            process.wait()
            return jsonify({'message': 'Scanned document saved successfully', 'file_path': full_file_path, "filename": file_name, "web_path": network_web_path}), 200
        except Exception as error:
            print(error)
            return jsonify({'error': 'there was a problem in scan_and_save c version'}), 400

if __name__ == "__main__":
    # Check if the network path exists
    if os.path.exists(network_path):
        # Access the network location
        try:
            app.run(debug=True, port=3030)
        except Exception as e:
            print(f"Error accessing network location: {e}")
    else:
        raise Exception(f"Network path '{network_path}' does not exist.")