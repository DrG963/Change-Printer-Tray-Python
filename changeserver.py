import win32print
# import win32gui
from flask import Flask, request, jsonify
import scanner

app = Flask(__name__)

device_name = "Brother MFC-J6955DW Printer"

@app.route('/changeprintertray')
def change_printer_tray():
    tray_number = request.args.get('tray', type=int)

    if tray_number is None and tray_number != 1 and tray_number != 2:
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
    return f"Changed printer tray to {tray_number}"

@app.route('/scan_and_save', methods=['POST'])
def scan_and_save():
    # Get data from the JSON request
    data = request.get_json()
    filename = data.get('filename')
    filetype = data.get('filetype')
    customername = data.get('customername')
    tpcID = data.get('tpcID')

    # Check if required data is missing
    if not filename or not filetype or not customername or not tpcID:
        return jsonify({'error': 'Missing data'}), 400

    # List available scanners
    available_scanners = scanner.Scanner().get_scanners()

    if available_scanners:
        selected_scanner = available_scanners[0]  # Choose the first available scanner

        # Configure scanning parameters (e.g., resolution, color mode, etc.)
        scan_settings = scanner.ScannerSettings()
        scan_settings.color_mode = scanner.ColorMode.COLOR
        scan_settings.resolution = 300  # DPI

        # Queue a scan
        scanned_image = selected_scanner.scan(scan_settings)

        # Save the scanned image to a file with the provided filename and extension
        file_path = f"{filename}.{filetype}"
        with open(file_path, "wb") as f:
            f.write(scanned_image)

        return jsonify({'message': 'Scanned document saved successfully', 'file_path': file_path})
    else:
        return jsonify({'error': 'No scanners found'})

if __name__ == "__main__":
    app.run(debug=True)