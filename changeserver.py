import win32print
# import win32gui
from flask import Flask, request

app = Flask(__name__)

device_name = "Brother MFC-J6955DW Printer"

@app.route('/changeprintertray')
def change_printer_tray():
    tray_number = request.args.get('tray', type=int)

    if tray_number is None:
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

if __name__ == "__main__":
    app.run(debug=True)