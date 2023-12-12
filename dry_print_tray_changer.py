import win32print
import win32gui

# Get a handle for the default printer
# device_name = win32print.GetDefaultPrinter()
device_name = "Brother MFC-J6955DW Printer"
PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}
handle = win32print.OpenPrinter(device_name, PRINTER_DEFAULTS)

# Get the default properties for the printer
properties = win32print.GetPrinter(handle, 2)
devmode = properties['pDevMode']

# Print the default paper source (prints '7' for 'Automatically select')
print(devmode.DefaultSource)

# Change the default paper source to '4' for 'Manual feed', 7 for auto? 
# Update! 1 is tray 1, 2 is tray 2. Easy enough.
devmode.DefaultSource = 1

# Write these changes back to the printer
# win32print.DocumentProperties(None, handle, device_name, devmode, devmode, DM_IN_BUFFER | DM_OUT_BUFFER)
properties["pDevMode"]=devmode #write the devmode back to properties
win32print.SetPrinter(handle,2,properties,0) #save the properties to the printer

# Confirm the changes were updated
print(devmode.DefaultSource) 

# Start printing with the device
hdc = win32gui.CreateDC('', device_name, devmode)
win32print.StartDoc(hdc, ('Test', None, None, 0))
win32print.StartPage(hdc)

# ... GDI drawing commands ...

win32print.EndPage(hdc)
win32print.EndDoc(hdc)
