import win32com.client
import os

def acquire_image_wia(scanner_name=None):
    WIA_IMG_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"

    # wia = win32com.client.Dispatch("WIA.CommonDialog")
    device_manager = win32com.client.Dispatch("WIA.DeviceManager")

    # Select the first available scanner or a specific one by name
    scanner_device = None
    for device in device_manager.DeviceInfos:
        if device.Type == 1:  # Type 1 corresponds to scanner
            if device.Properties["Name"].Value in scanner_name:
                scanner_device = device.Connect()
                break
            if scanner_name is None:
                pass
            else:
                scanner_device = device.Connect()
                break
        
    if not scanner_device:
        print("No WIA scanner device found.")
        return

    # Assuming the scanner is the first item (which it typically is)
    item = scanner_device.Items[1]

    image = item.Transfer(WIA_IMG_FORMAT_PNG)
    fname = 'C:\\Users\\grbsk\\OneDrive\\Desktop\\Test Photos\\wia-test.jpg'
    
    if os.path.exists(fname):
        os.remove(fname)
    
    image.SaveFile(fname)
    print(f"Image saved to {fname}")

if __name__ == "__main__":
    # Call the function with the name of the scanner if known, else it will pick the first available scanner
    acquire_image_wia("Your Scanner Name Here")
