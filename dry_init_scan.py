import win32com.client
import os

def acquire_image_wia(save_path, scanner_name=None):
    WIA_IMG_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    device_manager = win32com.client.Dispatch("WIA.DeviceManager")

    scanner_device = None
    for device in device_manager.DeviceInfos:
        if device.Type == 1:  # Type 1 corresponds to scanner
            device_name = device.Properties["Name"].Value

            # Check if the scanner name matches or if scanner_name is None
            if scanner_name is None or scanner_name in device_name:
                scanner_device = device.Connect()
                break
        
    if not scanner_device:
        print("No WIA scanner device found.")
        return

    # Assuming the scanner is the first item
    item = scanner_device.Items[1]
    image = item.Transfer(WIA_IMG_FORMAT_PNG)
    fname = save_path

    if os.path.exists(fname):
        os.remove(fname)
    
    image.SaveFile(fname)
    print(f"Image saved to {fname}")

if __name__ == "__main__":
    # Call the function with the name of the scanner if known, else it will pick the first available scanner
    acquire_image_wia(save_path='C:\\Users\\grbsk\\OneDrive\\Desktop\\Test Photos\\wia-test.jpg')
