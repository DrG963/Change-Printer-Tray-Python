using System;
using System.IO;
using Newtonsoft.Json.Linq;
using WIA; // Assuming you're using Windows Image Acquisition (WIA)

class Program
{
    static void Main()
    {
        // Get the current directory where the executable is located
        string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;

        // Construct the path to the configuration file
        string configPath = Path.Combine(exeDirectory, "savepath.json");

        // Check if the configuration file exists
        if (File.Exists(configPath))
        {
            // Read the configuration file
            string configText = File.ReadAllText(configPath);

            // Parse the JSON data
            JObject config = JObject.Parse(configText);

            // Extract the save directory path
            string saveDirectory = config["saveDirectory"].ToString();
            string saveFileName = config["filename"].ToString();

            // Use the save directory path as needed
            Console.WriteLine($"Save Directory: {saveDirectory}, save file name: {saveFileName}");

            // Perform the scanning operation
            ScanDocument(saveDirectory, saveFileName);
        }
        else
        {
            Console.WriteLine("Configuration file not found.");
        }
    }

    static void ScanDocument(string saveDirectory, string saveFileName)
    {
        try
        {
            // Create a new CommonDialog class
            CommonDialog dialog = new CommonDialog();

            // Show the scanner dialog
            Device device = dialog.ShowSelectDevice(WiaDeviceType.ScannerDeviceType, false, false);

            if (device != null)
            {
                // Create a new Item class
                Item item = device.Items[1] as Item;

                if (item != null)
                {
                    // Configure the scan settings (e.g., color, resolution, etc.)
                    // Modify these settings as needed
                    item.Properties["6146"].set_Value(1); // Color
                    item.Properties["6147"].set_Value(300); // DPI

                    // Perform the scan and save the image to the specified directory
                    ImageFile image = (ImageFile)item.Transfer(WiaFormatType.wiaFormatJPEG);
                    string fileName = Path.Combine(saveDirectory, saveFileName);
                    image.SaveFile(fileName);
                    Console.WriteLine($"Scanned image saved to: {fileName}");
                }
                else
                {
                    Console.WriteLine("No scanning item found.");
                }
            }
            else
            {
                Console.WriteLine("No scanning device selected.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error scanning document: {ex.Message}");
        }
    }
}
