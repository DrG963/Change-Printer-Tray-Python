using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using Newtonsoft.Json.Linq;
using WIA; // Windows Image Acquisition

class Program
{
    static void Main()
    {
        // Your existing code for reading configuration...
        string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string configPath = Path.Combine(exeDirectory, "config.json");

        if (File.Exists(configPath))
        {
            string configText = File.ReadAllText(configPath);
            JObject config = JObject.Parse(configText);

            string saveDirectory = config["saveDirectory"].ToString();
            string filename = config["filename"].ToString();

            Console.WriteLine($"Save Directory: {saveDirectory}");
            ScanDocument(saveDirectory, filename);
        }
        else
        {
            Console.WriteLine("Configuration file not found.");
        }
    }

    static void ScanDocument(string saveDirectory, string filename)
    {
        try
        {
            CommonDialog dialog = new CommonDialog();
            Device device = dialog.ShowSelectDevice(WiaDeviceType.ScannerDeviceType, false, false);

            if (device != null)
            {
                Item item = device.Items[1] as Item;

                if (item != null)
                {
                    // Your existing code for scanner settings...
                    item.Properties["6146"].set_Value(1); // Color
                    item.Properties["6147"].set_Value(300); // DPI

                    // Transfer the image in its native format (BMP)
                    ImageFile image = (ImageFile)item.Transfer(FormatID.wiaFormatBMP);

                    // Save the image temporarily
                    string tempPath = Path.Combine(saveDirectory, "temp.bmp");
                    image.SaveFile(tempPath);

                    // Convert and save the image as JPEG
                    string finalFileName = Path.Combine(saveDirectory, filename);
                    using (Image img = Image.FromFile(tempPath))
                    {
                        img.Save(finalFileName, ImageFormat.Jpeg);
                    }

                    // Delete the temporary BMP file
                    File.Delete(tempPath);

                    Console.WriteLine($"Scanned image saved to: {finalFileName}");
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
