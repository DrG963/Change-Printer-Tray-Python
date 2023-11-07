# info from here: https://superuser.com/questions/684061/automate-windows-scanning-feature

# Create object to access the scanned image later
$imageProcess = new-object -ComObject WIA.ImageProcess

# Store file format GUID strings
$wiaFormatBMP  = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatPNG  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatGIF  = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"

# Scan the image from scanner as BMP
foreach ($item in $device.Items) {
    $image = $item.Transfer() 
}

# set type to JPEG and quality/compression level
$imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
$imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatJPEG
$imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 5
$image = $imageProcess.Apply($image)

# Build filepath from desktop path and filename 'Scan 0'
# $filename = "$([Environment]::GetFolderPath("Desktop"))\Scan {0}.jpg"
$filename = "superpythonpointer"

# If a file named 'Scan 0' already exists, increment the index as long as needed
$index = 0
while (test-path ($filename -f $index)) {[void](++$index)}
$filename = $filename -f $index

# Save image to 'C:\Users\<username>\Desktop\Scan {x}'
$image.SaveFile($filename)

# Show image 
& $filename