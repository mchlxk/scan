$dpi = $args[0] -as [int]
$width = $args[1] -as [float]
$height = $args[2] -as [float]

$quality = 85

### access the scanner
$deviceManager = new-object -ComObject WIA.DeviceManager

if( $deviceManager.DeviceInfos.Count -eq 0)
{
    Write-Host "No scanner!"
    exit
}
$device = $deviceManager.DeviceInfos.Item(1).Connect()


# https://docs.microsoft.com/en-us/windows/win32/wia/-wia-wia-property-constant-definitions

$device.Items(1).Properties("6147").Value = $dpi   # horizontal DPI
$device.Items(1).Properties("6148").Value = $dpi   # vertical DPI

$device.Items(1).Properties("6149").Value = 0  # x point where to start scan
$device.Items(1).Properties("6150").Value = 0  # y point where to start scan

### scan width
$widthPx = [math]::floor($width * $dpi * 210 / 25.4)
$device.Items(1).Properties("6151").Value = [int]$widthPx 

### scan height
$heightPx = [math]::floor($height * $dpi * 297 / 25.4)
$device.Items(1).Properties("6152").Value = [int]$heightPx 

# $device.Items(1).Properties("6146").Value = 2   # colors
# $device.Items(1).Properties("4104").Value = 8   # bits per pixel


### Scan the image from scanner as BMP
foreach ($item in $device.Items)
{
    $image = $item.Transfer() 
}


### process scanned image
$BMP  = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
$PNG  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$GIF  = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
$JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
$TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"

$imageProcess = new-object -ComObject WIA.ImageProcess
$imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
$imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = [string]$JPEG
$imageProcess.Filters.Item(1).Properties.Item("Quality").Value = $quality
$image = $imageProcess.Apply($image)

# $folder = $([Environment]::GetFolderPath("Desktop"))
$folder = "."

$timestamp = New-TimeSpan "01 January 1970 00:00:00"
$timestamp = $timestamp.TotalSeconds
$timestamp = [uint64]$timestamp

$filename = "{0}\{1}.jpg" -f $folder, $timestamp

$image.SaveFile($filename)
Write-Host $filename

### Show image 
# & $filename

