$dpi = $args[0] -as [int]
$widthMm = $args[1] -as [float]
$heightMm = $args[2] -as [float]

$widthPx = [math]::floor($widthMm * $dpi / 25.4)
$heightPx = [math]::floor($heightMm * $dpi / 25.4)

$widthPx = [math]::min($widthPx, $dpi * 210)
$heightPx = [math]::min($heightPx, $dpi * 297)

$quality = 80

### access the scanner
$deviceManager = new-object -ComObject WIA.DeviceManager

if( $deviceManager.DeviceInfos.Count -eq 0)
{
    Write-Host "No scanner!"
    exit
}
$device = $deviceManager.DeviceInfos.Item(1).Connect()


#-------------------

# https://docs.microsoft.com/en-us/windows/win32/wia/-wia-wia-property-constant-definitions
# https://github.com/tpn/winsdk-7/blob/master/v7.1A/Include/WiaDef.h
# https://stackoverflow.com/questions/25371269/scan-automation-with-powershell-and-wia-how-to-set-png-as-image-type


# horizontal DPI
$device.Items(1).Properties("6147").Value = $dpi

# vertical DPI
$device.Items(1).Properties("6148").Value = $dpi

# X/Y pivot
$device.Items(1).Properties("6149").Value = 0
$device.Items(1).Properties("6150").Value = 0

# width
$device.Items(1).Properties("6151").Value = [int]$widthPx 

# height
$device.Items(1).Properties("6152").Value = [int]$heightPx 

# intent
# https://learn.microsoft.com/cs-cz/windows-hardware/drivers/image/wia-ips-cur-intent
$device.Items(1).Properties("6146").Value = 0x00000001 -bor 0x00020000

# data type
# https://learn.microsoft.com/en-us/windows-hardware/drivers/image/wia-ipa-datatype
$device.Items(1).Properties("4103").Value = 3

# bits per pixel
$device.Items(1).Properties("4104").Value = 24

# contrast
# https://learn.microsoft.com/en-us/windows-hardware/drivers/image/wia-ips-contrast
$device.Items(1).Properties("6155").Value = 0

# brightness
# https://learn.microsoft.com/en-us/windows-hardware/drivers/image/wia-ips-brightness
$device.Items(1).Properties("6154").Value = 0

#-------------------

$wiaFormatBmp  = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatPng  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatGif  = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatJpeg = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatTiff = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"


foreach ($item in $device.Items)
{
    $image = $item.Transfer($wiaFormatJpeg) 
}


### process scanned image
# $imageProcess = new-object -ComObject WIA.ImageProcess
# $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
# $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = [string]$wiaFormatJpeg
# $imageProcess.Filters.Item(1).Properties.Item("Quality").Value = $quality
# $image = $imageProcess.Apply($image)


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


