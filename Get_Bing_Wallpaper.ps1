#---------------------------------------------
# Get_Bing_Wallpaper.ps1 
# Download Bing Wallpapers in resolution 1920x1080 (16:9) and 1920x1200 (16:10) 
# Created by Borcon
#---------------------------------------------

# Use the Bing.com API. 
# The idx parameter determines the day: 0 is the current day, 1 is the previous day, etc. This goes back for max. 7 days. 
# The n parameter defines how many pictures you want to load. Usually, n=1 to get the latest picture (of today) only. 
# The mkt parameter defines the culture, like en-US, de-DE, etc.
$url                = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US"
$idx                = "7"
$downloadBaseFolder = "$([Environment]::GetFolderPath("MyPictures"))\Wallpapers\Bing"

# Check if download folders exists and otherwise create it
$downloadFolder_1920x1080 = $downloadBaseFolder+"\1920x1080\"
$downloadFolder_1920x1200 = $downloadBaseFolder+"\1920x1200\"
if (!(Test-Path $downloadBaseFolder)) {
    New-Item -ItemType Directory $downloadBaseFolder
}
if (!(Test-Path $downloadFolder_1920x1080)) {
    New-Item -ItemType Directory $downloadFolder_1920x1080
}
if (!(Test-Path $downloadFolder_1920x1200)) {
    New-Item -ItemType Directory $downloadFolder_1920x1200
}

$i = 0
for($i;$i -lt $idx;$i++) 

{
    # Create Download URL
    $url_New = $url.Replace("idx=0","idx=$i")

    # Get the picture metadata
    $response = Invoke-WebRequest -Method Get -Uri $url_New

    # Extract the image content
    $body = ConvertFrom-Json -InputObject $response.Content
    # $fileurl_16_9       = "https://www.bing.com/"+$body.images[0].url
    # $fileurl_16_10      = $fileurl_16_9.Replace("1080","1200")
    $fileurl_16_9       = "https://www.bing.com/"+$body.images[0].urlbase+"_1920x1080.jpg"
    $fileurl_16_10      = "https://www.bing.com/"+$body.images[0].urlbase+"_1920x1200.jpg"
    # $filename_title     = $body.images[0].title.Replace(" ","-").Replace("?","")+".jpg"
    $filename_copyright = $body.images[0].copyright
    $filename_copyright = $filename_copyright -replace '(.+)\s\(.*\)','$1'
    $filename_copyright = $filename_copyright+".jpg"

    # Download the picture
    $filepath_16_9  = $downloadFolder_1920x1080+$filename_copyright
    $filepath_16_10 = $downloadFolder_1920x1200+$filename_copyright
    if (!(Test-Path $filepath_16_9)) {
        Write-Host "Download Bing Image 16:9" $filename_copyright
        Write-Host "  --> Path:" $filepath_16_9
        Write-Host ""
        Invoke-WebRequest -Method Get -Uri $fileurl_16_9 -OutFile $filepath_16_9
    }
    if (!(Test-Path $filepath_16_10)) {
        Write-Host "Download Bing Image 16:10" $filename_copyright
        Write-Host "  --> Path:" $filepath_16_10
        Write-Host ""
        Invoke-WebRequest -Method Get -Uri $fileurl_16_10 -OutFile $filepath_16_10
    }
}
