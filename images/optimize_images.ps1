
Add-Type -AssemblyName System.Drawing

$details = @{
    quality = 75
    maxWidth = 1920
}

$encoder = [System.Drawing.Imaging.Encoder]::Quality
$encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
$encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($encoder, $details.quality)
$jpegCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }

$files = Get-ChildItem "presentation Slide*.png"

foreach ($file in $files) {
    Write-Host "Processing $($file.Name)..."
    $img = [System.Drawing.Image]::FromFile($file.FullName)
    
    # Calculate new dimensions
    $newWidth = $img.Width
    $newHeight = $img.Height
    
    if ($img.Width -gt $details.maxWidth) {
        $newWidth = $details.maxWidth
        $newHeight = [int]($img.Height * ($details.maxWidth / $img.Width))
    }
    
    # Resize
    $newImg = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graph = [System.Drawing.Graphics]::FromImage($newImg)
    $graph.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graph.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graph.DrawImage($img, 0, 0, $newWidth, $newHeight)
    
    # Save as JPG
    $newName = $file.FullName -replace '\.png$', '.jpg'
    $newImg.Save($newName, $jpegCodec, $encoderParams)
    
    # Cleanup
    $img.Dispose()
    $newImg.Dispose()
    $graph.Dispose()
    
    Write-Host "Saved to $newName"
}
