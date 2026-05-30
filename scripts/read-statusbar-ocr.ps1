# Crops the SAP status bar band from a screenshot and reads text via Windows OCR.
param(
    [Parameter(Mandatory = $true)][string]$ImagePath,
    [int]$Left = 0,
    [int]$Top = 0,
    [int]$Right = 0,
    [int]$Bottom = 0
)

$ErrorActionPreference = 'Stop'

function Await-Task {
    param($Task, [Type]$ResultType)
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1
    } | Select-Object -First 1
    if (-not $method) { throw 'AsTask extension not available' }
    $asTask = $method.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($Task))
    return $netTask.GetAwaiter().GetResult()
}

if (-not (Test-Path -LiteralPath $ImagePath)) {
    Write-Output ""
    exit 0
}

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType=WindowsRuntime]

$bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
try {
    $imgW = $bmp.Width
    $imgH = $bmp.Height

    if ($Right -gt $Left -and $Bottom -gt $Top) {
        $cropLeft = [Math]::Max(0, [Math]::Min($Left, $imgW - 1))
        $cropTop = [Math]::Max(0, [Math]::Min($Top, $imgH - 1))
        $cropRight = [Math]::Max($cropLeft + 1, [Math]::Min($Right, $imgW))
        $cropBottom = [Math]::Max($cropTop + 1, [Math]::Min($Bottom, $imgH))
        $regionH = $cropBottom - $cropTop
        $barH = [Math]::Max(28, [Math]::Min(72, [int][Math]::Round($regionH * 0.09)))
        $x = $cropLeft
        $y = [Math]::Max($cropTop, $cropBottom - $barH)
        $w = $cropRight - $cropLeft
        $h = [Math]::Min($barH, $cropBottom - $y)
    } else {
        $h = [Math]::Max(28, [Math]::Min(72, [int][Math]::Round($imgH * 0.06)))
        $x = 0
        $y = $imgH - $h
        $w = $imgW
    }

    if ($w -lt 10 -or $h -lt 10) {
        Write-Output ""
        exit 0
    }

    $cropRect = New-Object System.Drawing.Rectangle $x, $y, $w, $h
    $cropped = $bmp.Clone($cropRect, $bmp.PixelFormat)
    $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "sop-statusbar-" + [guid]::NewGuid().ToString() + ".png")
    $cropped.Save($tempFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $cropped.Dispose()

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new("en"))
    }
    if ($null -eq $engine) {
        Write-Output ""
        exit 0
    }

    $file = Await-Task ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tempFile)) ([Windows.Storage.StorageFile])
    $stream = Await-Task ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await-Task ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $sbmp = Await-Task ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $result = Await-Task ($engine.RecognizeAsync($sbmp)) ([Windows.Media.Ocr.OcrResult])

    $text = ($result.Lines | ForEach-Object { $_.Text }) -join " "
    $text = ($text -replace '\s+', ' ').Trim()
    Write-Output $text
}
finally {
    if ($bmp) { $bmp.Dispose() }
    if ($tempFile -and (Test-Path -LiteralPath $tempFile)) {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }
}
