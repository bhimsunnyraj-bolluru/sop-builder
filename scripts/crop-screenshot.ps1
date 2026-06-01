# Crops a PNG to the given screen rectangle (SAP window bounds).
param(
    [Parameter(Mandatory = $true)][string]$ImagePath,
    [Parameter(Mandatory = $true)][string]$OutPath,
    [int]$Left = 0,
    [int]$Top = 0,
    [int]$Right = 0,
    [int]$Bottom = 0
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ImagePath)) {
    Write-Error "Image not found: $ImagePath"
    exit 1
}

if ($Right -le $Left -or $Bottom -le $Top) {
    Copy-Item -LiteralPath $ImagePath -Destination $OutPath -Force
    exit 0
}

Add-Type -AssemblyName System.Drawing

$bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
try {
    $imgW = $bmp.Width
    $imgH = $bmp.Height

    $cropLeft = [Math]::Max(0, [Math]::Min($Left, $imgW - 1))
    $cropTop = [Math]::Max(0, [Math]::Min($Top, $imgH - 1))
    $cropRight = [Math]::Max($cropLeft + 1, [Math]::Min($Right, $imgW))
    $cropBottom = [Math]::Max($cropTop + 1, [Math]::Min($Bottom, $imgH))
    $w = $cropRight - $cropLeft
    $h = $cropBottom - $cropTop

    $crop = New-Object System.Drawing.Bitmap $w, $h
    try {
        $g = [System.Drawing.Graphics]::FromImage($crop)
        try {
            $srcRect = New-Object System.Drawing.Rectangle $cropLeft, $cropTop, $w, $h
            $destRect = New-Object System.Drawing.Rectangle 0, 0, $w, $h
            $g.DrawImage($bmp, $destRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
        } finally {
            $g.Dispose()
        }
        $outDir = Split-Path -Parent $OutPath
        if ($outDir -and -not (Test-Path $outDir)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        }
        $crop.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $crop.Dispose()
    }
} finally {
    $bmp.Dispose()
}
