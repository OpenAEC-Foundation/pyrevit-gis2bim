# PowerShell script om PNG icons te genereren met System.Drawing
Add-Type -AssemblyName System.Drawing

function Create-ScheduleExportIcon {
    param([string]$Path)
    
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    # Achtergrond
    $violet = [System.Drawing.Color]::FromArgb(255, 53, 14, 53)
    $g.Clear($violet)
    
    # Teal kleur
    $teal = [System.Drawing.Color]::FromArgb(255, 69, 182, 168)
    $penTeal = New-Object System.Drawing.Pen($teal, 2)
    $penThin = New-Object System.Drawing.Pen($teal, 1)
    
    # Spreadsheet
    $g.DrawRectangle($penTeal, 6, 6, 14, 16)
    
    # Grid lijnen horizontaal
    $g.DrawLine($penThin, 6, 10, 20, 10)
    $g.DrawLine($penThin, 6, 14, 20, 14)
    $g.DrawLine($penThin, 6, 18, 20, 18)
    
    # Grid lijnen verticaal
    $g.DrawLine($penThin, 11, 6, 11, 22)
    $g.DrawLine($penThin, 15, 6, 15, 22)
    
    # Export pijl naar rechts
    $g.DrawLine($penTeal, 21, 14, 27, 14)
    $points = @(
        (New-Object System.Drawing.Point(27, 14)),
        (New-Object System.Drawing.Point(24, 11)),
        (New-Object System.Drawing.Point(24, 17))
    )
    $brushTeal = New-Object System.Drawing.SolidBrush($teal)
    $g.FillPolygon($brushTeal, $points)
    
    $bmp.Save($Path)
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Created: $Path"
}

function Create-ScheduleImportIcon {
    param([string]$Path)
    
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    # Achtergrond
    $violet = [System.Drawing.Color]::FromArgb(255, 53, 14, 53)
    $g.Clear($violet)
    
    # Teal kleur
    $teal = [System.Drawing.Color]::FromArgb(255, 69, 182, 168)
    $penTeal = New-Object System.Drawing.Pen($teal, 2)
    $penThin = New-Object System.Drawing.Pen($teal, 1)
    
    # Spreadsheet
    $g.DrawRectangle($penTeal, 12, 6, 14, 16)
    
    # Grid lijnen horizontaal
    $g.DrawLine($penThin, 12, 10, 26, 10)
    $g.DrawLine($penThin, 12, 14, 26, 14)
    $g.DrawLine($penThin, 12, 18, 26, 18)
    
    # Grid lijnen verticaal
    $g.DrawLine($penThin, 17, 6, 17, 22)
    $g.DrawLine($penThin, 21, 6, 21, 22)
    
    # Import pijl naar links
    $g.DrawLine($penTeal, 5, 14, 11, 14)
    $points = @(
        (New-Object System.Drawing.Point(5, 14)),
        (New-Object System.Drawing.Point(8, 11)),
        (New-Object System.Drawing.Point(8, 17))
    )
    $brushTeal = New-Object System.Drawing.SolidBrush($teal)
    $g.FillPolygon($brushTeal, $points)
    
    $bmp.Save($Path)
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Created: $Path"
}

function Create-SheetParametersIcon {
    param([string]$Path)
    
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    # Achtergrond
    $violet = [System.Drawing.Color]::FromArgb(255, 53, 14, 53)
    $g.Clear($violet)
    
    # Teal kleur
    $teal = [System.Drawing.Color]::FromArgb(255, 69, 182, 168)
    $tealLight = [System.Drawing.Color]::FromArgb(80, 69, 182, 168)
    $penTeal = New-Object System.Drawing.Pen($teal, 2)
    $penThin = New-Object System.Drawing.Pen($teal, 1)
    $brushTeal = New-Object System.Drawing.SolidBrush($teal)
    $brushTealLight = New-Object System.Drawing.SolidBrush($tealLight)
    
    # A0 Sheet outline
    $g.DrawRectangle($penTeal, 4, 7, 24, 18)
    
    # Titleblock
    $g.FillRectangle($brushTealLight, 4, 22, 24, 3)
    $g.DrawRectangle($penThin, 4, 22, 24, 3)
    
    # Sheet lines
    $penLight = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(128, 69, 182, 168), 1)
    $g.DrawLine($penLight, 8, 11, 18, 11)
    $g.DrawLine($penLight, 8, 14, 15, 14)
    $g.DrawLine($penLight, 8, 17, 20, 17)
    $g.DrawLine($penLight, 20, 11, 24, 11)
    $g.DrawLine($penLight, 20, 14, 24, 14)
    
    # Settings gear in titleblock
    $g.DrawEllipse($penThin, 24, 22, 3, 3)
    
    $bmp.Save($Path)
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Created: $Path"
}

function Create-NAAKTIcon {
    param([string]$Path)
    
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    # Achtergrond
    $violet = [System.Drawing.Color]::FromArgb(255, 53, 14, 53)
    $g.Clear($violet)
    
    # Teal kleuren
    $teal = [System.Drawing.Color]::FromArgb(255, 69, 182, 168)
    $teal80 = [System.Drawing.Color]::FromArgb(204, 69, 182, 168)
    $teal60 = [System.Drawing.Color]::FromArgb(153, 69, 182, 168)
    $teal50 = [System.Drawing.Color]::FromArgb(128, 69, 182, 168)
    
    $penTeal = New-Object System.Drawing.Pen($teal, 1)
    $brushTeal80 = New-Object System.Drawing.SolidBrush($teal80)
    $brushTeal50 = New-Object System.Drawing.SolidBrush($teal50)
    $brushTeal60 = New-Object System.Drawing.SolidBrush($teal60)
    
    # Materiaal lagen
    $g.FillRectangle($brushTeal80, 6, 8, 20, 3)
    $g.FillRectangle($brushTeal50, 6, 11, 20, 3)
    
    # Dashed lijn voor luchtlaag
    for ($x = 6; $x -lt 26; $x += 3) {
        $g.DrawLine($penTeal, $x, 14, [Math]::Min($x+1, 26), 14)
        $g.DrawLine($penTeal, $x, 15, [Math]::Min($x+1, 26), 15)
    }
    
    $g.FillRectangle($brushTeal60, 6, 16, 20, 3)
    
    # Dimensie pijltjes
    $g.DrawLine($penTeal, 4, 8, 4, 19)
    $g.DrawLine($penTeal, 3, 9, 4, 8)
    $g.DrawLine($penTeal, 5, 9, 4, 8)
    $g.DrawLine($penTeal, 3, 18, 4, 19)
    $g.DrawLine($penTeal, 5, 18, 4, 19)
    
    # Tekst
    $font = New-Object System.Drawing.Font("Segoe UI", 5, [System.Drawing.FontStyle]::Bold)
    $brushTeal = New-Object System.Drawing.SolidBrush($teal)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString("NAA.K.T.", $font, $brushTeal, 16, 23, $format)
    
    $bmp.Save($Path)
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Created: $Path"
}

# Main
$base = "X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\extensions\bouwkunde.extension\Bouwkunde.tab"

Create-ScheduleExportIcon (Join-Path $base "Data Exchange.panel\ScheduleExport.pushbutton\icon.png")
Create-ScheduleImportIcon (Join-Path $base "Data Exchange.panel\ScheduleImport.pushbutton\icon.png")
Create-SheetParametersIcon (Join-Path $base "Document.panel\SheetParameters.pushbutton\icon.png")
Create-NAAKTIcon (Join-Path $base "Materialen.panel\NAAKTGenerator.pushbutton\icon.png")

Write-Host "`nAll PNG icons created! Now syncing to runtime..."
