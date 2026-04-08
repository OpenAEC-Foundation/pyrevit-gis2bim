# Helper script om Python te gebruiken in huidige sessie
# Ook als PATH nog niet herladen is

$pythonPath = "C:\Users\JochemK\AppData\Local\Programs\Python\Python314"

# Voeg toe aan PATH van deze sessie
if ($env:Path -notlike "*$pythonPath*") {
    $env:Path = "$pythonPath;$pythonPath\Scripts;" + $env:Path
    Write-Host "Python toegevoegd aan sessie PATH"
}

# Test
Write-Host "`nPython versie:"
python --version

Write-Host "`nPip versie:"
pip --version

Write-Host "`n✓ Python is nu beschikbaar in deze sessie"
Write-Host "  Gebruik: python <script.py>"
Write-Host "  Gebruik: pip install <package>"
