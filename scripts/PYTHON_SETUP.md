# Python Setup voor 3BM pyRevit Development

## Status
✅ Python 3.14 is geïnstalleerd  
✅ Python staat correct in Windows PATH  
✅ Scripts zijn beschikbaar

## Locatie
- Python: `C:\Users\JochemK\AppData\Local\Programs\Python\Python314`
- Scripts: `C:\Users\JochemK\AppData\Local\Programs\Python\Python314\Scripts`

## Gebruik

### Optie 1: Nieuwe PowerShell sessie (aanbevolen)
Start een nieuwe PowerShell sessie - Python werkt dan direct:
```powershell
python --version
pip --version
```

### Optie 2: Enable Python in huidige sessie
Als je in een oude PowerShell sessie zit:
```powershell
.\enable_python.ps1
```

### Optie 3: Volledig pad gebruiken
```powershell
C:\Users\JochemK\AppData\Local\Programs\Python\Python314\python.exe --version
```

## Icon Generator Scripts

### PowerShell versie (altijd beschikbaar)
```powershell
.\create_toolbar_icons.ps1
```
Gebruikt .NET System.Drawing - altijd beschikbaar

### Python versie (vereist PIL/Pillow)
```powershell
python create_toolbar_icons.py
```
Vereist: `pip install Pillow`

## Packages installeren
```powershell
pip install Pillow          # Voor icon generatie
pip install openpyxl        # Voor Excel bestanden
```

## Troubleshooting

**"python is not recognized"**
- Start nieuwe PowerShell sessie
- Of run: `.\enable_python.ps1`

**"No module named PIL"**
```powershell
pip install Pillow
```

**Welke Python versie?**
```powershell
python --version
```
