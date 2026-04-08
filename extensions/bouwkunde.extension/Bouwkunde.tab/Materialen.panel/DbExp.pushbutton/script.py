# -*- coding: utf-8 -*-
"""Export materialen database JSON naar CSV"""

__title__ = "Db\nExp"
__author__ = "3BM Bouwkunde"
__doc__ = "Exporteer materialen_database.json naar CSV in lib folder"

from pyrevit import forms
import os
import sys
import json

sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from bm_logger import get_logger

log = get_logger("DbExp")


def main():
    lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
    json_path = os.path.join(lib_path, 'materialen_database.json')
    csv_path = os.path.join(lib_path, 'materialen_database.csv')
    
    if not os.path.exists(json_path):
        forms.alert("Database niet gevonden:\n{}".format(json_path), warn_icon=True)
        return
    
    # Laad JSON
    with open(json_path, 'r') as f:
        data = json.load(f)
    
    materialen = data.get('materialen', [])
    versie = data.get('versie', '?')
    
    if not materialen:
        forms.alert("Geen materialen in database.", warn_icon=True)
        return
    
    # Schrijf CSV
    with open(csv_path, 'w') as f:
        f.write('\xef\xbb\xbf')  # UTF-8 BOM
        f.write("Categorie;Naam;Lambda (W/mK);Rd vast (m2K/W);Mu (-);Rho (kg/m3);Keywords\n")
        
        for mat in materialen:
            cat = mat.get('categorie', '')
            naam = mat.get('naam', '').replace(';', ',')
            lam = mat.get('lambda')
            rd = mat.get('rd_vast')
            mu = mat.get('mu')
            rho = mat.get('rho')
            kw = ', '.join(mat.get('keywords', []))
            
            lam_str = str(lam).replace('.', ',') if lam is not None else ''
            rd_str = str(rd).replace('.', ',') if rd is not None else ''
            
            f.write(u"{};{};{};{};{};{};{}\n".format(
                cat, naam, lam_str, rd_str,
                mu if mu is not None else '',
                rho if rho is not None else '',
                kw
            ))
    
    forms.alert(
        "Database geexporteerd!\n\n"
        "{} materialen (v{})\n\n"
        "Locatie:\n{}".format(len(materialen), versie, csv_path),
        title="Database Export"
    )
    os.startfile(csv_path)


if __name__ == '__main__':
    main()
