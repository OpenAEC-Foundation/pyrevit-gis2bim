import requests
import time
import os
import json

API = "https://api.pdok.nl/lv/bgt/download/v1_0"

# Complete lijst featuretypes, exact uit PDOK schema
BGT_FEATURES = [
    "bak",
    "begroeidterreindeel",
    "bord",
    "buurt",
    "functioneelgebied",
    "gebouwinstallatie",
    "installatie",
    "kast",
    "kunstwerkdeel",
    "mast",
    "onbegroeidterreindeel",
    "ondersteunendwaterdeel",
    "ondersteunendwegdeel",
    "ongeclassificeerdobject",
    "openbareruimte",
    "openbareruimtelabel",
    "overbruggingsdeel",
    "overigbouwwerk",
    "overigescheiding",
    "paal",
    "pand",
    "plaatsbepalingspunt",
    "put",
    "scheiding",
    "sensor",
    "spoor",
    "stadsdeel",
    "straatmeubilair",
    "tunneldeel",
    "vegetatieobject",
    "waterdeel",
    "waterinrichtingselement",
    "waterschap",
    "wegdeel",
    "weginrichtingselement",
    "wijk"
]


def bbox_to_wkt(b):
    xmin, ymin, xmax, ymax = b
    return (
        f"POLYGON (({xmin} {ymin}, "
        f"{xmax} {ymin}, "
        f"{xmax} {ymax}, "
        f"{xmin} {ymax}, "
        f"{xmin} {ymin}))"
    )


def start_bgt_download(bbox):
    url = f"{API}/full/custom"

    payload = {
        "featuretypes": BGT_FEATURES,
        "format": "citygml",       # Correcte enum, exact volgens PDOK
        "geofilter": bbox_to_wkt(bbox)
    }

    r = requests.post(url, json=payload)
    print("START RESPONSE:", r.text)
    r.raise_for_status()

    return r.json()["downloadRequestId"]


def poll_bgt(download_id):
    url = f"{API}/full/custom/{download_id}/status"

    while True:
        r = requests.get(url)
        r.raise_for_status()
        data = r.json()

        status = data.get("status")
        print("STATUS:", status)

        if status == "COMPLETED":
            return data["_links"]["download"]["href"]

        if status == "ERROR":
            raise Exception("BGT download error: " + json.dumps(data, indent=2))

        time.sleep(2)


def download_bgt(download_url, outdir, download_id):
    full_url = f"{API}{download_url}"

    r = requests.get(full_url)
    r.raise_for_status()

    os.makedirs(outdir, exist_ok=True)
    outfile = os.path.join(outdir, f"bgt_{download_id}.zip")

    with open(outfile, "wb") as f:
        f.write(r.content)

    return outfile


def run_bgt(bbox, outdir):
    download_id = start_bgt_download(bbox)
    dl_url = poll_bgt(download_id)
    zipfile = download_bgt(dl_url, outdir, download_id)

    return {
        "status": "OK",
        "download_id": download_id,
        "zipfile": zipfile
    }
