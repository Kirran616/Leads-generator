"""Lead Generation Tool - OpenStreetMap edition (no API key required).

Data sources:
  * Nominatim  -> turns a city name into a bounding box
  * Overpass   -> finds businesses (OSM nodes/ways/relations) inside that box
Data (c) OpenStreetMap contributors, licensed under ODbL.
"""
import csv
import io
import os
import re

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook

load_dotenv()
app = Flask(__name__)

CONTACT = os.getenv("CONTACT_EMAIL", "").strip() or "set-your-email-in-env-file"
HEADERS = {"User-Agent": f"SimpleLeadTool/1.0 ({CONTACT})"}
NOMINATIM = "https://nominatim.openstreetmap.org/search"
OVERPASS_URLS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
]
FIELDS = ["name", "address", "phone", "website", "email", "category", "osm_id", "osm_url"]
LABELS = ["Business Name", "Address", "Phone", "Website", "Email", "Category", "OSM ID", "OSM Link"]
TAG_KEYS = ["amenity", "shop", "office", "craft", "tourism", "healthcare", "leisure"]


class ApiError(Exception):
    def __init__(self, message, status=502):
        super().__init__(message)
        self.status = status


def _nominatim(location):
    r = requests.get(NOMINATIM, headers=HEADERS, timeout=15,
                     params={"q": location, "format": "json", "limit": 1})
    if r.status_code != 200:
        raise ApiError(f"Nominatim error {r.status_code}", r.status_code)
    data = r.json()
    if not data:
        return None
    s, n, w, e = data[0]["boundingbox"]  # south, north, west, east
    return f"{s},{w},{n},{e}", data[0]["display_name"]


def _photon(location):
    """Backup geocoder (also OpenStreetMap data) used if Nominatim blocks us."""
    r = requests.get("https://photon.komoot.io/api/", headers=HEADERS, timeout=15,
                     params={"q": location, "limit": 1})
    if r.status_code != 200:
        raise ApiError(f"Photon error {r.status_code}", r.status_code)
    feats = r.json().get("features", [])
    if not feats:
        return None
    p, (lon, lat) = feats[0]["properties"], feats[0]["geometry"]["coordinates"]
    if p.get("extent"):
        w, n, e, s = p["extent"]
    else:  # a single point: search about 5 km around it
        w, e, s, n = lon - 0.05, lon + 0.05, lat - 0.05, lat + 0.05
    name = ", ".join(x for x in [p.get("name"), p.get("state"), p.get("country")] if x)
    return f"{s},{w},{n},{e}", name


def geocode(location):
    found, failed = None, False
    for lookup in (_nominatim, _photon):
        try:
            found = lookup(location)
        except (requests.RequestException, ApiError):
            failed = True
            continue
        if found:
            return found
    if failed:
        raise ApiError("The location service refused the request. Set a real email in your .env "
                       "file (CONTACT_EMAIL), restart the app, and try again.", 502)
    raise ApiError(f'No place found for "{location}". Try a city name like "Lahore, Pakistan".', 404)


def build_query(keyword, bbox, limit):
    kw = re.sub(r"[^\w\s&-]", "", keyword, flags=re.UNICODE).strip()
    pattern = f"{kw}|{kw.replace(' ', '_')}" if " " in kw else kw
    lines = [f'nwr["name"~"{pattern}",i]({bbox});']
    lines += [f'nwr["{k}"~"{pattern}",i]["name"]({bbox});' for k in TAG_KEYS]
    return f"[out:json][timeout:50];({''.join(lines)});out center tags {limit};"


def run_overpass(query):
    last = "Business search failed."
    for url in OVERPASS_URLS:
        try:
            r = requests.post(url, data={"data": query}, headers=HEADERS, timeout=70)
        except requests.RequestException:
            last = "Business search timed out. Try a smaller area or a more specific keyword."
            continue
        if r.status_code == 200:
            return r.json().get("elements", [])
        if r.status_code == 429:
            last = "The free OpenStreetMap server is busy. Wait a minute and try again."
        elif r.status_code in (504, 408):
            last = "The search took too long. Try a smaller area or a more specific keyword."
        else:
            last = f"Business search returned an error ({r.status_code})."
    raise ApiError(last)


def to_lead(el):
    t = el.get("tags", {})
    parts = [t.get("addr:housenumber", ""), t.get("addr:street", ""), t.get("addr:suburb", ""),
             t.get("addr:city", ""), t.get("addr:postcode", "")]
    address = ", ".join(p for p in parts if p)
    category = next((t[k].replace("_", " ") for k in TAG_KEYS if k in t), "")
    oid = f'{el["type"]}/{el["id"]}'
    return {
        "name": t.get("name", ""),
        "address": address,
        "phone": t.get("phone") or t.get("contact:phone") or t.get("contact:mobile") or "",
        "website": t.get("website") or t.get("contact:website") or "",
        "email": t.get("email") or t.get("contact:email") or "",
        "category": category,
        "osm_id": oid,
        "osm_url": f"https://www.openstreetmap.org/{oid}",
    }


def dedupe(leads):
    seen, out = set(), []
    for l in leads:
        key = (l["name"].lower().strip(), l["phone"] or l["address"].lower())
        if l["osm_id"] in seen or key in seen:
            continue
        seen.update([l["osm_id"], key])
        out.append(l)
    return out


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/search")
def search():
    keyword = request.args.get("keyword", "").strip()
    location = request.args.get("location", "").strip()
    try:
        limit = min(max(int(request.args.get("limit", 100)), 10), 500)
    except ValueError:
        limit = 100
    if not keyword or not location:
        return jsonify(error="Enter both a business type and a location."), 400
    try:
        bbox, place = geocode(location)
        elements = run_overpass(build_query(keyword, bbox, limit))
    except ApiError as e:
        return jsonify(error=str(e)), e.status
    leads = dedupe([to_lead(el) for el in elements if el.get("tags", {}).get("name")])
    return jsonify(results=leads, total=len(leads), place=place)


def _rows(leads):
    return [[l.get(f, "") for f in FIELDS] for l in leads]


@app.route("/api/export/<fmt>", methods=["POST"])
def export(fmt):
    leads = (request.get_json(silent=True) or {}).get("results", [])
    if not leads:
        return jsonify(error="Nothing to export."), 400
    if fmt == "csv":
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(LABELS)
        w.writerows(_rows(leads))
        data = io.BytesIO(("\ufeff" + buf.getvalue()).encode("utf-8"))  # BOM so Excel reads UTF-8
        return send_file(data, mimetype="text/csv", as_attachment=True, download_name="leads.csv")
    if fmt == "xlsx":
        wb = Workbook()
        ws = wb.active
        ws.title = "Leads"
        ws.append(LABELS)
        for row in _rows(leads):
            ws.append(row)
        ws.append([])
        ws.append(["Data (c) OpenStreetMap contributors, ODbL - openstreetmap.org/copyright"])
        for col, width in zip("ABCDEFGH", [32, 45, 18, 32, 28, 18, 16, 38]):
            ws.column_dimensions[col].width = width
        data = io.BytesIO()
        wb.save(data)
        data.seek(0)
        return send_file(data, as_attachment=True, download_name="leads.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return jsonify(error="Unknown format."), 404


if __name__ == "__main__":
    app.run(debug=True)
