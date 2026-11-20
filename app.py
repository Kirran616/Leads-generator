
"""
Lead Generation Tool - OpenStreetMap Edition
Flask application configured for Render deployment.

Data sources:
- Nominatim: Location search
- Photon: Backup geocoder
- Overpass: Business search

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

# Contact email for OpenStreetMap services
CONTACT = (
    os.getenv("CONTACT_EMAIL", "").strip()
    or "set-your-email-in-env-file"
)

HEADERS = {
    "User-Agent": f"SimpleLeadTool/1.0 ({CONTACT})"
}

NOMINATIM = "https://nominatim.openstreetmap.org/search"

PHOTON = "https://photon.komoot.io/api/"

OVERPASS_URLS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
]

FIELDS = [
    "name",
    "address",
    "phone",
    "website",
    "email",
    "category",
    "osm_id",
    "osm_url"
]

LABELS = [
    "Business Name",
    "Address",
    "Phone",
    "Website",
    "Email",
    "Category",
    "OSM ID",
    "OSM Link"
]

TAG_KEYS = [
    "amenity",
    "shop",
    "office",
    "craft",
    "tourism",
    "healthcare",
    "leisure"
]


# Custom API error
class ApiError(Exception):

    def __init__(self, message, status=502):
        super().__init__(message)
        self.status = status


# Location search using Nominatim
def _nominatim(location):

    r = requests.get(
        NOMINATIM,
        headers=HEADERS,
        timeout=15,
        params={
            "q": location,
            "format": "json",
            "limit": 1
        }
    )

    if r.status_code != 200:
        raise ApiError(
            f"Nominatim error {r.status_code}",
            r.status_code
        )

    data = r.json()

    if not data:
        return None

    s, n, w, e = data[0]["boundingbox"]

    return (
        f"{s},{w},{n},{e}",
        data[0]["display_name"]
    )


# Backup location search using Photon
def _photon(location):

    r = requests.get(
        PHOTON,
        headers=HEADERS,
        timeout=15,
        params={
            "q": location,
            "limit": 1
        }
    )

    if r.status_code != 200:
        raise ApiError(
            f"Photon error {r.status_code}",
            r.status_code
        )

    features = r.json().get("features", [])

    if not features:
        return None

    properties = features[0]["properties"]

    lon, lat = features[0]["geometry"]["coordinates"]

    if properties.get("extent"):

        w, n, e, s = properties["extent"]

    else:

        # Approximately 5 km around the location
        w = lon - 0.05
        e = lon + 0.05
        s = lat - 0.05
        n = lat + 0.05

    name = ", ".join(
        x for x in [
            properties.get("name"),
            properties.get("state"),
            properties.get("country")
        ]
        if x
    )

    return (
        f"{s},{w},{n},{e}",
        name
    )


# Geocoding with backup service
def geocode(location):

    found = None
    failed = False

    for lookup in (_nominatim, _photon):

        try:
            found = lookup(location)

        except (requests.RequestException, ApiError):
            failed = True
            continue

        if found:
            return found

    if failed:
        raise ApiError(
            "The location service refused the request. "
            "Set a real email in your .env file "
            "(CONTACT_EMAIL) and try again.",
            502
        )

    raise ApiError(
        f'No place found for "{location}". '
        'Try a city name like "Lahore, Pakistan".',
        404
    )


# Build OpenStreetMap Overpass query
def build_query(keyword, bbox, limit):

    kw = re.sub(
        r'[^\w\s&-]',
        '',
        keyword,
        flags=re.UNICODE
    ).strip()

    # Escape special characters for Overpass regex
    escaped = re.escape(kw)

    pattern = (
        f"{escaped}|{re.escape(kw.replace(' ', '_'))}"
        if " " in kw
        else escaped
    )

    lines = [
        f'nwr["name"~"{pattern}",i]({bbox});'
    ]

    lines += [
        f'nwr["{key}"~"{pattern}",i]["name"]({bbox});'
        for key in TAG_KEYS
    ]

    return (
        f"[out:json][timeout:50];"
        f"({''.join(lines)});"
        f"out center tags {limit};"
    )


# Run Overpass search with backup server
def run_overpass(query):

    last = "Business search failed."

    for url in OVERPASS_URLS:

        try:

            r = requests.post(
                url,
                data={"data": query},
                headers=HEADERS,
                timeout=70
            )

        except requests.RequestException:

            last = (
                "Business search timed out. "
                "Try a smaller area or specific keyword."
            )

            continue

        if r.status_code == 200:

            try:
                return r.json().get("elements", [])

            except ValueError:
                last = "Invalid response from OpenStreetMap."
                continue

        if r.status_code == 429:

            last = (
                "The free OpenStreetMap server is busy. "
                "Wait a minute and try again."
            )

        elif r.status_code in (504, 408):

            last = (
                "The search took too long. "
                "Try a smaller area or specific keyword."
            )

        else:

            last = (
                f"Business search returned an error "
                f"({r.status_code})."
            )

    raise ApiError(last)


# Convert OSM result into lead
def to_lead(element):

    tags = element.get("tags", {})

    parts = [
        tags.get("addr:housenumber", ""),
        tags.get("addr:street", ""),
        tags.get("addr:suburb", ""),
        tags.get("addr:city", ""),
        tags.get("addr:postcode", "")
    ]

    address = ", ".join(
        part for part in parts if part
    )

    category = next(
        (
            tags[key].replace("_", " ")
            for key in TAG_KEYS
            if key in tags
        ),
        ""
    )

    oid = f'{element["type"]}/{element["id"]}'

    return {

        "name": tags.get("name", ""),

        "address": address,

        "phone": (
            tags.get("phone")
            or tags.get("contact:phone")
            or tags.get("contact:mobile")
            or ""
        ),

        "website": (
            tags.get("website")
            or tags.get("contact:website")
            or ""
        ),

        "email": (
            tags.get("email")
            or tags.get("contact:email")
            or ""
        ),

        "category": category,

        "osm_id": oid,

        "osm_url": (
            f"https://www.openstreetmap.org/{oid}"
        )
    }


# Remove duplicate businesses
def dedupe(leads):

    seen = set()
    output = []

    for lead in leads:

        key = (
            lead["name"].lower().strip(),
            lead["phone"] or lead["address"].lower()
        )

        if lead["osm_id"] in seen or key in seen:
            continue

        seen.add(lead["osm_id"])
        seen.add(key)

        output.append(lead)

    return output


# Home page
@app.route("/")
def index():

    return render_template("index.html")


# Business search API
@app.route("/api/search")
def search():

    keyword = request.args.get(
        "keyword",
        ""
    ).strip()

    location = request.args.get(
        "location",
        ""
    ).strip()

    try:

        limit = min(
            max(
                int(request.args.get("limit", 100)),
                10
            ),
            500
        )

    except ValueError:

        limit = 100

    if not keyword or not location:

        return jsonify(
            error="Enter both a business type and a location."
        ), 400

    try:

        bbox, place = geocode(location)

        elements = run_overpass(
            build_query(keyword, bbox, limit)
        )

    except ApiError as error:

        return jsonify(
            error=str(error)
        ), error.status

    leads = dedupe([
        to_lead(element)
        for element in elements
        if element.get("tags", {}).get("name")
    ])

    return jsonify(
        results=leads,
        total=len(leads),
        place=place
    )


# Prepare export rows
def _rows(leads):

    return [
        [lead.get(field, "") for field in FIELDS]
        for lead in leads
    ]


# Export CSV and Excel files
@app.route(
    "/api/export/<fmt>",
    methods=["POST"]
)
def export(fmt):

    leads = (
        request.get_json(silent=True) or {}
    ).get("results", [])

    if not leads:

        return jsonify(
            error="Nothing to export."
        ), 400

    if fmt == "csv":

        buffer = io.StringIO()

        writer = csv.writer(buffer)

        writer.writerow(LABELS)

        writer.writerows(_rows(leads))

        data = io.BytesIO(
            (
                "\ufeff" + buffer.getvalue()
            ).encode("utf-8")
        )

        return send_file(
            data,
            mimetype="text/csv",
            as_attachment=True,
            download_name="leads.csv"
        )

    if fmt == "xlsx":

        workbook = Workbook()

        worksheet = workbook.active

        worksheet.title = "Leads"

        worksheet.append(LABELS)

        for row in _rows(leads):

            worksheet.append(row)

        worksheet.append([])

        worksheet.append([
            "Data (c) OpenStreetMap contributors, "
            "ODbL - openstreetmap.org/copyright"
        ])

        widths = [
            32, 45, 18, 32, 28, 18, 16, 38
        ]

        for column, width in zip(
            "ABCDEFGH",
            widths
        ):

            worksheet.column_dimensions[
                column
            ].width = width

        data = io.BytesIO()

        workbook.save(data)

        data.seek(0)

        return send_file(
            data,
            as_attachment=True,
            download_name="leads.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    return jsonify(
        error="Unknown format."
    ), 404


# Health check
@app.route("/health")
def health():

    return jsonify(
        status="ok",
        message="Lead Generation Tool is running."
    )


# Render deployment configuration
if __name__ == "__main__":

    port = int(
        os.environ.get("PORT", 10000)
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
