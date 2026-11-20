# Local Lead Finder (OpenStreetMap)

Search businesses by keyword + city, view them in a dashboard, and export to CSV / Excel.
**No API key and no billing needed.** Uses Nominatim (city lookup) and Overpass (business search).

## Run it (Python 3.11.9)
```bash
python -m venv venv
venv\Scripts\activate          # Windows   (Mac/Linux: source venv/bin/activate)
pip install -r requirements.txt
copy .env.example .env         # Mac/Linux: cp .env.example .env  -> then put your email in it
python app.py
```
Open http://127.0.0.1:5000

## Notes
- **Fields:** name, address, phone, website, email, category, OSM ID (the OSM equivalent of a Place ID), OSM link.
- **Rating:** OpenStreetMap does not have ratings, so this column does not exist.
- **Coverage:** phone/website/address only appear where mappers added them. Use "Only with phone or website" to filter.
- **Pagination:** Overpass has no pages; use the "Max results" limit instead.
- **Duplicates:** removed by OSM ID and by name + phone/address.
- **License:** data is ODbL. Keep the "(c) OpenStreetMap contributors" credit (already in the footer and Excel export). You may store and export it, as long as you credit it.
- **Fair use:** the public servers are free but shared. Do not send many searches per minute, and set your email in `.env` (required by Nominatim's usage policy).
