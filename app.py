"""
Google Maps Scraper — Web UI
Flask application with file upload, real-time progress via SSE.
"""

import os
import json
import time
import queue
import threading
from datetime import datetime
from flask import Flask, render_template, request, jsonify, Response, send_file
import requests as http_requests

from scraper import scrape_postcodes, create_excel, load_existing_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB max upload

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# Global state
jobs = {}
uploaded_data = {}  # session_id -> { stores: [...], keys: set(), filename: str }


def fetch_postcodes_for_location(location):
    results = []
    seen = set()

    try:
        resp = http_requests.get(
            f"https://api.postcodes.io/places?q={location}&limit=5", timeout=10
        )
        if resp.status_code == 200:
            data = resp.json()
            places = data.get("result", []) or []
            for place in places:
                lat = place.get("latitude")
                lon = place.get("longitude")
                if lat and lon:
                    resp2 = http_requests.get(
                        f"https://api.postcodes.io/outcodes?lon={lon}&lat={lat}&limit=100&radius=25000",
                        timeout=10,
                    )
                    if resp2.status_code == 200:
                        outcodes = resp2.json().get("result", []) or []
                        for oc in outcodes:
                            outcode = oc.get("outcode", "")
                            if outcode and outcode not in seen:
                                seen.add(outcode)
                                results.append({
                                    "outcode": outcode,
                                    "admin_district": ", ".join(oc.get("admin_district", []) or []),
                                    "latitude": oc.get("latitude"),
                                    "longitude": oc.get("longitude"),
                                })
    except Exception as e:
        print(f"Place search error: {e}")

    location_upper = location.strip().upper()
    if len(location_upper) <= 4 and location_upper[0].isalpha():
        try:
            resp = http_requests.get(
                f"https://api.postcodes.io/outcodes/{location_upper}", timeout=10
            )
            if resp.status_code == 200:
                oc = resp.json().get("result", {})
                outcode = oc.get("outcode", "")
                if outcode and outcode not in seen:
                    seen.add(outcode)
                    results.append({
                        "outcode": outcode,
                        "admin_district": ", ".join(oc.get("admin_district", []) or []),
                        "latitude": oc.get("latitude"),
                        "longitude": oc.get("longitude"),
                    })
        except Exception:
            pass

    def sort_key(x):
        oc = x["outcode"]
        alpha = "".join(c for c in oc if c.isalpha())
        num = "".join(c for c in oc if c.isdigit())
        return (alpha, int(num) if num else 0)

    results.sort(key=sort_key)
    return results


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/postcodes", methods=["POST"])
def get_postcodes():
    data = request.json
    location = data.get("location", "").strip()
    if not location:
        return jsonify({"error": "Location is required"}), 400
    postcodes = fetch_postcodes_for_location(location)
    return jsonify({"postcodes": postcodes, "location": location})


@app.route("/api/upload", methods=["POST"])
def upload_file():
    """Upload an existing Excel file. Parse it and return store count + session ID."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".xlsx", ".xls"):
        return jsonify({"error": "Only .xlsx or .xls files are supported"}), 400

    # Save to uploads dir
    session_id = f"upload_{int(time.time() * 1000)}"
    save_path = os.path.join(UPLOAD_DIR, f"{session_id}{ext}")
    file.save(save_path)

    # Parse the file
    stores, keys, error = load_existing_excel(save_path)

    if error:
        os.remove(save_path)
        return jsonify({"error": f"Could not parse file: {error}"}), 400

    # Extract unique postcodes found in the file
    existing_postcodes = sorted(set(
        s.get("postcode", "—") for s in stores if s.get("postcode", "—") != "—"
    ))

    # Store in memory for later use
    uploaded_data[session_id] = {
        "stores": stores,
        "keys": keys,
        "filename": file.filename,
        "filepath": save_path,
    }

    # Return summary (don't send full store data to keep response light)
    with_phone = sum(1 for s in stores if s.get("phone", "N/A") != "N/A")
    sample = [
        {"name": s["name"], "address": s["address"], "phone": s["phone"]}
        for s in stores[:5]
    ]

    return jsonify({
        "session_id": session_id,
        "filename": file.filename,
        "total_stores": len(stores),
        "with_phone": with_phone,
        "postcodes_found": existing_postcodes,
        "sample": sample,
    })


@app.route("/api/upload/remove", methods=["POST"])
def remove_upload():
    """Remove an uploaded file from memory."""
    data = request.json
    session_id = data.get("session_id", "")
    if session_id in uploaded_data:
        filepath = uploaded_data[session_id].get("filepath")
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
        del uploaded_data[session_id]
    return jsonify({"ok": True})


@app.route("/api/scrape", methods=["POST"])
def start_scrape():
    """Start a scraping job, optionally merging with uploaded data."""
    data = request.json
    query = data.get("query", "").strip()
    location = data.get("location", "").strip()
    postcodes = data.get("postcodes", [])
    session_id = data.get("upload_session_id", "")

    if not query:
        return jsonify({"error": "Search query is required"}), 400
    if not postcodes:
        return jsonify({"error": "Select at least one postcode"}), 400

    # Load existing data if an upload was provided
    existing_stores = []
    existing_keys = set()
    if session_id and session_id in uploaded_data:
        existing_stores = uploaded_data[session_id]["stores"]
        existing_keys = uploaded_data[session_id]["keys"]

    job_id = f"job_{int(time.time() * 1000)}"
    progress_queue = queue.Queue()

    jobs[job_id] = {
        "status": "running",
        "queue": progress_queue,
        "output_file": None,
        "started_at": datetime.now().isoformat(),
    }

    if existing_stores:
        progress_queue.put({
            "type": "progress",
            "message": f"Loaded {len(existing_stores)} existing stores from uploaded file — these will be skipped during scraping",
        })

    def run_scraper():
        try:
            output_file = f"output_{job_id}.xlsx"
            output_path = os.path.join(STATIC_DIR, output_file)

            new_results, postcode_summary = scrape_postcodes(
                query=query,
                location=location,
                postcodes=postcodes,
                existing_keys=existing_keys,
                progress_callback=lambda msg: progress_queue.put({"type": "progress", "message": msg}),
                store_callback=lambda store: progress_queue.put({"type": "store", "data": store}),
            )

            # Always create output (even if 0 new — existing stores will be in the file)
            create_excel(
                all_results=new_results,
                postcode_summary=postcode_summary,
                output_path=output_path,
                query=query,
                existing_stores=existing_stores,
            )

            jobs[job_id]["output_file"] = output_path
            total_combined = len(existing_stores) + len(new_results)

            progress_queue.put({
                "type": "complete",
                "total_stores": total_combined,
                "new_stores": len(new_results),
                "existing_stores": len(existing_stores),
                "total_with_phone": (
                    len([s for s in new_results if s.get("phone", "N/A") != "N/A"]) +
                    len([s for s in existing_stores if s.get("phone", "N/A") != "N/A"])
                ),
                "file": output_file,
            })

            jobs[job_id]["status"] = "complete"

        except Exception as e:
            progress_queue.put({"type": "error", "message": str(e)})
            jobs[job_id]["status"] = "error"

    thread = threading.Thread(target=run_scraper, daemon=True)
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/api/progress/<job_id>")
def stream_progress(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404

    def generate():
        q = jobs[job_id]["queue"]
        while True:
            try:
                msg = q.get(timeout=60)
                yield f"data: {json.dumps(msg)}\n\n"
                if msg.get("type") in ("complete", "error"):
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"

    return Response(generate(), mimetype="text/event-stream")


@app.route("/api/download/<filename>")
def download_file(filename):
    filepath = os.path.join(STATIC_DIR, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name="scraped_stores.xlsx")
    return jsonify({"error": "File not found"}), 404


@app.route("/api/cleanup", methods=["POST"])
def cleanup_files():
    """Delete all generated files in static/ and uploads/ directories."""
    deleted = {"static": 0, "uploads": 0, "files": []}

    for dirname, key in [(STATIC_DIR, "static"), (UPLOAD_DIR, "uploads")]:
        if os.path.exists(dirname):
            for fname in os.listdir(dirname):
                fpath = os.path.join(dirname, fname)
                if os.path.isfile(fpath):
                    try:
                        os.remove(fpath)
                        deleted[key] += 1
                        deleted["files"].append(f"{key}/{fname}")
                    except Exception:
                        pass

    # Clear in-memory upload data
    uploaded_data.clear()

    total = deleted["static"] + deleted["uploads"]
    return jsonify({
        "ok": True,
        "total_deleted": total,
        "static_deleted": deleted["static"],
        "uploads_deleted": deleted["uploads"],
        "files": deleted["files"],
    })


@app.route("/api/cleanup/info", methods=["GET"])
def cleanup_info():
    """Return count and size of files that would be cleaned up."""
    total_files = 0
    total_size = 0

    for dirname in [STATIC_DIR, UPLOAD_DIR]:
        if os.path.exists(dirname):
            for fname in os.listdir(dirname):
                fpath = os.path.join(dirname, fname)
                if os.path.isfile(fpath):
                    total_files += 1
                    total_size += os.path.getsize(fpath)

    return jsonify({
        "total_files": total_files,
        "total_size_bytes": total_size,
        "total_size_mb": round(total_size / (1024 * 1024), 2),
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
