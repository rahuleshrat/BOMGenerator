import os, uuid, ezdxf, math, json, pandas as pd
from shapely.geometry import LineString
from collections import defaultdict
from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
import matplotlib.pyplot as plt
from fastapi.staticfiles import StaticFiles

def render_preview(dxf_path, out_file):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    fig, ax = plt.subplots()
    for e in msp:
        if e.dxftype() == "LINE":
            s, e_ = e.dxf.start, e.dxf.end
            ax.plot([s[0], e_[0]], [s[1], e_[1]], "k-")
        elif e.dxftype() == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, "b-")
        elif e.dxftype() == "POLYLINE":
            pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices()]
            if len(pts) > 1:
                xs, ys = zip(*pts)
                ax.plot(xs, ys, "g-")

    ax.set_aspect("equal")
    ax.axis("off")
    fig.savefig(out_file, bbox_inches="tight", dpi=200)
    plt.close(fig)
    return out_file



# ---------- DXF Parsing ----------
def entity_length(e):
    if e.dxftype() == "LINE":
        s, e_ = e.dxf.start, e.dxf.end
        return math.dist(s, e_)
    elif e.dxftype() in ("LWPOLYLINE", "POLYLINE"):
        pts = []
        if e.dxftype() == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in e.get_points()]
        else:
            pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices()]
        return LineString(pts).length if len(pts) > 1 else 0.0
    return 0.0

def explore_dxf(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    entity_counts = defaultdict(int)
    layer_counts = defaultdict(int)
    block_counts = defaultdict(int)

    for e in msp:
        t = e.dxftype()
        entity_counts[t] += 1
        if hasattr(e.dxf, "layer"):
            layer_counts[e.dxf.layer.upper()] += 1
        if t == "INSERT":
            block_counts[e.dxf.name.upper()] += 1

    return entity_counts, layer_counts, block_counts

def ensure_mapping(mapping_path, layers, blocks):
    """Generate mapping.json if missing"""
    if not os.path.exists(mapping_path):
        mapping = {
            "layers": {l: {"item": l, "unit": "m"} for l in layers.keys()},
            "blocks": {b: {"item": b, "unit": "pcs"} for b in blocks.keys()},
            "defaults": {"unit_length": "m", "length_precision": 2}
        }
        with open(mapping_path, "w") as f:
            json.dump(mapping, f, indent=2)
        print(f"✅ mapping.json generated with {len(layers)} layers and {len(blocks)} blocks.")

def parse_dxf(dxf_path, mapping_path="mapping.json"):
    with open(mapping_path, "r") as f:
        mapping = json.load(f)

    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    bom = {}

    for e in msp:
        if e.dxftype() in ("LINE", "LWPOLYLINE", "POLYLINE"):
            layer = (e.dxf.layer or "").upper()
            if layer in mapping["layers"]:
                length = entity_length(e) / 1000.0  # assume mm→m
                item = mapping["layers"][layer]["item"]
                unit = mapping["layers"][layer]["unit"]
                bom.setdefault(item, {"quantity": 0, "unit": unit, "source": "Layer"})
                bom[item]["quantity"] += length
        elif e.dxftype() == "INSERT":
            blk = (e.dxf.name or "").upper()
            if blk in mapping["blocks"]:
                item = mapping["blocks"][blk]["item"]
                unit = mapping["blocks"][blk]["unit"]
                bom.setdefault(item, {"quantity": 0, "unit": unit, "source": "Block"})
                bom[item]["quantity"] += 1

    for k, v in bom.items():
        if v["unit"] == "m":
            v["quantity"] = round(v["quantity"], mapping["defaults"].get("length_precision", 2))
        else:
            v["quantity"] = int(v["quantity"])
    return bom

def export_excel(bom, out_file="BoM.xlsx"):
    rows = [{"Item": k, "Quantity": v["quantity"], "Unit": v["unit"], "Source": v["source"]}
            for k, v in bom.items()]
    df = pd.DataFrame(rows)
    df.to_excel(out_file, index=False)
    return out_file, rows

# ---------- FastAPI ----------
app = FastAPI()
templates = Jinja2Templates(directory="templates")
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)
app.mount("/uploads", StaticFiles(directory="uploads"), name="uploads")

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})

@app.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".dxf"):
        return templates.TemplateResponse("upload.html", {"request": request, "error": "Upload DXF file"})

    uid = str(uuid.uuid4())[:8]
    dxf_path = os.path.join(UPLOAD_DIR, f"{uid}_{file.filename}")
    with open(dxf_path, "wb") as f:
        f.write(await file.read())

    # --- Explore DXF ---
    entity_summary, layers, blocks = explore_dxf(dxf_path)

    # --- Auto-generate mapping.json if missing ---
    ensure_mapping("mapping.json", layers, blocks)

    # --- Parse to BoM ---
    bom = parse_dxf(dxf_path, "mapping.json")
    out_excel = os.path.join(UPLOAD_DIR, f"{uid}_BoM.xlsx")
    excel_path, rows = export_excel(bom, out_excel)
    preview_file = os.path.join(UPLOAD_DIR, f"{uid}_preview.png")
    render_preview(dxf_path, preview_file)
    return templates.TemplateResponse("upload.html", {
        "request": request,
        "entity_summary": dict(entity_summary),
        "layers": sorted(layers.items(), key=lambda x: -x[1]),
        "blocks": sorted(blocks.items(), key=lambda x: -x[1]),
        "bom": rows,
        "preview": f"/uploads/{os.path.basename(preview_file)}",
        "download_link": f"/download/{os.path.basename(excel_path)}"
    })

@app.get("/download/{filename}")
def download_file(filename: str):
    file_path = os.path.join(UPLOAD_DIR, filename)
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)
