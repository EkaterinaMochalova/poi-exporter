import os, re, time, uuid, json, random, threading, zipfile
from datetime import datetime
from typing import Dict, Any, List, Optional

import requests
from fastapi import FastAPI, BackgroundTasks
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from pydantic import BaseModel
from openpyxl import Workbook

# -------------------------
# Config
# -------------------------
# Общие запросы можно фейловерить на несколько инстансов,
# но городские запросы (boundary->area) лучше не гонять на .fr (часто 403).
OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.openstreetmap.fr/api/interpreter",
]

CITY_OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
]

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)

JOBS: Dict[str, Dict[str, Any]] = {}
LOCK = threading.Lock()

CACHE_FILE = os.path.join(DATA_DIR, "cache_index.json")
try:
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        CACHE_INDEX = json.load(f)
except Exception:
    CACHE_INDEX = {}

CACHE_TTL_SECONDS = 7 * 24 * 3600  # 7 days

# -------------------------
# Cities (500k+ from your list)
# -------------------------
# key = internal id, label = what user sees, aliases = what we try in OSM
CITIES: List[Dict[str, Any]] = [
    {"key": "moscow", "label": "Москва", "aliases": ["Москва", "Moscow"]},
    {"key": "spb", "label": "Санкт-Петербург", "aliases": ["Санкт-Петербург", "Saint Petersburg", "St Petersburg", "Sankt-Peterburg"]},
    {"key": "novosibirsk", "label": "Новосибирск", "aliases": ["Новосибирск", "Novosibirsk"]},
    {"key": "ekaterinburg", "label": "Екатеринбург", "aliases": ["Екатеринбург", "Yekaterinburg", "Ekaterinburg"]},
    {"key": "kazan", "label": "Казань", "aliases": ["Казань", "Kazan"]},
    {"key": "krasnoyarsk", "label": "Красноярск", "aliases": ["Красноярск", "Krasnoyarsk"]},
    {"key": "nizhny_novgorod", "label": "Нижний Новгород", "aliases": ["Нижний Новгород", "Nizhny Novgorod"]},
    {"key": "chelyabinsk", "label": "Челябинск", "aliases": ["Челябинск", "Chelyabinsk"]},
    {"key": "ufa", "label": "Уфа", "aliases": ["Уфа", "Ufa"]},
    {"key": "krasnodar", "label": "Краснодар", "aliases": ["Краснодар", "Krasnodar"]},
    {"key": "samara", "label": "Самара", "aliases": ["Самара", "Samara"]},
    {"key": "rostov_on_don", "label": "Ростов-на-Дону", "aliases": ["Ростов-на-Дону", "Rostov-on-Don", "Rostov on Don"]},
    {"key": "omsk", "label": "Омск", "aliases": ["Омск", "Omsk"]},
    {"key": "voronezh", "label": "Воронеж", "aliases": ["Воронеж", "Voronezh"]},
    {"key": "perm", "label": "Пермь", "aliases": ["Пермь", "Perm"]},
    {"key": "volgograd", "label": "Волгоград", "aliases": ["Волгоград", "Volgograd"]},
    {"key": "saratov", "label": "Саратов", "aliases": ["Саратов", "Saratov"]},
    {"key": "tyumen", "label": "Тюмень", "aliases": ["Тюмень", "Tyumen", "Tyumen’"]},
    {"key": "tolyatti", "label": "Тольятти", "aliases": ["Тольятти", "Tolyatti", "Togliatti"]},
    {"key": "makhachkala", "label": "Махачкала", "aliases": ["Махачкала", "Makhachkala"]},
    {"key": "barnaul", "label": "Барнаул", "aliases": ["Барнаул", "Barnaul"]},
    {"key": "izhevsk", "label": "Ижевск", "aliases": ["Ижевск", "Izhevsk"]},
    {"key": "khabarovsk", "label": "Хабаровск", "aliases": ["Хабаровск", "Khabarovsk"]},
    {"key": "ulyanovsk", "label": "Ульяновск", "aliases": ["Ульяновск", "Ulyanovsk"]},
    {"key": "irkutsk", "label": "Иркутск", "aliases": ["Иркутск", "Irkutsk"]},
    {"key": "vladivostok", "label": "Владивосток", "aliases": ["Владивосток", "Vladivostok"]},
    {"key": "yaroslavl", "label": "Ярославль", "aliases": ["Ярославль", "Yaroslavl", "Yaroslavl’"]},
    {"key": "stavropol", "label": "Ставрополь", "aliases": ["Ставрополь", "Stavropol"]},
    {"key": "sevastopol", "label": "Севастополь", "aliases": ["Севастополь", "Sevastopol"]},
    {"key": "naberezhnye_chelny", "label": "Набережные Челны", "aliases": ["Набережные Челны", "Naberezhnye Chelny"]},
    {"key": "tomsk", "label": "Томск", "aliases": ["Томск", "Tomsk"]},
    {"key": "balashikha", "label": "Балашиха", "aliases": ["Балашиха", "Balashikha"]},
    {"key": "kemerovo", "label": "Кемерово", "aliases": ["Кемерово", "Kemerovo"]},
    {"key": "orenburg", "label": "Оренбург", "aliases": ["Оренбург", "Orenburg"]},
    {"key": "novokuznetsk", "label": "Новокузнецк", "aliases": ["Новокузнецк", "Novokuznetsk"]},
    {"key": "ryazan", "label": "Рязань", "aliases": ["Рязань", "Ryazan", "Ryazan’"]},
]

CITY_BY_KEY = {c["key"]: c for c in CITIES}

# -------------------------
# Helpers
# -------------------------
def now_iso() -> str:
    return datetime.utcnow().isoformat()

def persist_cache():
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(CACHE_INDEX, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def request_signature(req: Dict[str, Any]) -> str:
    norm = {
        "q": (req.get("q") or "").strip(),
        "scope_type": req.get("scope_type"),
        "city_key": (req.get("city_key") or "").strip(),
        "limit": int(req.get("limit") or 0),
        "format": req.get("format"),
    }
    return json.dumps(norm, ensure_ascii=False, sort_keys=True)

def overpass_request(query: str, endpoints: List[str], max_tries: int = 8) -> Dict[str, Any]:
    last_err = None
    for attempt in range(1, max_tries + 1):
        endpoint = endpoints[(attempt - 1) % len(endpoints)]
        try:
            r = requests.post(endpoint, data={"data": query}, timeout=180)

            if r.status_code in (401, 403, 429) or r.status_code >= 500:
                last_err = RuntimeError(f"{r.status_code} from {endpoint}")
                time.sleep(min(120, 6 * attempt) + random.uniform(0, 2))
                continue

            r.raise_for_status()
            data = r.json()
            if "elements" not in data:
                raise RuntimeError(f"No elements in response: {list(data.keys())}")
            return data

        except Exception as e:
            last_err = e
            time.sleep(min(180, 8 * attempt) + random.uniform(0, 3))

    raise RuntimeError(f"Overpass failed after retries. Last error: {last_err}")

def escape_overpass_regex(s: str) -> str:
    return (s or "").replace('"', '\\"').strip()

def build_query_russia(q: str, limit: int) -> str:
    q = escape_overpass_regex(q)
    if not q:
        raise ValueError("Empty query")

    lim = f" {int(limit)}" if int(limit or 0) > 0 else ""
    return f"""
[out:json][timeout:180];
area(3600060189)->.a;  // Russia (relation 60189)
(
  nwr(area.a)["name"~"{q}",i];
  nwr(area.a)["brand"~"{q}",i];
  nwr(area.a)["operator"~"{q}",i];
);
out center{lim};
"""

def build_query_city(q: str, alias: str, limit: int) -> str:
    q = escape_overpass_regex(q)
    alias = (alias or "").replace('"', '\\"').strip()
    if not q:
        raise ValueError("Empty query")
    if not alias:
        raise ValueError("Empty city alias")

    lim = f" {int(limit)}" if int(limit or 0) > 0 else ""
    return f"""
[out:json][timeout:180];
(
  rel["boundary"="administrative"]["admin_level"~"8|6|4"]["name"="{alias}"];
  rel["boundary"="administrative"]["admin_level"~"8|6|4"]["name:ru"="{alias}"];
  rel["boundary"="administrative"]["admin_level"~"8|6|4"]["name:en"="{alias}"];
  way["boundary"="administrative"]["admin_level"~"8|6|4"]["name"="{alias}"];
  way["boundary"="administrative"]["admin_level"~"8|6|4"]["name:ru"="{alias}"];
  way["boundary"="administrative"]["admin_level"~"8|6|4"]["name:en"="{alias}"];

  nwr["place"="city"]["name"="{alias}"];
  nwr["place"="city"]["name:ru"="{alias}"];
  nwr["place"="city"]["name:en"="{alias}"];
)->.r;

.r map_to_area->.a;

(
  nwr(area.a)["name"~"{q}",i];
  nwr(area.a)["brand"~"{q}",i];
  nwr(area.a)["operator"~"{q}",i];
);
out center{lim};
"""

def extract_pois(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for el in data.get("elements", []):
        tags = el.get("tags", {}) or {}
        lat = el.get("lat")
        lon = el.get("lon")
        center = el.get("center") or {}
        if lat is None or lon is None:
            lat = center.get("lat")
            lon = center.get("lon")
        if lat is None or lon is None:
            continue

        name = tags.get("name") or tags.get("brand") or tags.get("operator") or "POI"
        out.append({
            "Название": str(name),
            "Широта": float(lat),
            "Долгота": float(lon),
            "osm_type": el.get("type", ""),
            "osm_id": str(el.get("id", "")),
        })

    seen = set()
    ded = []
    for r in out:
        k = (r["osm_type"], r["osm_id"])
        if k in seen:
            continue
        seen.add(k)
        ded.append(r)
    return ded

def write_xlsx(path: str, rows: List[Dict[str, Any]]):
    wb = Workbook()
    ws = wb.active
    ws.title = "POI"
    ws.append(["Название", "Широта", "Долгота"])
    for r in rows:
        ws.append([r["Название"], r["Широта"], r["Долгота"]])
    wb.save(path)

def write_geojson(path: str, rows: List[Dict[str, Any]]):
    features = []
    for r in rows:
        features.append({
            "type": "Feature",
            "properties": {"name": r["Название"]},
            "geometry": {"type": "Point", "coordinates": [r["Долгота"], r["Широта"]]},
        })
    fc = {"type": "FeatureCollection", "features": features}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(fc, f, ensure_ascii=False)

def list_history():
    items = []
    for job_id in os.listdir(DATA_DIR):
        d = os.path.join(DATA_DIR, job_id)
        z = os.path.join(d, "result.zip")
        meta = os.path.join(d, "meta.json")
        if os.path.isfile(z) and os.path.isfile(meta):
            try:
                with open(meta, "r", encoding="utf-8") as f:
                    m = json.load(f)
                st = os.stat(z)
                items.append({
                    "job_id": job_id,
                    "created_at": m.get("created_at"),
                    "q": m.get("q"),
                    "scope_type": m.get("scope_type"),
                    "city_key": m.get("city_key"),
                    "city_label": m.get("city_label"),
                    "count": m.get("count"),
                    "size_bytes": st.st_size,
                    "mtime": int(st.st_mtime),
                })
            except Exception:
                pass
    items.sort(key=lambda x: x["mtime"], reverse=True)
    return items

# -------------------------
# API models
# -------------------------
class ExportRequest(BaseModel):
    q: str
    scope_type: str = "ru"     # "ru" | "city"
    city_key: Optional[str] = None
    limit: int = 50           # default 50
    format: str = "xlsx"      # "xlsx" | "geojson" | "both"

# -------------------------
# Worker
# -------------------------
def run_job(job_id: str, req: Dict[str, Any]):
    job_dir = os.path.join(DATA_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    def set_job(**kw):
        with LOCK:
            JOBS[job_id].update(kw)

    set_job(status="running", message="Preparing query…", started_at=now_iso(), progress=5)

    try:
        q = req["q"]
        scope_type = req.get("scope_type") or "ru"
        limit = int(req.get("limit") or 0)

        query_tried = None

        if scope_type == "ru":
            set_job(message="Querying Russia…", progress=10)
            query_tried = build_query_russia(q, limit)
            data = overpass_request(query_tried, OVERPASS_ENDPOINTS)

        elif scope_type == "city":
            city_key = (req.get("city_key") or "").strip()
            if not city_key or city_key not in CITY_BY_KEY:
                raise ValueError("City is required (choose from list).")

            city = CITY_BY_KEY[city_key]
            aliases = city.get("aliases") or [city["label"]]

            set_job(message=f"Geocoding city: {city['label']}…", progress=10)
            last_err = None
            data = None

            for idx, alias in enumerate(aliases, 1):
                set_job(
                    message=f"City match {idx}/{len(aliases)}: {alias}",
                    progress=15 + int(30 * idx / max(1, len(aliases)))
                )
                query_try = build_query_city(q, alias, limit)
                query_tried = query_try
                try:
                    data = overpass_request(query_try, CITY_OVERPASS_ENDPOINTS)
                    break
                except Exception as e:
                    last_err = e
                    time.sleep(1.2 + random.uniform(0, 0.7))

            if data is None:
                raise RuntimeError(f"City geocode failed: {last_err}")

        else:
            raise ValueError("Unknown scope_type")

        # Save query for debugging
        with open(os.path.join(job_dir, "query.overpassql"), "w", encoding="utf-8") as f:
            f.write(query_tried or "")

        set_job(message="Parsing response…", progress=60)
        pois = extract_pois(data)

        set_job(message=f"Found {len(pois)} POI. Writing files…", progress=75)

        out_files = []
        fmt = req.get("format") or "xlsx"
        if fmt in ("xlsx", "both"):
            xlsx_path = os.path.join(job_dir, "poi.xlsx")
            write_xlsx(xlsx_path, pois)
            out_files.append(("poi.xlsx", xlsx_path))

        if fmt in ("geojson", "both"):
            gj_path = os.path.join(job_dir, "poi.geojson")
            write_geojson(gj_path, pois)
            out_files.append(("poi.geojson", gj_path))

        set_job(message="Packing ZIP…", progress=90)

        zip_path = os.path.join(job_dir, "result.zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for arcname, p in out_files:
                z.write(p, arcname=arcname)
            z.writestr(
                "README.txt",
                "Data source: OpenStreetMap contributors (ODbL). Generated via Overpass API.\n"
                "Automated extraction; may be incomplete.\n"
            )

        meta = {
            "created_at": now_iso(),
            "q": req.get("q"),
            "scope_type": req.get("scope_type"),
            "city_key": req.get("city_key"),
            "city_label": CITY_BY_KEY.get(req.get("city_key"), {}).get("label") if req.get("city_key") else None,
            "limit": req.get("limit"),
            "format": fmt,
            "count": len(pois),
        }
        with open(os.path.join(job_dir, "meta.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        set_job(status="done", message="Ready", finished_at=now_iso(), zip_path=zip_path, count=len(pois), progress=100)

    except Exception as e:
        set_job(status="error", message=str(e), finished_at=now_iso(), progress=100)

# -------------------------
# FastAPI
# -------------------------
app = FastAPI(title="POI Exporter")

@app.get("/", response_class=HTMLResponse)
def index():
    city_options = "\n".join([f'<option value="{c["key"]}">{c["label"]}</option>' for c in CITIES])

    html = r"""
<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>POI Exporter</title>
<style>
  :root{
    --bg1:#f5f6f8;
    --bg2:#eef1f6;
    --card: rgba(255,255,255,.62);
    --stroke: rgba(255,255,255,.55);
    --text:#0b0b0f;
    --muted: rgba(20,20,30,.62);
    --shadow: 0 18px 45px rgba(0,0,0,.08);
    --shadow2: 0 2px 10px rgba(0,0,0,.06);
    --r: 18px;
  }
  *{box-sizing:border-box}
  body{
    font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial;
    margin:0;
    color:var(--text);
    background:
      radial-gradient(1200px 800px at 20% 10%, #ffffff 0%, transparent 55%),
      radial-gradient(900px 700px at 85% 25%, #dbe6ff 0%, transparent 55%),
      radial-gradient(1000px 700px at 65% 85%, #ffe6f2 0%, transparent 55%),
      linear-gradient(180deg, var(--bg1), var(--bg2));
    min-height:100vh;
  }
  .wrap{max-width:1040px;margin:34px auto;padding:0 18px 40px;}
  h1{margin:0 0 10px 0;letter-spacing:-0.02em;font-size:34px;line-height:1.1;}
  .muted{color:var(--muted);font-size:14px;line-height:1.35;}
  .glass{
    background:var(--card);
    border:1px solid var(--stroke);
    box-shadow:var(--shadow);
    border-radius:calc(var(--r) + 6px);
    backdrop-filter:blur(18px);
    -webkit-backdrop-filter:blur(18px);
    overflow:hidden;
  }
  .card{padding:16px;}
  .card + .card{margin-top:14px;}
  label{display:block;font-size:13px;font-weight:620;margin:0 0 8px 0;}
  input, select{
    width:100%;
    padding:11px 12px;
    border-radius:14px;
    border:1px solid rgba(0,0,0,.10);
    background:rgba(255,255,255,.72);
    box-shadow:var(--shadow2);
    outline:none;
    transition:border-color .15s ease, background .15s ease, transform .08s ease;
  }
  input:focus, select:focus{border-color:rgba(0,0,0,.22);background:rgba(255,255,255,.88);}
  input::placeholder{color:rgba(0,0,0,.35);}
  .grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;align-items:start;}
  .grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;align-items:end;}
  .field{min-width:0;}
  .actions{display:flex;gap:12px;align-items:center;justify-content:flex-start;flex-wrap:wrap;}
  button{
    padding:11px 14px;
    border-radius:14px;
    border:1px solid rgba(0,0,0,.16);
    background:rgba(20,20,24,.92);
    color:#fff;
    font-weight:650;
    cursor:pointer;
    box-shadow:0 14px 28px rgba(0,0,0,.14);
    transition:transform .08s ease, opacity .2s ease;
  }
  button:hover{transform:translateY(-1px);}
  button:disabled{opacity:.55;cursor:not-allowed;transform:none;}
  .status{
    font-family: ui-monospace, Menlo, Monaco, Consolas, monospace;
    color: rgba(20,20,30,.55);
    font-size: 13px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 100%;
  }
  .hint{margin-top:6px;color:var(--muted);font-size:13px;}
  .links{margin-top:12px;font-family:ui-monospace,Menlo,Monaco,Consolas,monospace;font-size:14px;}
  a{color:#0b57d0;text-decoration:none;}
  a:hover{text-decoration:underline;}
  table{width:100%;border-collapse:collapse;margin-top:10px;overflow:hidden;border-radius:14px;}
  th,td{border-bottom:1px solid rgba(0,0,0,.06);padding:10px;text-align:left;font-size:13px;}
  th{background: rgba(255,255,255,.55);}

  .progressWrap{
    width:100%;
    height:10px;
    border-radius:999px;
    background: rgba(0,0,0,.06);
    overflow:hidden;
    border:1px solid rgba(255,255,255,.55);
  }
  .progressBar{
    height:100%;
    width:0%;
    background: rgba(20,20,24,.88);
    border-radius:999px;
    transition: width .25s ease;
  }

  .presets{
    display:flex;
    flex-wrap:wrap;
    gap:8px;
    margin-top:10px;
  }

  .preset{
    padding:6px 12px;
    border-radius:999px;
    font-size:13px;
    background: rgba(255,255,255,.6);
    border:1px solid rgba(255,255,255,.7);
    box-shadow: 0 4px 10px rgba(0,0,0,.06);
    cursor:pointer;
    transition: all .15s ease;
    backdrop-filter: blur(10px);
    user-select:none;
  }

  .preset:hover{
    background: rgba(255,255,255,.9);
    transform: translateY(-1px);
  }

  .preset.active{
    background: rgba(20,20,24,.9);
    color:white;
  }

  .hidden{display:none;}

  @media (max-width: 860px){
    .grid2{grid-template-columns:1fr;}
    .grid3{grid-template-columns:1fr;}
    .status{white-space:normal;}
  }
</style>
</head>
<body>
<div class="wrap">
  <h1>POI Exporter 🧭</h1>
  <div class="muted">Поиск POI по OpenStreetMap (Overpass). Экспорт: <b>Название / Широта / Долгота</b>. © OpenStreetMap contributors.</div>

  <div class="glass card">
    <div class="grid2">
      <div class="field">
        <label>Что ищем (бренд/название)</label>
        <input id="q" placeholder="Магнит Косметик | Starbucks | Лента" />
        <div class="hint">Ищет в тегах name/brand/operator. Можно regex.</div>

        <div class="presets">
          <div class="preset" data-q="Магнит">Магнит</div>
          <div class="preset" data-q="Пятёрочка">Пятёрочка</div>
          <div class="preset" data-q="Starbucks">Starbucks</div>
          <div class="preset" data-q="Магнит Косметик">Магнит Косметик</div>
          <div class="preset" data-q="АЗС">АЗС</div>
        </div>
      </div>

      <div class="field">
        <label>Формат</label>
        <select id="format">
          <option value="xlsx">Excel (.xlsx)</option>
          <option value="geojson">GeoJSON</option>
          <option value="both">ZIP: XLSX + GeoJSON</option>
        </select>
        <div class="hint">Скачивание одним ZIP.</div>
      </div>
    </div>

    <div class="grid2" style="margin-top:14px;">
      <div class="field">
        <label>Регион</label>
        <select id="scope_type">
          <option value="ru">Вся Россия</option>
          <option value="city">Город (500k+)</option>
        </select>
      </div>

      <div class="field" id="city_field">
        <label>Город</label>
        <select id="city_key">
          <option value="">Выберите город…</option>
          __CITY_OPTIONS__
        </select>
        <div class="hint">Выбери город, в котором нужны точки.</div>
      </div>
    </div>

    <div class="grid3" style="margin-top:14px;">
      <div class="field">
        <label>Limit</label>
        <input id="limit" type="number" min="1" value="50" />
        <div class="hint">Сколько точек максимум выгружать.</div>
      </div>

      <div class="field actions">
        <button id="go">Сгенерировать</button>
        <div id="status" class="status"></div>
      </div>

      <div class="field">
        <label>Прогресс</label>
        <div class="progressWrap"><div id="pbar" class="progressBar"></div></div>
      </div>
    </div>

    <div id="links" class="links"></div>
  </div>

  <div class="glass card">
    <div style="font-weight:650;font-size:14px;margin:0 0 10px 0;">История</div>
    <div class="muted">Старые выгрузки доступны тут же.</div>
    <div id="history"></div>
  </div>
</div>

<script>
const el = (id)=>document.getElementById(id);
const statusEl = el("status");
const linksEl = el("links");
const historyEl = el("history");
const pbar = el("pbar");
let timer = null;

function setStatus(t){ statusEl.textContent = t || ""; }
function setProgress(p){
  const v = Math.max(0, Math.min(100, parseInt(p || 0, 10)));
  pbar.style.width = v + "%";
}

function updateCityVisibility(){
  const st = el("scope_type").value;
  el("city_field").classList.toggle("hidden", st !== "city");
}
el("scope_type").addEventListener("change", updateCityVisibility);
updateCityVisibility();

document.querySelectorAll(".preset").forEach(p => {
  p.addEventListener("click", () => {
    document.querySelectorAll(".preset").forEach(x => x.classList.remove("active"));
    p.classList.add("active");
    el("q").value = p.dataset.q;
    el("q").focus();
  });
});

async function refreshHistory(){
  const r = await fetch("/api/history");
  const data = await r.json();
  const items = data.items || [];
  if (!items.length){
    historyEl.innerHTML = "<div class='muted'>Пока нет выгрузок.</div>";
    return;
  }
  let html = "<table><thead><tr><th>Когда</th><th>Запрос</th><th>Скоуп</th><th>POI</th><th></th></tr></thead><tbody>";
  for (const it of items.slice(0, 15)){
    const when = (it.created_at || "").replace("T"," ").slice(0,19);
    let scope = "Россия";
    if (it.scope_type === "city") scope = "Город: " + (it.city_label || it.city_key || "");
    html += `<tr>
      <td>${when}</td>
      <td>${it.q || ""}</td>
      <td>${scope}</td>
      <td>${it.count ?? ""}</td>
      <td><a href="/api/download/${it.job_id}">download</a></td>
    </tr>`;
  }
  html += "</tbody></table>";
  historyEl.innerHTML = html;
}

async function start(){
  linksEl.innerHTML = "";
  setProgress(0);
  setStatus("creating job…");
  el("go").disabled = true;

  const q = el("q").value.trim();
  if (!q){ setStatus("введи запрос"); el("go").disabled=false; return; }

  const scope_type = el("scope_type").value;
  const city_key = el("city_key").value;
  const format = el("format").value;
  const limit = Math.max(1, parseInt(el("limit").value || "50", 10) || 50);

  if (scope_type === "city" && !city_key){
    setStatus("выбери город");
    el("go").disabled=false;
    return;
  }

  const payload = { q, scope_type, city_key, limit, format };

  const res = await fetch("/api/export", {
    method:"POST",
    headers:{ "Content-Type":"application/json" },
    body: JSON.stringify(payload)
  });

  const data = await res.json();
  const job_id = data.job_id;

  setStatus("job: " + job_id);

  timer = setInterval(async ()=>{
    const sr = await fetch("/api/status/" + job_id);
    const s = await sr.json();
    setStatus(`${s.status} | ${s.message || ""}`);
    setProgress(s.progress || 0);

    if (s.status === "done"){
      clearInterval(timer);
      el("go").disabled = false;
      linksEl.innerHTML = `✅ <a href="/api/download/${job_id}">Скачать результат (ZIP)</a>`;
      refreshHistory();
    }
    if (s.status === "error"){
      clearInterval(timer);
      el("go").disabled = false;
      linksEl.innerHTML = `❌ ${s.message || "error"}`;
      refreshHistory();
    }
  }, 1100);
}

el("go").addEventListener("click", start);
refreshHistory();
</script>
</body>
</html>
"""
    html = html.replace("__CITY_OPTIONS__", city_options)
    return HTMLResponse(html)

@app.get("/api/cities")
def api_cities():
    return JSONResponse({"cities": [{"key": c["key"], "label": c["label"]} for c in CITIES]})

@app.post("/api/export")
def api_export(req: ExportRequest, background: BackgroundTasks):
    payload = req.model_dump()
    sig = request_signature(payload)
    now = time.time()

    cached_job = CACHE_INDEX.get(sig)
    if cached_job:
        meta_path = os.path.join(DATA_DIR, cached_job, "meta.json")
        zip_path = os.path.join(DATA_DIR, cached_job, "result.zip")
        if os.path.exists(meta_path) and os.path.exists(zip_path):
            try:
                if now - os.stat(zip_path).st_mtime < CACHE_TTL_SECONDS:
                    return JSONResponse({"job_id": cached_job, "cached": True})
            except Exception:
                pass

    job_id = uuid.uuid4().hex[:12]
    job_dir = os.path.join(DATA_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    with LOCK:
        JOBS[job_id] = {
            "job_id": job_id,
            "status": "queued",
            "message": "Queued",
            "created_at": now_iso(),
            "zip_path": None,
            "count": None,
            "progress": 0,
        }

    CACHE_INDEX[sig] = job_id
    persist_cache()

    background.add_task(run_job, job_id, payload)
    return JSONResponse({"job_id": job_id, "cached": False})

@app.get("/api/status/{job_id}")
def api_status(job_id: str):
    with LOCK:
        job = JOBS.get(job_id)

    if not job:
        meta_path = os.path.join(DATA_DIR, job_id, "meta.json")
        if os.path.exists(meta_path):
            with open(meta_path, "r", encoding="utf-8") as f:
                m = json.load(f)
            z = os.path.join(DATA_DIR, job_id, "result.zip")
            return JSONResponse({
                "job_id": job_id,
                "status": "done" if os.path.exists(z) else "unknown",
                "message": "Stored job",
                "count": m.get("count"),
                "progress": 100 if os.path.exists(z) else 0,
            })
        return JSONResponse({"status": "error", "message": "unknown job_id", "progress": 100}, status_code=404)

    return JSONResponse({
        "job_id": job_id,
        "status": job.get("status"),
        "message": job.get("message"),
        "count": job.get("count"),
        "progress": job.get("progress", 0),
    })

@app.get("/api/download/{job_id}")
def api_download(job_id: str):
    zip_path = os.path.join(DATA_DIR, job_id, "result.zip")
    if not os.path.exists(zip_path):
        return JSONResponse({"status": "error", "message": "not ready"}, status_code=400)
    return FileResponse(zip_path, media_type="application/zip", filename=f"poi_{job_id}.zip")

@app.get("/api/history")
def api_history():
    return JSONResponse({"items": list_history()})