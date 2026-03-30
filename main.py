import os
import urllib.request
import urllib.parse
import json
import re
import csv
import io
from fastapi import FastAPI
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

ZAICO_TOKEN = os.environ.get("ZAICO_TOKEN", "3pVDbsHCdn932LtmeVJzJ1xqreKPXGXe")
ZAICO_BASE = "https://web.zaico.co.jp/api/v1"
SALES_SHEET_ID = "1Ap81MeYTrNG78wO8Xs9Pm68msqxq5McbkvEj5fNBvaI"
SALES_SHEET_NAME = "販売実績"


def fetch_sold_inventory_ids() -> set:
    """販売実績シートから販売済みの在庫IDセットを取得する"""
    sheet_name = urllib.parse.quote(SALES_SHEET_NAME)
    url = f"https://docs.google.com/spreadsheets/d/{SALES_SHEET_ID}/export?format=csv&sheet={sheet_name}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        content = r.read().decode("utf-8")
    reader = csv.DictReader(io.StringIO(content))
    sold_ids = set()
    for row in reader:
        inv_id = row.get("在庫ID", "").strip()
        if inv_id:
            sold_ids.add(int(inv_id))
    return sold_ids


def parse_product_code(title: str) -> dict:
    """YS001-399-0001 → {model: YS001, material_code: 3, color_code: 99, serial: 0001}"""
    m = re.match(r"^([A-Z]+\d+)-(\d)(\d{2})(?:-(\d+))?$", title)
    if m:
        return {
            "model": m.group(1),
            "material_code": m.group(2),
            "color_code": m.group(3),
            "serial": m.group(4),
        }
    return {"model": title, "material_code": None, "color_code": None, "serial": None}


@app.post("/api/analyze")
def analyze():
    # 1. 全inventoryを取得
    print("inventories取得中...")
    inventories_raw = []
    page = 1
    while True:
        url = f"{ZAICO_BASE}/inventories?per_page=1000&page={page}"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {ZAICO_TOKEN}"})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.loads(r.read())
        if not data:
            break
        inventories_raw.extend(data)
        if len(data) < 1000:
            break
        page += 1

    inv_map = {item["id"]: item for item in inventories_raw}
    print(f"inventory: {len(inv_map)}件")

    # 2. 販売済みIDを取得
    print("販売実績シート取得中...")
    sold_ids = fetch_sold_inventory_ids()
    print(f"販売済み: {len(sold_ids)}件")

    # 3. 全packing_slipsを取得
    print("packing_slips取得中...")
    slips = []
    page = 1
    while True:
        url = f"{ZAICO_BASE}/packing_slips?per_page=100&page={page}"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {ZAICO_TOKEN}"})
        with urllib.request.urlopen(req, timeout=30) as r:
            data = json.loads(r.read())
        if not data:
            break
        slips.extend(data)
        if len(data) < 100:
            break
        page += 1
    print(f"packing_slips: {len(slips)}件")

    # 4. 各inventory_idについて「最新の出庫先」を記録
    latest_slip = {}
    for slip in slips:
        customer = slip["customer_name"] or "(不明)"
        date = slip["delivery_date"] or ""
        num = slip["num"] or ""
        for item in slip.get("deliveries", []):
            inv_id = item["inventory_id"]
            if inv_id not in latest_slip or date > latest_slip[inv_id]["date"]:
                latest_slip[inv_id] = {
                    "customer": customer,
                    "date": date,
                    "slip_num": num,
                    "slip_id": slip["id"],
                }

    # 4. quantity=0 かつ packing_slips記録あり かつ 販売済みでない → 出庫先にある商品
    at_customer = []
    for inv_id, slip_info in latest_slip.items():
        inv = inv_map.get(inv_id)
        if inv is None:
            continue
        raw_qty = inv.get("quantity", 0)
        qty = float(raw_qty) if raw_qty is not None else 0.0
        if qty > 0:
            continue  # 返品済み
        if inv_id in sold_ids:
            continue  # 販売済み
        parsed = parse_product_code(inv["title"])
        cats = inv.get("categories", [])
        at_customer.append({
            "inventory_id": inv_id,
            "title": inv["title"],
            "model": parsed["model"],
            "material_code": parsed["material_code"],
            "color_code": parsed["color_code"],
            "serial": parsed["serial"],
            "category_type": cats[0] if cats else "",
            "customer": slip_info["customer"],
            "shipped_date": slip_info["date"],
            "slip_num": slip_info["slip_num"],
        })

    # 6. 出庫先ごとに集計
    by_customer = {}
    for item in at_customer:
        c = item["customer"]
        if c not in by_customer:
            by_customer[c] = []
        by_customer[c].append(item)

    summary = sorted(
        [{"customer": k, "count": len(v), "items": v} for k, v in by_customer.items()],
        key=lambda x: -x["count"],
    )

    return JSONResponse({
        "total": len(at_customer),
        "customer_count": len(by_customer),
        "summary": summary,
    })


@app.post("/api/stock")
def stock():
    from datetime import date, datetime

    print("inventories取得中（社内在庫）...")
    inventories_raw = []
    page = 1
    while True:
        url = f"{ZAICO_BASE}/inventories?per_page=1000&page={page}"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {ZAICO_TOKEN}"})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.loads(r.read())
        if not data:
            break
        inventories_raw.extend(data)
        if len(data) < 1000:
            break
        page += 1

    today = date.today()

    # quantity > 0 の商品を型番+素材+色でグループ化
    groups = {}
    for item in inventories_raw:
        raw_qty = item.get("quantity", 0)
        qty = float(raw_qty) if raw_qty is not None else 0.0
        if qty <= 0:
            continue
        parsed = parse_product_code(item["title"])
        key = f"{parsed['model']}-{parsed['material_code']}{parsed['color_code']}"
        cats = item.get("categories", [])
        created_str = item.get("created_at", "")
        try:
            created = datetime.fromisoformat(created_str).date()
        except Exception:
            created = today

        if key not in groups:
            groups[key] = {
                "key": key,
                "model": parsed["model"],
                "material_code": parsed["material_code"],
                "color_code": parsed["color_code"],
                "category_type": cats[0] if cats else "",
                "count": 0,
                "oldest_date": created,
                "items": [],
            }
        groups[key]["count"] += qty
        groups[key]["items"].append(item["title"])
        if created < groups[key]["oldest_date"]:
            groups[key]["oldest_date"] = created

    # 3個以上のみ・件数順ソート
    result = []
    for g in groups.values():
        if g["count"] < 3:
            continue
        delta = today - g["oldest_date"]
        months = delta.days // 30
        result.append({
            "key": g["key"],
            "model": g["model"],
            "material_code": g["material_code"],
            "color_code": g["color_code"],
            "category_type": g["category_type"],
            "count": int(g["count"]),
            "oldest_date": g["oldest_date"].isoformat(),
            "months_in_stock": months,
            "items": g["items"],
        })

    result.sort(key=lambda x: -x["count"])

    return JSONResponse({
        "total_groups": len(result),
        "total_units": sum(r["count"] for r in result),
        "items": result,
    })


@app.get("/", response_class=HTMLResponse)
def index():
    here = os.path.dirname(os.path.abspath(__file__))
    return open(os.path.join(here, "index.html")).read()
