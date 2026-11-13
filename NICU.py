"""
NICU 營養計算器 (Streamlit)
複製自 NICU 計算.xlsx 每床頁 (1–27、預備床) 的計算邏輯:
  - 餵食熱量 (kcal/mL): PF = 0.8 x 濃度/16、Similac PF = kcal/30mL、
    MF = 2/3 x 濃度/14、HM = 0.67 + fortifier 包數 x 0.13
  - 餵食 SIR = 熱量 x 1000 / BW(kg) / 2 / 4 / 1440 (假設一半熱量來自醣類)
  - IV / PN SIR (mg/kg/min) = 速率 (mL/hr) x Glucose% / BW(kg) / 6
  - PN 熱量 = mL/day x (Glu% x 3.4 + AA% x 4) / 100;IF 20% = 2 kcal/mL
  - PN 滲透壓 = Glu% x 50 + AA% x 100 + Na x 2 + K x 2
  - 電解質、Ca/P、Fe、Zn 與奶水成分 (每公升) / 母乳添加劑 (每 100 mL) 小計
另含「RI 計算」與參考表 (IVF、本院 PN 套餐、奶水成分)。
登入 (選用): 帳號 PICU (床位 PI-01~PI-25) / NBC (床位 NBC-01~NBC-27),
  登入後可點選床位、自動讀取該床最近一次紀錄、儲存與查看歷次紀錄。
  儲存位置: secrets 有 [gcp_service_account] 與 [gsheet] 時存到 Google Sheets
  (需 pip install gspread),否則存本機 SQLite (nicu_records.db,環境變數 NICU_DB 可改路徑)。
  密碼可在 .streamlit/secrets.toml 以 [passwords] PICU = "..." 覆寫。
執行: streamlit run streamlit_app.py
"""
import base64
import hmac
import json
import os
import sqlite3
from contextlib import closing
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# 常數與參考表 (取自 Excel)
# ----------------------------------------------------------------------
GLU_KCAL_PER_G = 3.4
AA_KCAL_PER_G = 4.0
FAT20_G_PER_ML = 0.2
FAT20_KCAL_PER_ML = 2.0      # 0.2 g/mL x 10 kcal/g
FE_MG_PER_GTT = 2.5          # 鐵劑 1 gtt = 2.5 mg

# 本院每包 PN 每天固定添加
PN_FIXED_ADD = {"Zn (mg)": 1.1, "Vitamin A (IU)": 200, "Vitamin D2 (IU)": 200,
                "Vitamin E (mg)": 1, "Vitamin C (mg)": 100}

# 參考範圍 (Excel 標示)
REF = {
    "na": (2, 4), "k": (2, 3), "ca_total": 200, "p_total": 100,
    "ca_pn": 60, "p_pn": (45, 60), "cap_ratio": (1.3, 1.7),
    "fe": 2, "zn": 400, "aa_tc": 0.15, "fat_tc": 0.30,
}
PERIPHERAL_OSM = 900  # 一般建議周邊靜脈 PN 滲透壓上限 (非 Excel 原有)

# IVF 種類: Na, K (mEq/L), Glucose (%)
IVF_TABLE = {
    "自訂": None,
    "D5W": (0, 0, 5),
    "D10W": (0, 0, 10),
    "3% NaCl": (513, 0, 0),
    "N/S (0.9% NaCl)": (154, 0, 0),
    "Half Saline (0.45% NaCl)": (77, 0, 0),
    "0.45 G/S": (77, 0, 2.5),
    "0.33 G/S": (56, 0, 5),
    "Taita No.1": (25, 18, 3.8),
    "Taita No.5": (36, 18, 10),
    "Lactated Ringer's": (130, 4, 0),
}

# 本院現成 PN 套餐: Glu%, AA%, Na, K, Cl, Mg, Ca, Acetate, P
PN_SETS = pd.DataFrame(
    [
        [5, 1.5, 19.9, 18.5, 19.9, 3.1, 9.7, 8.4, 6.9],
        [7.5, 1.5, 19.9, 18.5, 19.9, 3.1, 9.7, 8.4, 6.9],
        [10, 2, 20.7, 17.8, 20.7, 3.3, 7.3, 7.25, 7.2],
        [12.5, 2, 20.7, 17.8, 20.7, 3.3, 7.3, 7.25, 7.2],
        [20, 2.5, 21.5, 19, 22, 3.3, 7.3, 7.5, 7.5],
    ],
    columns=["Glucose (%)", "AA (%)", "Na", "K", "Cl", "Mg", "Ca", "Acetate", "P"],
)

# 奶水成分 (每公升; 添加劑為每 100 mL)
# kcal, Protein g, Ca mg, P mg, Na mg, K mg, Fe mg, Zn mg
NUTRIENTS = ["Protein (g)", "Calcium (mg)", "Phosphorus (mg)", "Sodium (mg)",
             "Potassium (mg)", "Iron (mg)", "Zinc (mg)"]
MILK_TABLE = {
    "母乳 (1 wk)": [670, 24, 250, 140, 506, 702, 0, 5],
    "成熟母乳 (1 mo)": [670, 14, 250, 130, 253, 585, 1.2, 3.4],
    "Alfare 雀巢好敏瑞": [700, 21, 540, 360, 341, 880, 7, 7],
    "S-26 SMA LBW 惠氏出生低體重": [820, 22, 1010, 610, 420, 850, 8, 8],
    "Enfalac LBW 優生早產兒": [810, 24, 970, 530, 460, 810, 14, 8.1],
    "Similac Special Care 20": [680, 20, 1220, 680, 290, 870, 12, 10],
    "Similac Special Care 24": [810, 24, 1460, 810, 350, 1050, 15, 12],
    "SSC 26": [876.7, 26, 1583.3, 876.7, 380, 1136.7, 16, 13],
    "SSC 28": [943.3, 28, 1706.7, 943.3, 410, 1223.3, 17, 14],
    "Similac Special Care 30": [1010, 30, 1830, 1010, 440, 1310, 18, 15],
    "Enfalac A+ 優生A+": [670, 14, 530, 300, 173, 710, 8.2, 5.3],
    "S-26 SMA Gold 惠氏": [670, 14, 490, 280, 160, 650, 8, 6],
    "Similac High Q plus": [680, 14, 530, 280, 180, 830, 12, 5],
    "Similac Advance HH 水解蛋白": [710, 18.7, 670, 350, 290, 710, 8.5, 4.7],
    "Enfalac A+ H.P. 水解蛋白": [660, 16.5, 440, 300, 260, 770, 7.9, 6.6],
    "NAN H.A. 能恩水解蛋白": [670, 15.1, 440, 240, 180, 670, 6.6, 6.2],
    "Similac NeoSure 出院後": [740, 19, 780, 460, 250, 1060, 13, 9],
    "S-26 SMA PDF Gold 出院後": [730, 19, 730, 420, 270, 710, 12, 7.3],
    "SSC 24 with HP": [810, 27, 1460, 810, 350, 1050, 15, 12],
    "Similac Special Care 24 (120 mL, 添加鐵質)": [830, 22, 1460, 730, 370, 1250, 15, 12],
}
FORTIFIER_TABLE = {
    "無": [0] * 8,
    "Similac HMF 心美力母乳營養補充品": [14, 1, 117, 67, 15, 63, 0.35, 1],
    "Enfamil HMF 優生母乳強化營養品": [14, 1.1, 90, 50, 16, 29, 1.44, 0.72],
}


# ----------------------------------------------------------------------
# 登入與床位紀錄
# ----------------------------------------------------------------------
TW_TZ = timezone(timedelta(hours=8))
UNITS = {
    "PICU": [f"PI-{i:02d}" for i in range(1, 26)],
    "NBC": [f"NBC-{i:02d}" for i in range(1, 28)],
}
DEFAULT_PASSWORD = "hmmch"
DB_PATH = Path(os.environ.get("NICU_DB", Path(__file__).parent / "nicu_records.db"))


def get_password(unit: str) -> str:
    """優先使用 st.secrets["passwords"][unit],否則用預設密碼。"""
    try:
        return str(st.secrets["passwords"][unit])
    except Exception:
        return DEFAULT_PASSWORD


def check_login(unit: str, password: str) -> bool:
    return unit in UNITS and hmac.compare_digest(password.encode(), get_password(unit).encode())


# --- 儲存後端: 有設定 Google Sheets secrets 就用 Google Sheets,否則用本機 SQLite ---
GSHEET_TAB = "records"
GSHEET_HEADER = ["id", "unit", "bed", "saved_at", "data"]
GSHEET_CACHE_TTL = 30  # 秒;避免每次畫面更新都讀 Google Sheets (API 讀取上限約 60 次/分鐘)


def use_gsheets() -> bool:
    try:
        return "gcp_service_account" in st.secrets and "gsheet" in st.secrets
    except Exception:
        return False


def storage_name() -> str:
    return "Google Sheets" if use_gsheets() else "本機 SQLite"


@st.cache_resource(show_spinner=False)
def _gsheet_ws():
    import gspread  # 只有使用 Google Sheets 時才需要安裝
    gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    cfg = st.secrets["gsheet"]
    sh = gc.open_by_url(cfg["url"]) if "url" in cfg else gc.open_by_key(cfg["key"])
    try:
        ws = sh.worksheet(GSHEET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=GSHEET_TAB, rows=1000, cols=len(GSHEET_HEADER))
        ws.append_row(GSHEET_HEADER)
    return ws


@st.cache_data(ttl=GSHEET_CACHE_TTL, show_spinner=False)
def _gsheet_rows() -> list:
    """[(sheet 列號, id, unit, bed, saved_at, data), ...]"""
    values = _gsheet_ws().get_all_values()
    rows = []
    for i, v in enumerate(values[1:], start=2):
        v = (v + [""] * 5)[:5]
        try:
            rows.append((i, int(v[0]), v[1], v[2], v[3], v[4]))
        except ValueError:
            continue
    return rows


def _db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.execute("""CREATE TABLE IF NOT EXISTS records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        unit TEXT NOT NULL, bed TEXT NOT NULL, saved_at TEXT NOT NULL, data TEXT NOT NULL)""")
    con.execute("CREATE INDEX IF NOT EXISTS idx_bed ON records(unit, bed, id)")
    return con


def _unit_rows(unit: str) -> list:
    """該單位所有紀錄 [(id, bed, saved_at, data_json), ...],由舊到新。"""
    if use_gsheets():
        rows = [(i, b, t, d) for _, i, u, b, t, d in _gsheet_rows() if u == unit]
        return sorted(rows)
    with closing(_db()) as con:
        return con.execute("SELECT id, bed, saved_at, data FROM records WHERE unit=? "
                           "ORDER BY id", (unit,)).fetchall()


def save_record(unit: str, bed: str, data: dict) -> str:
    now = datetime.now(TW_TZ)
    saved_at = now.strftime("%Y-%m-%d %H:%M")
    payload = json.dumps(data, ensure_ascii=False)
    if use_gsheets():
        rec_id = int(now.timestamp() * 1000)  # 以毫秒時間當 id,免讀取即可新增
        _gsheet_ws().append_row([rec_id, unit, bed, saved_at, payload],
                                value_input_option="RAW")
        _gsheet_rows.clear()
    else:
        with closing(_db()) as con, con:
            con.execute("INSERT INTO records (unit, bed, saved_at, data) VALUES (?, ?, ?, ?)",
                        (unit, bed, saved_at, payload))
    return saved_at


def list_records(unit: str, bed: str) -> list:
    """[(id, saved_at), ...] 由新到舊。"""
    return [(i, t) for i, b, t, _ in reversed(_unit_rows(unit)) if b == bed]


def load_record(unit: str, rec_id: int) -> dict:
    for i, _, _, d in _unit_rows(unit):
        if i == rec_id:
            return json.loads(d)
    return {}


def latest_by_bed(unit: str) -> dict:
    """{bed: (saved_at, data)} 每床最近一筆。"""
    out = {}
    for _, b, t, d in _unit_rows(unit):   # 由舊到新,後者覆蓋前者
        out[b] = (t, d)
    return {b: (t, json.loads(d)) for b, (t, d) in out.items()}


def clear_bed(unit: str, bed: str):
    if use_gsheets():
        _gsheet_rows.clear()
        ws = _gsheet_ws()
        targets = [r for r, _, u, b, _, _ in _gsheet_rows() if u == unit and b == bed]
        for r in sorted(targets, reverse=True):   # 由下往上刪,列號才不會位移
            ws.delete_rows(r)
        _gsheet_rows.clear()
    else:
        with closing(_db()) as con, con:
            con.execute("DELETE FROM records WHERE unit=? AND bed=?", (unit, bed))


# ----------------------------------------------------------------------
# 核心計算 (與 Excel 公式一致)
# ----------------------------------------------------------------------
def calc_nicu(p: dict) -> dict:
    """p: 所有輸入 (BW 以 g)。回傳所有計算結果。"""
    bw_kg = p["bw_g"] / 1000
    q = p["q_h"]
    r = {}

    # --- 餵食 (每餐 mL -> mL/day, kcal/day) ---
    feeds = {
        "PF": (p["pf_vol"], 0.8 * p["pf_conc"] / 16),
        "Similac PF": (p["spf_vol"], p["spf_kcal30"] / 30),
        "MF": (p["mf_vol"], 2 / 3 * p["mf_conc"] / 14),
        "HM": (p["hm_vol"], 0.67 + p["hm_pk"] * 0.13),
    }
    feed_rows = {}
    for name, (vol, kcal_ml) in feeds.items():
        ml_day = vol * 24 / q
        kcal = ml_day * kcal_ml
        feed_rows[name] = (ml_day, kcal, kcal * 1000 / bw_kg / 2 / 4 / 1440)
    r["feed_rows"] = feed_rows
    r["feed_ml"] = sum(v[0] for v in feed_rows.values())
    r["feed_kcal"] = sum(v[1] for v in feed_rows.values())
    r["feed_sir"] = sum(v[2] for v in feed_rows.values())

    # --- IV-1 / IV-2 ---
    for k in ("iv1", "iv2"):
        rate, glu = p[f"{k}_rate"], p[f"{k}_glu"]
        r[f"{k}_ml"] = rate * 24
        r[f"{k}_kcal"] = rate * 24 * glu / 100 * GLU_KCAL_PER_G
        r[f"{k}_sir"] = rate * glu / bw_kg / 6
        r[f"{k}_na"] = r[f"{k}_ml"] * p[f"{k}_na"] / 1000 / bw_kg
        r[f"{k}_k"] = r[f"{k}_ml"] * p[f"{k}_k"] / 1000 / bw_kg

    # --- PN + Fat ---
    pn_ml = p["pn_rate"] * 24
    r["pn_ml"] = pn_ml
    r["pn_kcal_glu"] = pn_ml * p["pn_glu"] / 100 * GLU_KCAL_PER_G
    r["pn_kcal_aa"] = pn_ml * p["pn_aa"] / 100 * AA_KCAL_PER_G
    r["pn_kcal"] = r["pn_kcal_glu"] + r["pn_kcal_aa"]
    r["pn_sir"] = p["pn_glu"] * p["pn_rate"] / bw_kg / 6
    r["fat_kcal"] = p["fat_ml"] * FAT20_KCAL_PER_ML
    r["pn_osm"] = p["pn_glu"] * 50 + p["pn_aa"] * 100 + p["pn_na"] * 2 + p["pn_k"] * 2

    # --- 總量 ---
    r["total_ml"] = r["feed_ml"] + r["iv1_ml"] + pn_ml + p["fat_ml"] + r["iv2_ml"]
    r["total_kcal"] = (r["feed_kcal"] + r["iv1_kcal"] + r["pn_kcal"]
                       + r["fat_kcal"] + r["iv2_kcal"])
    r["tf"] = r["total_ml"] / bw_kg                 # mL/kg/day
    r["cal"] = r["total_kcal"] / bw_kg              # kcal/kg/day (吃 + 打)
    r["sir"] = r["feed_sir"] + r["iv1_sir"] + r["pn_sir"] + r["iv2_sir"]

    r["aa_pn"] = p["pn_aa"] * pn_ml / 100 / bw_kg   # g/kg/day (只計 PN)
    r["fat_kg"] = p["fat_ml"] * FAT20_G_PER_ML / bw_kg
    tc_iv = r["iv1_kcal"] + r["iv2_kcal"] + r["pn_kcal"] + r["fat_kcal"]  # TC 只計 PN + IVF
    r["aa_tc"] = r["pn_kcal_aa"] / tc_iv if tc_iv else 0
    r["fat_tc"] = r["fat_kcal"] / tc_iv if tc_iv else 0

    # --- 奶水成分小計 (per kg/day) ---
    milk_ml = p["milk_ml_day"]
    ratio = p["milk_ratio"]
    milk = {}
    for i, n in enumerate(NUTRIENTS):
        amt = (p["milk_comp"][i] / 1000 * milk_ml
               + p["fort_comp"][i] / 100 * milk_ml * p["milk_pk"]) * ratio / bw_kg
        milk[n] = amt
    r["milk"] = milk
    milk_na = milk["Sodium (mg)"] / 23     # mEq/kg/day
    milk_k = milk["Potassium (mg)"] / 39

    # --- 電解質 ---
    r["pn_na"] = pn_ml * p["pn_na"] / 1000 / bw_kg
    r["pn_k"] = pn_ml * p["pn_k"] / 1000 / bw_kg
    r["na_iv_pn"] = r["iv1_na"] + r["pn_na"] + r["iv2_na"]
    r["k_iv_pn"] = r["iv1_k"] + r["pn_k"] + r["iv2_k"]
    r["na_total"] = r["na_iv_pn"] + milk_na
    r["k_total"] = r["k_iv_pn"] + milk_k
    pn_l = pn_ml / 1000
    r["pn_mg_mg"] = p["pn_mg"] * 24 / 2 * pn_l / bw_kg
    r["pn_ca_mg"] = p["pn_ca"] * 40 / 2 * pn_l / bw_kg
    r["pn_p_mg"] = p["pn_p"] * 31 * pn_l / bw_kg
    r["ca_total"] = r["pn_ca_mg"] + milk["Calcium (mg)"]
    r["p_total"] = r["pn_p_mg"] + milk["Phosphorus (mg)"]
    r["cap_ratio"] = r["ca_total"] / r["p_total"] if r["p_total"] else 0
    r["protein"] = r["aa_pn"] + milk["Protein (g)"]

    # --- Fe (奶水 + 鐵劑) ---
    # Excel J20 誤引用奶水 Na (I51),此處改用奶水 Fe
    r["fe_drops"] = p["fe_gtt"] * p["fe_times"] * FE_MG_PER_GTT / bw_kg
    r["fe_total"] = milk["Iron (mg)"] + r["fe_drops"]

    # --- PN 固定添加 (依實際輸注量 / 實報 PN 量 比例) ---
    bag = p["pn_bag_ml"] if p["pn_bag_ml"] > 0 else pn_ml
    frac = pn_ml / bag if bag else 0
    r["pn_frac"] = frac
    r["add_zn"] = 1.1 * frac / bw_kg                   # mg/kg/day
    r["add_vita"] = 200 * frac / bw_kg                 # IU/kg/day
    r["add_vitd"] = 200 * frac                         # IU/day
    r["add_vite"] = 1 * frac / bw_kg                   # mg/kg/day
    r["add_vitc"] = 100 * frac / bw_kg                 # mg/kg/day
    r["zn_total"] = (r["add_zn"] + milk["Zinc (mg)"]) * 1000  # mcg/kg/day
    return r


def fmt(x) -> str:
    """數字去尾零 (仿 Excel 字串串接)。"""
    return f"{x:g}"


def summary_strings(p: dict, r: dict) -> dict:
    """Excel「今日結果」列 (N3:W3)。"""
    q = fmt(p["q_h"])
    feed = []
    if p["hm_vol"]:
        feed.append(f"HM {fmt(p['hm_vol'])}ml q{q}h" if not p["hm_pk"]
                    else f"HM+{fmt(p['hm_pk'])} pk fortifier {fmt(p['hm_vol'])}ml q{q}h")
    if p["pf_vol"]:
        feed.append(f"{fmt(p['pf_conc'])} % PF {fmt(p['pf_vol'])} ml q{q}h")
    if p["spf_vol"]:
        feed.append(f"SimilacPF ({fmt(p['spf_kcal30'])} Kcal) {fmt(p['spf_vol'])} ml q{q}h")
    if p["mf_vol"]:
        feed.append(f"{fmt(p['mf_conc'])} % MF {fmt(p['mf_vol'])} ml q{q}h")
    ivf = []
    for k in ("iv1", "iv2"):
        if p[f"{k}_rate"]:
            ivf.append(f"{p[f'{k}_name']} run {fmt(p[f'{k}_rate'])} ml/hr")
    return {
        "Feeding": ", ".join(feed) if feed else "NPO",
        "IVF": "; ".join(ivf) if ivf else "-",
        "AA": f"{r['aa_pn']:.2f}/{r['protein']:.2f}",
        "PN": f"{fmt(p['pn_rate'])}({fmt(p['pn_glu'])}/{fmt(p['pn_aa'])}/{r['aa_tc'] * 100:.1f})",
        "Fat": f"{fmt(p['fat_ml'])}({r['fat_kg']:.2f}/{r['fat_tc'] * 100:.2f})",
        "TF": f"{r['tf']:.0f}",
        "Cal": f"{r['cal']:.0f}",
        "SIR": f"{r['sir']:.2f}",
    }


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
DEFAULTS = {
    "bed": "", "bw_g": 2400, "q_h": 3,
    "pf_conc": 16, "pf_vol": 0, "spf_kcal30": 26, "spf_vol": 0,
    "mf_conc": 14, "mf_vol": 20, "hm_pk": 0, "hm_vol": 0,
    "iv1_preset": "自訂", "iv1_name": "", "iv1_glu": 10.0, "iv1_rate": 0.0,
    "iv1_na": 0.0, "iv1_k": 0.0,
    "iv2_preset": "自訂", "iv2_name": "", "iv2_glu": 0.0, "iv2_rate": 0.0,
    "iv2_na": 0.0, "iv2_k": 0.0,
    "pn_preset": "自訂", "pn_glu": 10.0, "pn_aa": 2.5, "pn_rate": 8.5, "fat_ml": 20,
    "pn_na": 20.7, "pn_k": 22.0, "pn_mg": 3.3, "pn_ca": 10.0, "pn_p": 5.0, "pn_bag_ml": 0,
    "milk_type": "Enfalac LBW 優生早產兒", "fort_type": "Similac HMF 心美力母乳營養補充品",
    "milk_link": True, "milk_conc": 14, "milk_std": 14, "milk_vol": 20,
    "milk_q": 3, "milk_pk": 0, "fe_gtt": 0, "fe_times": 0,
}

# 輸入精度:整數欄位 = DEFAULTS 中為 int 者 (另加 RI 的 U 數);
# 一位小數 = IV / PN glucose、PN AA、電解質、各滴速 (mL/hr)
INT_KEYS = {k for k, v in DEFAULTS.items() if type(v) is int} | {"ri_u"}
DEC2_KEYS = {"ri_tgt"}   # RI 目標劑量 (U/kg/hr) 需兩位小數


def to_int(v, default=0) -> int:
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return default


def to_dec1(v, default=0.0) -> float:
    try:
        return round(float(v), 1)
    except (TypeError, ValueError):
        return default


def apply_ivf(k: str):
    name = st.session_state[f"{k}_preset"]
    vals = IVF_TABLE.get(name)
    if vals:
        st.session_state[f"{k}_na"], st.session_state[f"{k}_k"], st.session_state[f"{k}_glu"] = (
            float(v) for v in vals)
        st.session_state[f"{k}_name"] = name


def apply_pn():
    label = st.session_state["pn_preset"]
    if label == "自訂":
        return
    row = PN_SETS.iloc[pn_labels().index(label) - 1]
    for key, col in [("pn_glu", "Glucose (%)"), ("pn_aa", "AA (%)"), ("pn_na", "Na"),
                     ("pn_k", "K"), ("pn_mg", "Mg"), ("pn_ca", "Ca"), ("pn_p", "P")]:
        st.session_state[key] = float(row[col])


def pn_labels():
    return ["自訂"] + [f"{g:g}% / AA {a:g}%" for g, a in
                      zip(PN_SETS["Glucose (%)"], PN_SETS["AA (%)"])]


def apply_record(data: dict):
    """把紀錄寫回輸入欄位 (空 dict = 重設為預設值)。只能在 callback 中呼叫。"""
    ss = st.session_state
    valid = {"iv1_preset": IVF_TABLE, "iv2_preset": IVF_TABLE, "pn_preset": pn_labels(),
             "milk_type": MILK_TABLE, "fort_type": FORTIFIER_TABLE}
    for k, default in DEFAULTS.items():
        v = data.get(k, default)
        if k in valid and v not in valid[k]:
            v = default
        if type(default) is int:
            v = to_int(v, default)
        elif type(default) is float:
            v = to_dec1(v, default)
        ss[k] = v
    ss["comp_override"] = ({"milk_type": ss["milk_type"], "fort_type": ss["fort_type"],
                            "milk": data.get("milk_comp"), "fort": data.get("fort_comp")}
                           if data else None)
    ss["load_n"] = ss.get("load_n", 0) + 1   # 讓奶水成分表重新載入


def load_latest():
    ss = st.session_state
    ss["clear_pending"] = None
    unit, bed = ss.get("unit"), ss.get("bed_sel")
    recs = list_records(unit, bed) if unit and bed else []
    if recs:
        apply_record(load_record(unit, recs[0][0]))
        ss["loaded_info"] = f"{bed}:已載入 {recs[0][1]} 的紀錄"
    else:
        apply_record({})
        ss["loaded_info"] = f"{bed}:尚無紀錄,已重設為預設值"


def load_selected_history():
    ss = st.session_state
    unit, bed = ss["unit"], ss["bed_sel"]
    rec_id = ss.get(f"hist_{unit}_{bed}")
    if rec_id is not None:
        apply_record(load_record(unit, rec_id))
        t = dict(list_records(unit, bed)).get(rec_id, "")
        ss["loaded_info"] = f"{bed}:已載入 {t} 的紀錄"


def do_login():
    ss = st.session_state
    unit = ss.get("login_user", "").strip().upper()
    if check_login(unit, ss.get("login_pw", "")):
        ss["unit"] = unit
        ss["bed_sel"] = UNITS[unit][0]
        ss["login_error"] = ""
        load_latest()
    else:
        ss["login_error"] = "帳號或密碼錯誤"
    ss["login_pw"] = ""


def do_logout():
    ss = st.session_state
    ss.pop("unit", None)
    apply_record({})          # 登出後清空畫面上的病人資料
    ss["loaded_info"] = ""


def ask_clear_bed():
    st.session_state["clear_pending"] = st.session_state.get("bed_sel")


def cancel_clear_bed():
    st.session_state["clear_pending"] = None


def do_clear_bed():
    """刪除該床所有紀錄,並把畫面輸入重設為預設值。"""
    ss = st.session_state
    bed = ss["bed_sel"]
    n = len(list_records(ss["unit"], bed))
    if n:
        clear_bed(ss["unit"], bed)
    ss["clear_pending"] = None
    apply_record({})
    ss["loaded_info"] = ""
    ss["flash"] = f"{bed} 已清空" + (f" (刪除 {n} 筆紀錄)" if n else " (畫面已重設)")


def do_reset_inputs():
    """未登入時:只把畫面輸入重設為預設值。"""
    apply_record({})


def num(container, label, key, step=None, min_value=0, max_value=None, **kw):
    """依欄位決定精度:整數 / 一位小數 / (RI 目標) 兩位小數。"""
    if key in INT_KEYS:
        if "value" in kw:
            kw["value"] = to_int(kw["value"])
        return container.number_input(
            label, min_value=int(min_value), step=max(1, int(step or 1)), format="%d",
            max_value=None if max_value is None else int(max_value), key=key, **kw)
    if "value" in kw:
        kw["value"] = float(kw["value"])
    if key in DEC2_KEYS:
        kw.setdefault("format", "%.2f")
        step = step or 0.01
    else:
        kw["format"] = "%.1f"
        step = max(0.1, float(step or 0.1))
    return container.number_input(
        label, min_value=float(min_value), step=float(step),
        max_value=None if max_value is None else float(max_value), key=key, **kw)


def status(ok: bool) -> str:
    return "✅" if ok else "⚠️"


def main():
    _logo = Path(__file__).parent / "logo.png"
    st.set_page_config(page_title="NICU 營養計算器",
                       page_icon=str(_logo) if _logo.exists() else "🍼", layout="wide",
                       initial_sidebar_state="expanded")
    st.markdown(
        """
        <style>
        [data-testid="stSidebarCollapseButton"],
        [data-testid="stSidebarHeader"] button { display: none !important; }
        [data-testid="stSidebarHeader"] { display: none !important; }
        [data-testid="stSidebarUserContent"] { padding-top: 1rem !important; }
        .block-container,
        [data-testid="stMainBlockContainer"] { padding-top: 1.5rem !important; }
        header[data-testid="stHeader"] { background: transparent; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    logo_html = ""
    if _logo.exists():
        logo_b64 = base64.b64encode(_logo.read_bytes()).decode()
        logo_html = (f"<img src='data:image/png;base64,{logo_b64}' "
                     "style='height:64px; vertical-align:middle; margin-right:12px;'>")
    st.markdown(
        f"<h1 style='margin-bottom:0; display:flex; align-items:center; flex-wrap:wrap;'>"
        f"{logo_html}NICU 營養計算器</h1>",
        unsafe_allow_html=True,
    )

    ss = st.session_state
    for k, v in DEFAULTS.items():
        ss.setdefault(k, v)

    # --- 登入 / 床位 ---
    sb = st.sidebar
    unit = ss.get("unit")
    save_slot = None
    if unit:
        beds = UNITS[unit]
        latest = latest_by_bed(unit)
        a, b = sb.columns([3, 2])
        a.markdown(f"**單位:{unit}**")
        b.button("登出", on_click=do_logout, use_container_width=True)
        sb.caption(f"儲存位置:{storage_name()}")

        def bed_label(bed):
            if bed in latest:
                t, d = latest[bed]
                return f"{bed} ● {d.get('bw_g', 0):g} g · {t[5:]}"
            return f"{bed} (空床)"

        sb.selectbox("床位", beds, key="bed_sel", format_func=bed_label,
                     on_change=load_latest)
        bed = ss["bed_sel"]
        recs = list_records(unit, bed)
        if recs:
            rec_times = dict(recs)
            c1, c2 = sb.columns([3, 2])
            c1.selectbox("歷次紀錄", [r_[0] for r_ in recs], key=f"hist_{unit}_{bed}",
                         format_func=lambda i: rec_times.get(i, ""),
                         label_visibility="collapsed")
            c2.button("載入", on_click=load_selected_history, use_container_width=True)
        save_slot = sb.container()
        if ss.get("loaded_info"):
            sb.caption(ss["loaded_info"])
        if ss.get("clear_pending") == bed:
            sb.warning(f"確定要清空 {bed}?"
                       + (f"將刪除 {len(recs)} 筆紀錄,無法復原。" if recs else "此床尚無紀錄,只會重設畫面。"))
            c1, c2 = sb.columns(2)
            c1.button("確認清空", on_click=do_clear_bed, type="primary",
                      use_container_width=True)
            c2.button("取消", on_click=cancel_clear_bed, use_container_width=True)
        else:
            sb.button(f"🗑️ 清空 {bed} 資料", on_click=ask_clear_bed,
                      use_container_width=True)
    else:
        with sb.expander("🔐 登入 (儲存 / 讀取床位紀錄)"):
            with st.form("login_form", clear_on_submit=False):
                st.text_input("帳號 (PICU / NBC)", key="login_user")
                st.text_input("密碼", type="password", key="login_pw")
                st.form_submit_button("登入", on_click=do_login, use_container_width=True)
            if ss.get("login_error"):
                st.error(ss["login_error"])
        sb.caption("未登入也可直接計算,但不會儲存紀錄")
        sb.button("🗑️ 清空輸入 (重設為預設值)", on_click=do_reset_inputs,
                  use_container_width=True)
    if ss.get("flash"):
        sb.success(ss.pop("flash"))

    # --- 病人資料 ---
    sb.header("病人資料")
    sb.text_input("備註 (請勿輸入姓名 / 病歷號)", key="bed")
    num(sb, "體重 BW (g)", "bw_g", step=10.0, min_value=300.0, max_value=8000.0)
    num(sb, "餵食間隔 q (h)", "q_h", step=1.0, min_value=1.0, max_value=24.0)
    sb.caption("IV / PN 速率單位為 mL/hr (微滴 gtt/min = mL/hr)")

    tab_names = ["營養計算", "RI 計算", "參考表"] + (["床位總覽"] if unit else [])
    tab_main, tab_ri, tab_ref, *tab_beds = st.tabs(tab_names)

    if unit:
        with tab_beds[0]:
            rows = []
            for b_ in UNITS[unit]:
                if b_ in latest:
                    t, d = latest[b_]
                    sm = d.get("summary", {})
                    rows.append({"床位": b_, "上次儲存": t, "BW (g)": d.get("bw_g"),
                                 "備註": d.get("bed", ""), **sm})
                else:
                    rows.append({"床位": b_, "上次儲存": "—"})
            st.caption(f"{unit} 各床最近一次紀錄 (於側邊欄選擇床位即可載入)")
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True,
                         height=min(40 + 35 * len(rows), 950))

    with tab_main:
        top = st.container()
        st.divider()

        # --- 餵食 ---
        st.subheader("餵食 Feeding (每餐 mL)")
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            st.markdown("**PF 早產兒配方**")
            num(st, "濃度 (%) (標準 16)", "pf_conc", step=1.0)
            num(st, "每餐 (mL)", "pf_vol", step=1.0, help="kcal/mL = 0.8 × 濃度/16")
        with f2:
            st.markdown("**Similac PF (26 kcal 以上)**")
            num(st, "Kcal / 30 mL", "spf_kcal30", step=1.0)
            num(st, "每餐 (mL) ", "spf_vol", step=1.0)
        with f3:
            st.markdown("**MF 一般配方**")
            num(st, "濃度 (%) (標準 14)", "mf_conc", step=1.0)
            num(st, "每餐 (mL)  ", "mf_vol", step=1.0, help="kcal/mL = 2/3 × 濃度/14")
        with f4:
            st.markdown("**HM 母乳 (+ fortifier)**")
            num(st, "pk fortifier", "hm_pk", step=1.0, help="kcal/mL = 0.67 + 包數 × 0.13")
            num(st, "每餐 (mL)   ", "hm_vol", step=1.0)

        # --- IV / PN ---
        st.subheader("IV / PN")
        c_iv, c_pn, c_fat = st.columns([1, 1.4, 1])
        with c_iv:
            st.markdown("**IV-1**")
            st.selectbox("IV 種類", list(IVF_TABLE), key="iv1_preset",
                         on_change=apply_ivf, args=("iv1",))
            st.text_input("IV 名稱 (醫囑用)", key="iv1_name")
            num(st, "Glucose (%)", "iv1_glu", step=0.5)
            num(st, "速率 (mL/hr)", "iv1_rate", step=0.1)
            a, b = st.columns(2)
            num(a, "Na (mEq/L)", "iv1_na", step=1.0)
            num(b, "K (mEq/L)", "iv1_k", step=1.0)
        with c_pn:
            st.markdown("**PN**")
            st.selectbox("本院 PN 套餐", pn_labels(), key="pn_preset", on_change=apply_pn)
            a, b, c = st.columns(3)
            num(a, "PN Glu (%)", "pn_glu", step=0.5)
            num(b, "PN AA (%)", "pn_aa", step=0.5)
            num(c, "速率 (mL/hr)", "pn_rate", step=0.1)
            a, b, c = st.columns(3)
            num(a, "Na (mEq/L)", "pn_na", step=0.1)
            num(b, "K (mEq/L)", "pn_k", step=0.1)
            num(c, "Mg (mEq/L)", "pn_mg", step=0.1)
            a, b, c = st.columns(3)
            num(a, "Ca (mEq/L)", "pn_ca", step=0.5)
            num(b, "P (mEq/L)", "pn_p", step=0.5, help="Excel 以 × 31 換算為 mg")
            num(c, "實報 PN 量 (mL)", "pn_bag_ml", step=10.0,
                help="用於計算 PN 固定添加 (Zn、維生素) 實際輸入比例;0 = 同每日 PN 量")
        with c_fat:
            st.markdown("**Fat — IF 20%**")
            num(st, "mL qd", "fat_ml", step=1.0, help="0.2 g/mL、2 kcal/mL")
            with st.expander("第三條 IV line (IV-2)", expanded=ss["iv2_rate"] > 0):
                st.selectbox("IV-2 種類", list(IVF_TABLE), key="iv2_preset",
                             on_change=apply_ivf, args=("iv2",))
                st.text_input("IV-2 名稱", key="iv2_name")
                num(st, "Glucose (%) ", "iv2_glu", step=0.5)
                num(st, "速率 (mL/hr) ", "iv2_rate", step=0.1)
                num(st, "Na (mEq/L) ", "iv2_na", step=1.0)
                num(st, "K (mEq/L) ", "iv2_k", step=1.0)

        # --- 奶水成分 ---
        with st.expander("奶水成分 (計算 Protein / Ca / P / Na / K / Fe / Zn) 與鐵劑",
                         expanded=False):
            a, b = st.columns(2)
            a.selectbox("奶水種類", list(MILK_TABLE), key="milk_type")
            b.selectbox("母乳添加劑", list(FORTIFIER_TABLE), key="fort_type")
            st.checkbox("奶量同上方餵食總量", key="milk_link")
            a, b, c, d, e = st.columns(5)
            num(a, "奶水濃度 (%)", "milk_conc", step=1.0,
                help="與標準濃度的比例會乘到所有成分;母乳 / 現成奶水請兩欄填相同值")
            num(b, "標準濃度 (PF 16 / MF 14)", "milk_std", step=1.0, min_value=1.0)
            num(c, "添加劑包數", "milk_pk", step=1.0)
            num(d, "每餐 (mL)", "milk_vol", step=1.0, disabled=ss["milk_link"])
            num(e, "q (h)", "milk_q", step=1.0, min_value=1.0, disabled=ss["milk_link"])
            milk_vals = MILK_TABLE[ss["milk_type"]][1:]
            fort_vals = FORTIFIER_TABLE[ss["fort_type"]][1:]
            ov = ss.get("comp_override")   # 由紀錄載入、可能經手動修改的成分
            if ov and ov.get("milk_type") == ss["milk_type"] and ov.get("milk"):
                milk_vals = ov["milk"]
            if ov and ov.get("fort_type") == ss["fort_type"] and ov.get("fort"):
                fort_vals = ov["fort"]
            comp = pd.DataFrame({
                "成分": NUTRIENTS,
                "奶水 (每公升)": [float(v) for v in milk_vals],
                "添加劑 (每 100 mL)": [float(v) for v in fort_vals],
            })
            comp = st.data_editor(
                comp, hide_index=True, use_container_width=True, disabled=["成分"],
                key=f"comp_{ss['milk_type']}_{ss['fort_type']}_{ss.get('load_n', 0)}")
            st.markdown("**鐵劑** (1 gtt = 2.5 mg)")
            a, b = st.columns(2)
            num(a, "每次 gtt", "fe_gtt", step=1.0)
            num(b, "一天次數", "fe_times", step=1.0)

        # --- 計算 ---
        p = {k: (to_dec1(ss[k]) if type(DEFAULTS[k]) is float else ss[k]) for k in DEFAULTS}
        feed_ml_day = sum(ss[k] for k in ("pf_vol", "spf_vol", "mf_vol", "hm_vol")) * 24 / ss["q_h"]
        p["milk_ml_day"] = feed_ml_day if ss["milk_link"] else ss["milk_vol"] * 24 / ss["milk_q"]
        p["milk_ratio"] = ss["milk_conc"] / ss["milk_std"]
        p["milk_comp"] = comp["奶水 (每公升)"].fillna(0).tolist()
        p["fort_comp"] = comp["添加劑 (每 100 mL)"].fillna(0).tolist()
        p["iv1_name"] = ss["iv1_name"] or "IVF"
        p["iv2_name"] = ss["iv2_name"] or "IV-2"
        r = calc_nicu(p)
        s = summary_strings(p, r)

        # --- 儲存 (需登入) ---
        if save_slot is not None:
            if save_slot.button(f"💾 儲存 {ss['bed_sel']} 紀錄", type="primary",
                                use_container_width=True):
                data = {k: (to_dec1(ss[k]) if type(DEFAULTS[k]) is float else ss[k])
                        for k in DEFAULTS}
                data["milk_comp"] = p["milk_comp"]
                data["fort_comp"] = p["fort_comp"]
                data["summary"] = s
                t = save_record(ss["unit"], ss["bed_sel"], data)
                ss["flash"] = f"已儲存 {ss['bed_sel']} ({t})"
                ss["loaded_info"] = ""
                st.rerun()

        # --- 頂部摘要 ---
        with top:
            if unit:
                st.markdown(f"#### 🛏️ {unit} · {ss['bed_sel']}"
                            + (f" — {ss['bed']}" if ss["bed"] else ""))
            m = st.columns(5)
            m[0].metric("TF (mL/kg/day)", f"{r['tf']:.0f}")
            m[1].metric("Cal* (kcal/kg/day)", f"{r['cal']:.1f}")
            m[2].metric("SIR (mg/kg/min)", f"{r['sir']:.2f}")
            m[3].metric("Protein (g/kg/day)", f"{r['protein']:.2f}",
                        f"PN AA {r['aa_pn']:.2f}", delta_color="off")
            m[4].metric("Fat (g/kg/day)", f"{r['fat_kg']:.2f}")
            st.caption("*Cal 為所有吃 + 打的熱量")

        st.divider()
        st.subheader("計算結果")
        left, right = st.columns([3, 2])
        with left:
            rows = []
            for name, (ml, kcal, sir) in r["feed_rows"].items():
                if ml:
                    rows.append([name, ml, kcal, sir])
            if r["iv1_ml"]:
                rows.append(["IV-1", r["iv1_ml"], r["iv1_kcal"], r["iv1_sir"]])
            if r["pn_ml"]:
                rows.append(["PN", r["pn_ml"], r["pn_kcal"], r["pn_sir"]])
            if ss["fat_ml"]:
                rows.append(["IF 20%", ss["fat_ml"], r["fat_kcal"], 0.0])
            if r["iv2_ml"]:
                rows.append(["IV-2", r["iv2_ml"], r["iv2_kcal"], r["iv2_sir"]])
            rows.append(["合計", r["total_ml"], r["total_kcal"], r["sir"]])
            detail = pd.DataFrame(rows, columns=["來源", "mL/day", "kcal/day", "SIR (mg/kg/min)"])
            detail["mL/kg/day"] = detail["mL/day"] / (ss["bw_g"] / 1000)
            detail["kcal/kg/day"] = detail["kcal/day"] / (ss["bw_g"] / 1000)
            st.dataframe(
                detail[["來源", "mL/day", "mL/kg/day", "kcal/day", "kcal/kg/day",
                        "SIR (mg/kg/min)"]].style.format(
                    {"mL/day": "{:.1f}", "mL/kg/day": "{:.1f}", "kcal/day": "{:.1f}",
                     "kcal/kg/day": "{:.1f}", "SIR (mg/kg/min)": "{:.2f}"}),
                hide_index=True, use_container_width=True)
        with right:
            st.markdown("**PN 醫囑**")
            st.code(f"({fmt(ss['pn_glu'])}%/{fmt(ss['pn_aa'])}%) PN run {fmt(ss['pn_rate'])}ml qh"
                    f" , fat run {fmt(ss['fat_ml'])} ml qd", language=None)
            st.metric("PN 滲透壓 (mOsm/L)", f"{r['pn_osm']:.0f}")
            if r["pn_ml"] and r["pn_osm"] > PERIPHERAL_OSM:
                st.info(f"PN 滲透壓 > {PERIPHERAL_OSM} mOsm/L,經周邊靜脈給予時請留意 "
                        "(一般建議;Excel 未設此門檻)")

        # --- 目標範圍對照 ---
        st.subheader("目標範圍對照")
        has_pn = r["pn_ml"] > 0
        checks = [
            ("AA/TC** (%)", r["aa_tc"] * 100, "< 15", r["aa_tc"] < REF["aa_tc"]),
            ("Fat/TC** (%)", r["fat_tc"] * 100, "< 30", r["fat_tc"] < REF["fat_tc"]),
            ("Na: IV + PN (mEq/kg/day)", r["na_iv_pn"], "2 – 4", None),
            ("K: IV + PN (mEq/kg/day)", r["k_iv_pn"], "2 – 3", None),
            ("Na: 奶水 + IV + PN (mEq/kg/day)", r["na_total"], "2 – 4",
             REF["na"][0] <= r["na_total"] <= REF["na"][1]),
            ("K: 奶水 + IV + PN (mEq/kg/day)", r["k_total"], "2 – 3",
             REF["k"][0] <= r["k_total"] <= REF["k"][1]),
            ("PN Mg (mg/kg/day)", r["pn_mg_mg"], "—", None),
            ("PN Ca (mg/kg/day)", r["pn_ca_mg"], "60", None),
            ("PN P (mg/kg/day)", r["pn_p_mg"], "45 – 60", None),
            ("Ca: 奶水 + PN (mg/kg/day)", r["ca_total"], "200", None),
            ("P: 奶水 + PN (mg/kg/day)", r["p_total"], "100", None),
            ("Ca : P", r["cap_ratio"], "1.3 – 1.7",
             REF["cap_ratio"][0] <= r["cap_ratio"] <= REF["cap_ratio"][1]
             if r["p_total"] else None),
            ("Fe: 奶水 + 鐵劑 (mg/kg/day)", r["fe_total"], "2", None),
            ("Zn: 奶水 + PN (mcg/kg/day)", r["zn_total"], "400", None),
        ]
        ref_df = pd.DataFrame(
            [[n, round(v, 2), rng, "" if ok is None else status(ok)] for n, v, rng, ok in checks],
            columns=["項目", "目前", "建議", ""])
        a, b = st.columns([3, 2])
        a.dataframe(ref_df, hide_index=True, use_container_width=True, height=530)
        a.caption("**TC 只計 PN + IVF (不含奶水)")
        with b:
            st.markdown("**奶水成分小計** (per kg/day)")
            milk_df = pd.DataFrame({
                "成分": NUTRIENTS,
                "小計": list(r["milk"].values()),
                "單位": ["g", "mg", "mg", "mg", "mg", "mg", "mg"],
            })
            milk_df.loc[3, ["小計", "單位"]] = [r["milk"]["Sodium (mg)"] / 23, "mEq"]
            milk_df.loc[4, ["小計", "單位"]] = [r["milk"]["Potassium (mg)"] / 39, "mEq"]
            st.dataframe(milk_df.style.format({"小計": "{:.2f}"}), hide_index=True,
                         use_container_width=True)
            st.caption(f"奶量 {p['milk_ml_day']:.0f} mL/day,濃度比例 {p['milk_ratio']:.2f}")
            st.markdown("**本院 PN 固定添加** (實際輸入)")
            add_df = pd.DataFrame({
                "成分": ["Zn", "Vitamin A", "Vitamin D2", "Vitamin E", "Vitamin C"],
                "每包每天": ["1.1 mg", "200 IU", "200 IU", "1 mg", "100 mg"],
                "實際": [r["add_zn"], r["add_vita"], r["add_vitd"], r["add_vite"], r["add_vitc"]],
                "單位": ["mg/kg/day", "IU/kg/day", "IU/day", "mg/kg/day", "mg/kg/day"],
            })
            st.dataframe(add_df.style.format({"實際": "{:.2f}"}), hide_index=True,
                         use_container_width=True)

        warns = []
        if has_pn or r["iv1_ml"] or r["iv2_ml"]:
            if r["aa_tc"] >= REF["aa_tc"]:
                warns.append(f"AA/TC {r['aa_tc'] * 100:.1f}% ≥ 15%")
            if r["fat_tc"] >= REF["fat_tc"]:
                warns.append(f"Fat/TC {r['fat_tc'] * 100:.1f}% ≥ 30%")
        if r["na_total"] and not REF["na"][0] <= r["na_total"] <= REF["na"][1]:
            warns.append(f"Na {r['na_total']:.2f} mEq/kg/day 不在 2–4 範圍")
        if r["k_total"] and not REF["k"][0] <= r["k_total"] <= REF["k"][1]:
            warns.append(f"K {r['k_total']:.2f} mEq/kg/day 不在 2–3 範圍")
        if r["p_total"] and not REF["cap_ratio"][0] <= r["cap_ratio"] <= REF["cap_ratio"][1]:
            warns.append(f"Ca:P {r['cap_ratio']:.2f} 不在 1.3–1.7 範圍")
        for w in warns:
            st.warning("⚠️ " + w)
        if not warns:
            st.success("AA/TC、Fat/TC、Na、K、Ca:P 皆在建議範圍內")

        # --- 今日結果 ---
        st.subheader("今日結果")
        st.dataframe(pd.DataFrame([{"BW": ss["bw_g"], **s}]), hide_index=True,
                     use_container_width=True)
        st.code("\t".join([fmt(ss["bw_g"])] + list(s.values())), language=None)
        st.caption("可直接複製 (Tab 分隔) 貼到 Excel 的 N–W 欄")

    # ------------------------------------------------------------------
    with tab_ri:
        bw_kg = ss["bw_g"] / 1000
        st.caption(f"BW {ss['bw_g']:g} g (由側邊欄設定)")
        st.subheader("RI 濃度 → 劑量")
        a, b, c = st.columns(3)
        ri_u = num(a, "RI (U in 500 mL)", "ri_u", step=1.0, value=10.0)
        ri_glu = num(b, "稀釋液 Glucose (%)", "ri_glu", step=0.5, value=10.0)
        ri_rate = num(c, "速率 (mL/hr)", "ri_rate", step=0.1, value=1.0)
        dose = ri_u / 500 * ri_rate / bw_kg if bw_kg else 0
        g_per_u = 5 * ri_glu / ri_u if ri_u else 0
        a, b = st.columns(2)
        a.metric("RI 劑量 (U/kg/hr)", f"{dose:.3f}")
        b.metric("1 U 對應 Glucose (g)", f"{g_per_u:.2f}")

        st.subheader("目標劑量 → RI 濃度")
        a, b = st.columns(2)
        tgt = num(a, "目標 (U/kg/hr)", "ri_tgt", step=0.01, value=0.05, format="%.2f")
        tgt_rate = num(b, "速率 (mL/hr) ", "ri_tgt_rate", step=0.1, value=1.0, min_value=0.1)
        st.metric("RI 需加入 (U in 500 mL)", f"{tgt * bw_kg / tgt_rate * 500:.1f}")
        tbl = pd.DataFrame({"U/kg/hr": [0.01, 0.03, 0.05, 0.07, 0.1]})
        tbl["U in 500 mL"] = tbl["U/kg/hr"] * bw_kg / tgt_rate * 500
        st.dataframe(tbl.style.format({"U/kg/hr": "{:.2f}", "U in 500 mL": "{:.1f}"}),
                     hide_index=True)

        st.subheader("合併 SIR")
        a, b, c, d = st.columns(4)
        g1 = num(a, "RI 液 Glucose (%)", "sir_g1", step=0.5, value=10.0)
        v1 = num(b, "RI 液速率 (mL/hr)", "sir_v1", step=0.1, value=0.3)
        g2 = num(c, "IVF/PN Glucose (%)", "sir_g2", step=0.5, value=10.0)
        v2 = num(d, "IVF/PN 速率 (mL/hr)", "sir_v2", step=0.1, value=1.0)
        st.metric("SIR (mg/kg/min)", f"{(g1 * v1 + g2 * v2) / 6 / bw_kg:.2f}")

    # ------------------------------------------------------------------
    with tab_ref:
        st.subheader("本院現成 PN 套餐")
        pn_tbl = PN_SETS.copy()
        pn_tbl.insert(2, "Kcal/mL", (pn_tbl["Glucose (%)"] * 3.4 + pn_tbl["AA (%)"] * 4) / 100)
        pn_tbl["Osmolarity"] = (pn_tbl["Glucose (%)"] * 50 + pn_tbl["AA (%)"] * 100
                                + pn_tbl["Na"] * 2 + pn_tbl["K"] * 2 + pn_tbl["Mg"]
                                + pn_tbl["Ca"] * 1.4)
        st.dataframe(pn_tbl.style.format(precision=2), hide_index=True,
                     use_container_width=True)
        st.caption("電解質單位 mEq/L。本院 PN 固定添加:Zn 1.1 mg、Vit A 200 IU、"
                   "Vit D2 200 IU、Vit E 1 mg、Vit C 100 mg (每天)")
        st.caption("Osmolarity 係數:Glucose 50 / %、AA 100 / %、Na 2、K 2、Mg 1、Ca 1.4 "
                   "(per mEq);Lipid 1.7 / g (IF 20% ≈ 340 mOsm/L)")

        st.subheader("IVF 種類")
        ivf_df = pd.DataFrame(
            [[n, *v] for n, v in IVF_TABLE.items() if v],
            columns=["IVF", "Na (mEq/L)", "K (mEq/L)", "Glucose (%)"])
        st.dataframe(ivf_df, hide_index=True, use_container_width=True)
        st.caption("KCl: 2 mEq/mL;每 100 mL IVF 加 x mL KCl,K 增加 20x mEq/L。"
                   "以 D50W 提高濃度:(50 × D50W mL + 原濃度 × 原 mL) / 總 mL")

        st.subheader("奶水成分 (每公升;添加劑每 100 mL)")
        milk_ref = pd.DataFrame({**MILK_TABLE, **{k: v for k, v in FORTIFIER_TABLE.items()
                                                  if k != "無"}},
                                index=["Calorie (kcal)"] + NUTRIENTS).T
        st.dataframe(milk_ref.style.format(precision=1), use_container_width=True)

        st.subheader("Ca 補充")
        st.caption("10% Calcium gluconate = Ca 0.45 mEq/mL = Ca 9 mg/mL;"
                   "稀釋至 50 mg/mL,120–240 mg/kg 輸注 1 小時以上 "
                   "(Avery's Diseases of the Newborn 8th ed, Ch. 69)")

    st.caption("※ 本工具僅供輔助計算,實際處方請由臨床醫師/營養師確認。")
    st.caption("Neil JJC HMMCH 製作")


if __name__ == "__main__":
    main()
