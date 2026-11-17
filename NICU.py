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
登入 (選用): 帳號 PICU (床位 PI-01~PI-26) / NBC (床位 NBC-01~NBC-27),
  兩單位皆無 04、13、14、24 床;
  在標題列「登入 / 註冊」登入後,可點選床位、自動讀取該床最近一次紀錄、儲存與查看歷次紀錄。
  各單位可新增自訂床位 (只有自訂床位可刪除,刪除時連同紀錄);
  自訂床位存在 SQLite 的 custom_beds 表,或 Google Sheets 的 beds 工作表。
  密碼可在 .streamlit/secrets.toml 以 [passwords] PICU = "..." 覆寫。
個人帳號 (移植自 TPN 計算器;與單位帳號共用同一個登入對話框):
  - 標題旁的按鈕: 登入 / 註冊 (帳號、密碼皆需 4 個字以上,密碼以 PBKDF2 雜湊儲存)
  - 登入後可選擇病患、儲存、看歷史紀錄 (載入 / 刪除)、新增 / 管理 / 刪除病患
儲存位置 (單位帳號與個人帳號都一樣):
  - secrets 有 Google 試算表設定時,全部存到 Google Sheets (需 pip install gspread):
        [gsheets]
        spreadsheet = "試算表網址或 ID"
        gcp_json = '''
        { ...整個服務帳號金鑰 .json 檔的內容原樣貼上... }
        '''
    (舊寫法 [gsheet] url / key + [gcp_service_account] 也可以)
    工作表: records、beds (單位帳號);acct_users、acct_patients、acct_records (個人帳號)
    ※ 請用一份「NICU 專用」的試算表,不要和 TPN 計算器共用 (兩者都有 records 工作表)
  - 沒有設定時存本機 SQLite: 單位帳號 nicu_records.db (環境變數 NICU_DB 可改路徑)、
    個人帳號 nicu_accounts.db (環境變數 NICU_ACCT_DB_DIR 可指定資料夾)
  - 需要 Streamlit >= 1.40 (st.dialog、具 key 的 container)
執行: streamlit run streamlit_app.py
"""
import base64
import functools
import hashlib
import hmac
import html
import json
import os
import re
import secrets
import sqlite3
import tempfile
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
SKIP_BED_NUMS = {4, 13, 14, 24}  # 不使用的床號
UNITS = {
    "PICU": [f"PI-{i:02d}" for i in range(1, 27) if i not in SKIP_BED_NUMS],
    "NBC": [f"NBC-{i:02d}" for i in range(1, 28) if i not in SKIP_BED_NUMS],
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
    """secrets 有 Google 試算表設定就用 Google Sheets。支援兩種寫法:
    [gsheets] spreadsheet + gcp_json (同 TPN 計算器),或舊的 [gsheet] url/key + [gcp_service_account]。"""
    try:
        return "gsheets" in st.secrets or (
            "gcp_service_account" in st.secrets and "gsheet" in st.secrets)
    except Exception:
        return False


def storage_name() -> str:
    return "Google Sheets" if use_gsheets() else "本機 SQLite"


@st.cache_resource(show_spinner=False)
def _gs_book():
    """連線到 Google 試算表 (整個 app 共用一個連線)。"""
    import gspread  # 只有使用 Google Sheets 時才需要安裝
    try:
        if "gsheets" in st.secrets:
            cfg = st.secrets["gsheets"]
            ref = str(cfg["spreadsheet"]).strip()
            creds = (json.loads(cfg["gcp_json"]) if "gcp_json" in cfg
                     else dict(st.secrets["gcp_service_account"]))
        else:
            cfg = st.secrets["gsheet"]
            ref = str(cfg["url"] if "url" in cfg else cfg["key"]).strip()
            creds = dict(st.secrets["gcp_service_account"])
    except Exception as e:
        raise RuntimeError("Google 試算表的 secrets 設定不完整或格式錯誤 "
                           f"({type(e).__name__}:{e})") from e
    gc = gspread.service_account_from_dict(creds)
    return gc.open_by_url(ref) if ref.startswith("http") else gc.open_by_key(ref)


@st.cache_resource(show_spinner=False)
def _gsheet_ws():
    import gspread  # 只有使用 Google Sheets 時才需要安裝
    sh = _gs_book()
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


# --- 自訂床位 (各單位可新增;只有自訂床位可以刪除,預設床位固定) ---
GSHEET_BEDS_TAB = "beds"
GSHEET_BEDS_HEADER = ["unit", "bed", "created_at"]
BED_MAX_LEN = 20


@st.cache_resource(show_spinner=False)
def _gsheet_beds_ws():
    import gspread
    sh = _gs_book()
    try:
        ws = sh.worksheet(GSHEET_BEDS_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=GSHEET_BEDS_TAB, rows=200, cols=len(GSHEET_BEDS_HEADER))
        ws.append_row(GSHEET_BEDS_HEADER)
    return ws


@st.cache_data(ttl=GSHEET_CACHE_TTL, show_spinner=False)
def _gsheet_beds() -> list:
    """[(sheet 列號, unit, bed), ...]"""
    values = _gsheet_beds_ws().get_all_values()
    return [(i, v[0], v[1]) for i, v in enumerate(values[1:], start=2)
            if len(v) >= 2 and v[0] and v[1]]


def _beds_db() -> sqlite3.Connection:
    con = _db()
    con.execute("""CREATE TABLE IF NOT EXISTS custom_beds (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        unit TEXT NOT NULL, bed TEXT NOT NULL, created_at TEXT NOT NULL,
        UNIQUE (unit, bed))""")
    return con


def custom_beds(unit: str) -> list:
    """該單位自訂床位 (依新增順序)。"""
    if use_gsheets():
        return [b for _, u, b in _gsheet_beds() if u == unit]
    with closing(_beds_db()) as con:
        return [r[0] for r in con.execute(
            "SELECT bed FROM custom_beds WHERE unit=? ORDER BY id", (unit,))]


def all_beds(unit: str) -> list:
    return UNITS[unit] + custom_beds(unit)


def add_custom_bed(unit: str, bed: str):
    """成功回傳 None,失敗回傳錯誤訊息。"""
    bed = bed.strip()
    if not bed:
        return "請輸入床位名稱"
    if len(bed) > BED_MAX_LEN:
        return f"床位名稱最多 {BED_MAX_LEN} 個字"
    if bed.upper() in (b.upper() for b in all_beds(unit)):
        return f"已經有床位「{bed}」"
    now = datetime.now(TW_TZ).strftime("%Y-%m-%d %H:%M")
    if use_gsheets():
        _gsheet_beds_ws().append_row([unit, bed, now], value_input_option="RAW")
        _gsheet_beds.clear()
    else:
        with closing(_beds_db()) as con, con:
            con.execute("INSERT INTO custom_beds (unit, bed, created_at) VALUES (?, ?, ?)",
                        (unit, bed, now))
    return None


def delete_custom_bed(unit: str, bed: str) -> None:
    """刪除自訂床位及其所有紀錄;預設床位不會被刪除。"""
    if bed in UNITS[unit]:
        return
    clear_bed(unit, bed)
    if use_gsheets():
        _gsheet_beds.clear()
        targets = [r for r, u, b in _gsheet_beds() if u == unit and b == bed]
        ws = _gsheet_beds_ws()
        for r in sorted(targets, reverse=True):
            ws.delete_rows(r)
        _gsheet_beds.clear()
    else:
        with closing(_beds_db()) as con, con:
            con.execute("DELETE FROM custom_beds WHERE unit=? AND bed=?", (unit, bed))


def delete_unit_record(unit: str, rec_id: int):
    """刪除單位帳號的單筆紀錄。"""
    if use_gsheets():
        _gsheet_rows.clear()
        for r, i, u, _, _, _ in _gsheet_rows():
            if i == rec_id and u == unit:
                _gsheet_ws().delete_rows(r)
                break
        _gsheet_rows.clear()
    else:
        with closing(_db()) as con, con:
            con.execute("DELETE FROM records WHERE id=? AND unit=?", (rec_id, unit))


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


# ======================================================================
# 個人帳號 (移植自 TPN 計算器): 註冊 / 登入、管理病患、儲存、歷史紀錄
# 與上方 PICU / NBC 單位登入、床位紀錄互相獨立 (那部分完全沒有更動)
# 資料存在獨立的 SQLite 檔 nicu_accounts.db (環境變數 NICU_ACCT_DB_DIR 可指定資料夾)
# ======================================================================
ss = st.session_state
ACCT_DB_FILENAME = "nicu_accounts.db"
MIN_LEN = 4  # 帳號、密碼最少字數
PBKDF2_ROUNDS = 200_000


HISTORY_UNITS_NOTE = (
    "**單位**:BW g ｜ Feeding 每餐 mL、q 小時 ｜ PN Glucose % / AA % 速率 mL/hr ｜ "
    "IVF 速率 mL/hr ｜ TF mL/kg/day ｜ Cal kcal/kg/day (吃 + 打) ｜ SIR mg/kg/min ｜ "
    "AA PN 胺基酸 g/kg/day (占 TC %) ｜ Fat IF 20% 脂肪 g/kg/day (占 TC %)  \n"
    "TC (總熱量) 只計 PN + IVF + Fat,不含奶水。"
)


def _history_cols(d: dict) -> dict:
    """歷史紀錄的資料欄 (順序: Feeding、PN、IVF、TF、Cal、SIR、AA、Fat)。
    數值取自已存的「今日結果」字串,舊紀錄也能顯示:
      AA「PN AA/總蛋白」、PN「速率(Glu/AA/AA%TC)」、Fat「mL(g/kg/Fat%TC)」。"""
    sm = d.get("summary", {})

    def num_at(text: str, pattern: str):
        m = re.search(pattern, text or "")
        try:
            return float(m.group(1)) if m else None
        except ValueError:
            return None

    def gkg_pct(gkg, pct) -> str:
        if gkg is None:
            return "—"
        return f"{gkg:.2f} ({pct:.1f}%)" if pct is not None else f"{gkg:.2f}"

    has_pn = bool(d.get("pn_rate"))
    has_fat = bool(d.get("fat_ml"))
    aa_gkg = num_at(sm.get("AA"), r"^\s*([^/]+)/")             # PN AA g/kg/day
    aa_pct = num_at(sm.get("PN"), r"/([^/()]+)\)\s*$")          # AA 占 TC %
    fat_gkg = num_at(sm.get("Fat"), r"\(([^/()]+)/")            # Fat g/kg/day
    fat_pct = num_at(sm.get("Fat"), r"/([^/()]+)\)\s*$")        # Fat 占 TC %
    return {
        "Feeding": sm.get("Feeding", "—"),
        "PN": (f"{fmt(d.get('pn_glu', 0))}% / {fmt(d.get('pn_aa', 0))}% "
               f"{fmt(d.get('pn_rate', 0))}") if has_pn else "—",
        "IVF": sm.get("IVF", "—"),
        "TF": sm.get("TF", "—"),
        "Cal": sm.get("Cal", "—"),
        "SIR": sm.get("SIR", "—"),
        "AA": gkg_pct(aa_gkg, aa_pct) if has_pn else "—",
        "Fat": gkg_pct(fat_gkg, fat_pct) if has_fat else "—",
    }


def _pick_acct_db_path() -> Path:
    """找一個可以正常寫入的位置放資料庫 (同 TPN 計算器)。
    Streamlit Cloud 的程式資料夾可能無法寫入 SQLite,所以優先放家目錄或暫存資料夾。"""
    candidates = []
    if os.environ.get("NICU_ACCT_DB_DIR"):
        candidates.append(Path(os.environ["NICU_ACCT_DB_DIR"]))
    candidates += [Path.home() / ".nicu_calc", Path(tempfile.gettempdir()) / "nicu_calc",
                   Path(__file__).parent]
    for folder in candidates:
        path = folder / ACCT_DB_FILENAME
        try:
            folder.mkdir(parents=True, exist_ok=True)
            with closing(sqlite3.connect(path)) as c:
                c.execute("BEGIN IMMEDIATE")
                c.execute("CREATE TABLE IF NOT EXISTS _write_test (x TEXT)")
                c.execute("INSERT INTO _write_test VALUES ('ok')")
                c.rollback()
            return path
        except (sqlite3.Error, OSError):
            continue
    raise RuntimeError("找不到可寫入的位置來存放帳號資料庫")


@functools.lru_cache(maxsize=1)
def acct_db_path() -> Path:
    """只有使用 SQLite 時才需要決定資料庫位置。"""
    path = _pick_acct_db_path()
    acct_init_db(path)
    return path


def acct_storage_name() -> str:
    return "Google Sheets" if use_gsheets() else str(acct_db_path())


def _acct_db() -> sqlite3.Connection:
    conn = sqlite3.connect(acct_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def acct_init_db(path: Path) -> None:
    with closing(sqlite3.connect(path)) as c, c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                username   TEXT NOT NULL UNIQUE COLLATE NOCASE,
                pw_hash    TEXT NOT NULL,
                salt       TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS patients (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                name       TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE (user_id, name)
            );
            CREATE TABLE IF NOT EXISTS records (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                patient_id INTEGER NOT NULL REFERENCES patients(id) ON DELETE CASCADE,
                created_at TEXT NOT NULL,
                note       TEXT,
                data       TEXT NOT NULL
            );
            """
        )


# --- 個人帳號的 Google Sheets 儲存 ---
# 每個表是一個工作表,第 1 列為欄名;寫入一律用 RAW (文字不會被當成公式或數字)。
ACCT_TABLES = {
    "acct_users": ["id", "username", "pw_hash", "salt", "created_at"],
    "acct_patients": ["id", "user_id", "name", "created_at"],
    "acct_records": ["id", "patient_id", "created_at", "note", "data"],
}


@st.cache_resource(show_spinner=False)
def _acct_sheets() -> dict:
    """確認個人帳號的工作表存在 (不存在就建立),回傳 {表名: worksheet}。"""
    import gspread
    book = _gs_book()
    wss = {}
    for name, cols in ACCT_TABLES.items():
        try:
            ws = book.worksheet(name)
        except gspread.exceptions.WorksheetNotFound:
            ws = book.add_worksheet(title=name, rows=1000, cols=len(cols))
        if ws.row_values(1) != cols:
            ws.update(range_name="A1", values=[cols], value_input_option="RAW")
        wss[name] = ws
    return wss


def _acct_rows(name: str, values: list) -> list:
    cols = ACCT_TABLES[name]
    out = []
    for row in values[1:]:
        if any(row):
            row = list(row) + [""] * (len(cols) - len(row))
            out.append(dict(zip(cols, row)))
    return out


def _acct_read_fresh(name: str) -> list:
    """直接讀取 (不經暫存),回傳 [{欄名: 值(字串)}, ...],依寫入順序。"""
    return _acct_rows(name, _acct_sheets()[name].get_all_values())


@st.cache_data(ttl=GSHEET_CACHE_TTL, show_spinner=False)
def _acct_read(name: str) -> list:
    """讀取並暫存;本程式寫入後會立即清除暫存。"""
    return _acct_read_fresh(name)


def _acct_insert(name: str, row: dict) -> None:
    _acct_sheets()[name].append_row(
        [str(row.get(c, "")) for c in ACCT_TABLES[name]], value_input_option="RAW",
        insert_data_option="INSERT_ROWS", table_range="A1")
    _acct_read.clear()


def _acct_delete_where(name: str, pred) -> int:
    """刪除符合條件的所有列 (一次 API 呼叫),回傳刪除筆數。"""
    ws = _acct_sheets()[name]
    cols = ACCT_TABLES[name]
    row_nums = []
    for i, row in enumerate(ws.get_all_values()[1:], start=2):
        row = list(row) + [""] * (len(cols) - len(row))
        if any(row) and pred(dict(zip(cols, row))):
            row_nums.append(i)
    if row_nums:  # 由下往上刪,前面的列號才不會位移
        ws.spreadsheet.batch_update({"requests": [
            {"deleteDimension": {"range": {"sheetId": ws.id, "dimension": "ROWS",
                                           "startIndex": r - 1, "endIndex": r}}}
            for r in sorted(row_nums, reverse=True)]})
    _acct_read.clear()
    return len(row_nums)


def _acct_new_id(prefix: str) -> str:
    return f"{prefix}{secrets.token_hex(6)}"  # 英文字母開頭,避免被當成數字


def _acct_owned(user_id) -> set:
    return {p["id"] for p in _acct_read("acct_patients") if p["user_id"] == user_id}


def _now_str() -> str:
    return datetime.now(TW_TZ).strftime("%Y-%m-%d %H:%M:%S")


def _hash_pw(password: str, salt_hex: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt_hex),
                               PBKDF2_ROUNDS).hex()


def acct_create_user(username: str, password: str):
    """成功回傳 (user_id, None);失敗回傳 (None, 錯誤訊息)。"""
    username = username.strip()
    if len(username) < MIN_LEN:
        return None, f"帳號至少需要 {MIN_LEN} 個字"
    if len(password) < MIN_LEN:
        return None, f"密碼至少需要 {MIN_LEN} 個字"
    salt = secrets.token_hex(16)
    if use_gsheets():
        if any(u["username"].lower() == username.lower()
               for u in _acct_read_fresh("acct_users")):
            return None, "此帳號已被使用"
        uid = _acct_new_id("u")
        _acct_insert("acct_users", {"id": uid, "username": username,
                                    "pw_hash": _hash_pw(password, salt), "salt": salt,
                                    "created_at": _now_str()})
        return uid, None
    try:
        with closing(_acct_db()) as c, c:
            cur = c.execute(
                "INSERT INTO users (username, pw_hash, salt, created_at) VALUES (?, ?, ?, ?)",
                (username, _hash_pw(password, salt), salt, _now_str()),
            )
            return cur.lastrowid, None
    except sqlite3.IntegrityError:
        return None, "此帳號已被使用"


def acct_verify_user(username: str, password: str):
    if use_gsheets():
        name = username.strip().lower()
        row = next((u for u in _acct_read("acct_users") if u["username"].lower() == name),
                   None)
    else:
        with closing(_acct_db()) as c:
            row = c.execute("SELECT * FROM users WHERE username = ?",
                            (username.strip(),)).fetchone()
    if row and hmac.compare_digest(row["pw_hash"], _hash_pw(password, row["salt"])):
        return {"id": row["id"], "username": row["username"]}
    return None


def acct_list_patients(user_id) -> list:
    if use_gsheets():
        counts = {}
        for r in _acct_read("acct_records"):
            counts[r["patient_id"]] = counts.get(r["patient_id"], 0) + 1
        return [{"id": p["id"], "name": p["name"], "n": counts.get(p["id"], 0)}
                for p in _acct_read("acct_patients") if p["user_id"] == user_id]
    with closing(_acct_db()) as c:
        return c.execute(
            """SELECT p.id, p.name, COUNT(r.id) AS n FROM patients p
               LEFT JOIN records r ON r.patient_id = p.id
               WHERE p.user_id = ? GROUP BY p.id ORDER BY p.created_at, p.id""",
            (user_id,),
        ).fetchall()


def acct_create_patient(user_id, name: str):
    if use_gsheets():
        if any(p["user_id"] == user_id and p["name"] == name
               for p in _acct_read_fresh("acct_patients")):
            return None, f"已經有名為「{name}」的病患"
        pid = _acct_new_id("p")
        _acct_insert("acct_patients", {"id": pid, "user_id": user_id, "name": name,
                                       "created_at": _now_str()})
        return pid, None
    try:
        with closing(_acct_db()) as c, c:
            cur = c.execute("INSERT INTO patients (user_id, name, created_at) VALUES (?, ?, ?)",
                            (user_id, name, _now_str()))
            return cur.lastrowid, None
    except sqlite3.IntegrityError:
        return None, f"已經有名為「{name}」的病患"


def acct_delete_patient(user_id, patient_id) -> None:
    if use_gsheets():
        if patient_id not in _acct_owned(user_id):
            return
        _acct_delete_where("acct_records", lambda r: r["patient_id"] == patient_id)
        _acct_delete_where("acct_patients",
                           lambda p: p["id"] == patient_id and p["user_id"] == user_id)
        return
    with closing(_acct_db()) as c, c:  # 紀錄透過 ON DELETE CASCADE 一併刪除
        c.execute("DELETE FROM patients WHERE id = ? AND user_id = ?", (patient_id, user_id))


def acct_add_record(user_id, patient_id, note: str, data: dict) -> bool:
    if use_gsheets():
        if patient_id not in _acct_owned(user_id):
            return False
        _acct_insert("acct_records", {"id": _acct_new_id("r"), "patient_id": patient_id,
                                      "created_at": _now_str(), "note": note,
                                      "data": json.dumps(data, ensure_ascii=False)})
        return True
    with closing(_acct_db()) as c, c:
        owner = c.execute("SELECT 1 FROM patients WHERE id = ? AND user_id = ?",
                          (patient_id, user_id)).fetchone()
        if not owner:
            return False
        c.execute("INSERT INTO records (patient_id, created_at, note, data) VALUES (?, ?, ?, ?)",
                  (patient_id, _now_str(), note, json.dumps(data, ensure_ascii=False)))
        return True


def acct_list_records(user_id, patient_id) -> list:
    if use_gsheets():
        if patient_id not in _acct_owned(user_id):
            return []
        rows = [r for r in _acct_read("acct_records") if r["patient_id"] == patient_id]
        # 由新到舊;同時間者後寫入的排前面
        return [r for _, r in sorted(enumerate(rows), key=lambda t: (t[1]["created_at"], t[0]),
                                     reverse=True)]
    with closing(_acct_db()) as c:
        return c.execute(
            """SELECT r.id, r.created_at, r.note, r.data FROM records r
               JOIN patients p ON p.id = r.patient_id
               WHERE r.patient_id = ? AND p.user_id = ?
               ORDER BY r.created_at DESC, r.id DESC""",
            (patient_id, user_id),
        ).fetchall()


def acct_delete_record(user_id, record_id) -> None:
    if use_gsheets():
        owned = _acct_owned(user_id)
        _acct_delete_where("acct_records",
                           lambda r: r["id"] == record_id and r["patient_id"] in owned)
        return
    with closing(_acct_db()) as c, c:
        c.execute(
            """DELETE FROM records WHERE id = ? AND patient_id IN
               (SELECT id FROM patients WHERE user_id = ?)""",
            (record_id, user_id),
        )

# --- 狀態輔助 ---
def acct_flash(kind: str, msg: str, seconds: float = 0) -> None:
    """記下訊息,下次整頁更新時顯示在標題列。"""
    ss.setdefault("acct_flash", []).append((kind, msg, seconds))


def acct_apply(data: dict) -> None:
    """排定要載入的輸入值;真正寫入是在下一次整頁執行的最前面 (任何輸入欄位畫出來之前),
    這樣從對話框按的也能正確更新所有欄位。空 dict = 重設為預設值。"""
    ss._acct_pending = dict(data)


def acct_load_latest(pid: int) -> None:
    recs = acct_list_records(ss.user["id"], pid)
    if recs:
        acct_apply(json.loads(recs[0]["data"]))
        acct_flash("info", f"已載入最近紀錄 ({recs[0]['created_at'][:16]})")
    else:
        acct_apply({})


def request_app_rerun() -> None:
    """對話框裡的 callback 用:之後整頁重新整理並關閉對話框。"""
    ss._acct_rerun = True


def rerun_if_requested() -> None:
    if ss.pop("_acct_rerun", False):
        st.rerun()


def _safe_storage_label() -> str:
    try:
        return storage_name() if ss.get("unit") else acct_storage_name()
    except Exception:
        return storage_name()


def db_safe(func):
    """資料庫出錯時不讓整個 app 當掉,改把錯誤訊息顯示在畫面上。"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:  # SQLite / Google Sheets 等錯誤
            ss._acct_db_err = (f"資料存取錯誤 ({type(e).__name__}):{e}\n\n"
                               f"儲存位置:{_safe_storage_label()}"
                               + ("\n\nGoogle Sheets:請確認試算表已「共用」給服務帳號 (編輯者)"
                                  if use_gsheets() else ""))
    return wrapper


def show_db_error() -> None:
    err = ss.pop("_acct_db_err", None)
    if err:
        st.error(err)


# --- 登入狀態 ---
# 單位帳號 (PICU / NBC): ss.unit + ss.bed_sel (床位);個人帳號: ss.user + ss.patient_id (病患)
def _clear_login() -> None:
    for k in ["unit", "user", "patient_id", "confirm_del_patient", "confirm_del_record",
              "confirm_del_bed"]:
        ss.pop(k, None)


def _unit_load_latest() -> None:
    """載入目前床位最近一筆紀錄;沒有紀錄就回到預設值。"""
    unit, bed = ss.unit, ss.bed_sel
    recs = [(t, d) for _, b, t, d in _unit_rows(unit) if b == bed]  # 由舊到新
    if recs:
        acct_apply(json.loads(recs[-1][1]))
        acct_flash("info", f"{bed}:已載入 {recs[-1][0]} 的紀錄")
    else:
        acct_apply({})
        acct_flash("info", f"{bed}:尚無紀錄,已重設為預設值")


def target_records(target: tuple) -> list:
    """[(id, 時間, 備註, data), ...] 由新到舊。
    target: ("unit", 單位, 床位) 或 ("acct", 病患 id)。"""
    if target[0] == "unit":
        _, unit, bed = target
        out = []
        for i, b, t, d in reversed(_unit_rows(unit)):
            if b == bed:
                dd = json.loads(d)
                out.append((i, t[:16], dd.get("note", ""), dd))
        return out
    return [(r["id"], r["created_at"][:16], r["note"] or "", json.loads(r["data"]))
            for r in acct_list_records(ss.user["id"], target[1])]


# --- callbacks: 登入 / 註冊 / 登出 ---
PERSONAL = "個人帳號"
LOGIN_KINDS = list(UNITS) + [PERSONAL]  # 登入選單: PICU / NBC / 個人帳號
LOGIN_KIND_LABELS = {"PICU": "PICU (單位帳號)", "NBC": "NBC (單位帳號)",
                     PERSONAL: "個人帳號 (自行管理病患)"}


@db_safe
def cb_login() -> None:
    kind = ss.get("acct_li_kind", LOGIN_KINDS[0])
    name = ss.get("acct_li_user", "").strip() if kind == PERSONAL else kind
    pw = ss.get("acct_li_pw", "")
    ss.acct_li_pw = ""
    if name.upper() in UNITS:  # 單位帳號 PICU / NBC (選單選擇,或在個人帳號欄輸入單位名稱)
        unit = name.upper()
        if not check_login(unit, pw):
            ss._acct_auth_err = "帳號或密碼錯誤"
            return
        _clear_login()
        ss.unit = unit
        ss.bed_sel = UNITS[unit][0]
        _unit_load_latest()
        acct_flash("success", f"已登入單位 {unit}")
        request_app_rerun()
        return
    user = acct_verify_user(name, pw)  # 個人帳號
    if not user:
        ss._acct_auth_err = "帳號或密碼錯誤"
        return
    _clear_login()
    ss.user = user
    plist = acct_list_patients(user["id"])
    if plist:
        ss.patient_id = plist[0]["id"]
        acct_load_latest(plist[0]["id"])
    acct_flash("success", f"歡迎,{user['username']}")
    request_app_rerun()


@db_safe
def cb_register() -> None:
    u, p, p2 = ss.get("acct_rg_user", ""), ss.get("acct_rg_pw", ""), ss.get("acct_rg_pw2", "")
    if u.strip().upper() in UNITS:
        ss._acct_auth_err = f"「{u.strip()}」是單位帳號,請直接在「登入」頁以單位密碼登入"
        return
    if p != p2:
        ss._acct_auth_err = "兩次輸入的密碼不一致"
        return
    uid, err = acct_create_user(u, p)
    if err:
        ss._acct_auth_err = err
        return
    _clear_login()
    ss.user = {"id": uid, "username": u.strip()}
    ss.acct_rg_pw = ss.acct_rg_pw2 = ""
    acct_flash("success", "帳號已建立,請新增第一位病患")
    request_app_rerun()


def cb_logout() -> None:
    _clear_login()
    acct_apply({})  # 登出後清空畫面上的病人資料
    acct_flash("info", "已登出")


# --- callbacks: 床位 / 病患 ---
@db_safe
def cb_select_bed() -> None:
    _unit_load_latest()


@db_safe
def cb_clear_bed(unit: str, bed: str) -> None:
    """刪除該床所有紀錄,並把畫面輸入重設為預設值。"""
    n = len(target_records(("unit", unit, bed)))
    if n:
        clear_bed(unit, bed)
    acct_apply({})
    acct_flash("success", f"{bed} 已清空" + (f" (刪除 {n} 筆紀錄)" if n else " (畫面已重設)"))
    request_app_rerun()


@db_safe
def cb_add_bed() -> None:
    unit = ss.unit
    bed = ss.get("acct_new_bed", "").strip()
    err = add_custom_bed(unit, bed)
    if err:
        ss._acct_bed_err = err
        return
    ss.acct_new_bed = ""
    ss.bed_sel = bed
    acct_apply({})  # 新床位從預設值開始
    acct_flash("success", f"已新增床位 {bed}")
    request_app_rerun()


@db_safe
def cb_delete_bed(bed: str) -> None:
    unit = ss.unit
    n = len(target_records(("unit", unit, bed)))
    delete_custom_bed(unit, bed)
    ss.pop("confirm_del_bed", None)
    if ss.get("bed_sel") == bed:
        ss.bed_sel = UNITS[unit][0]
        _unit_load_latest()
    acct_flash("success", f"已刪除床位 {bed}" + (f" 及 {n} 筆紀錄" if n else ""))
    request_app_rerun()


@db_safe
def cb_select_patient() -> None:
    acct_load_latest(ss.patient_id)


@db_safe
def cb_add_patient() -> None:
    name = ss.get("acct_new_patient", "").strip()
    if not name:
        ss._acct_pat_err = "請輸入病患名稱"
        return
    pid, err = acct_create_patient(ss.user["id"], name)
    if err:
        ss._acct_pat_err = err
        return
    ss.patient_id = pid
    ss.acct_new_patient = ""
    acct_apply({})  # 新病患從預設值開始
    acct_flash("success", f"已新增病患「{name}」")
    request_app_rerun()


@db_safe
def cb_choose_patient(pid: int) -> None:
    ss.patient_id = pid
    acct_load_latest(pid)
    request_app_rerun()


@db_safe
def cb_delete_patient(pid: int, name: str) -> None:
    acct_delete_patient(ss.user["id"], pid)
    ss.pop("confirm_del_patient", None)
    if ss.get("patient_id") == pid:
        ss.pop("patient_id", None)
    acct_flash("success", f"已刪除病患「{name}」及其紀錄")
    request_app_rerun()


# --- callbacks: 紀錄 ---
@db_safe
def cb_save_record(target: tuple, payload: dict) -> None:
    note = ss.get("acct_note", "").strip()
    if target[0] == "unit":
        _, unit, bed = target
        t = save_record(unit, bed, {**payload, "note": note})
        ok, msg = True, f"已儲存 {bed} ({t})"
    else:
        ok, msg = acct_add_record(ss.user["id"], target[1], note, payload), "已儲存紀錄"
    if ok:
        ss.acct_note = ""
        acct_flash("success", msg, seconds=1.5)
    else:
        acct_flash("error", "儲存失敗:找不到此病患")
    request_app_rerun()


@db_safe
def cb_load_record(data: dict, created_at: str) -> None:
    acct_apply(data)
    acct_flash("info", f"已載入 {created_at} 的紀錄")
    request_app_rerun()


@db_safe
def cb_delete_record(target: tuple, rid: int) -> None:
    if target[0] == "unit":
        delete_unit_record(target[1], rid)
    else:
        acct_delete_record(ss.user["id"], rid)
    ss.pop("confirm_del_record", None)  # 對話框保持開啟,可以繼續刪


def cb_set(key: str, value) -> None:
    ss[key] = value


# --- 對話框 ---
@st.dialog("登入 / 註冊")
def auth_dialog() -> None:
    rerun_if_requested()
    show_db_error()
    err = ss.pop("_acct_auth_err", None)
    if err:
        st.error(err)
    tab_login, tab_reg = st.tabs(["登入", "建立帳號"])
    with tab_login:
        # 選單放在 form 外面,切換時才能立即顯示 / 隱藏帳號欄位
        kind = st.selectbox("登入類型", LOGIN_KINDS, key="acct_li_kind",
                            format_func=lambda k: LOGIN_KIND_LABELS.get(k, k))
        with st.form("acct_login_form", border=False):
            if kind == PERSONAL:
                st.text_input("帳號", key="acct_li_user")
            st.text_input("單位密碼" if kind in UNITS else "密碼", type="password",
                          key="acct_li_pw")
            st.form_submit_button("登入", type="primary", use_container_width=True,
                                  on_click=cb_login)
        if kind in UNITS:
            beds = UNITS[kind]
            st.caption(f"床位 {beds[0]} – {beds[-1]} 及自訂床位,依床位儲存紀錄。")
        else:
            st.caption("個人帳號請先到「建立帳號」註冊,可自行管理病患。")
    with tab_reg:
        with st.form("acct_register_form", border=False):
            st.text_input("帳號", key="acct_rg_user", help=f"至少 {MIN_LEN} 個字")
            st.text_input("密碼", type="password", key="acct_rg_pw", help=f"至少 {MIN_LEN} 個字")
            st.text_input("再次輸入密碼", type="password", key="acct_rg_pw2")
            st.form_submit_button("建立帳號", type="primary", use_container_width=True,
                                  on_click=cb_register)
        st.caption("單位帳號 (PICU / NBC) 不需註冊。")


@st.dialog("管理病患")
def patients_dialog() -> None:
    rerun_if_requested()
    show_db_error()
    with st.form("acct_new_patient_form", border=False):
        c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
        c1.text_input("新增病患", key="acct_new_patient", max_chars=50,
                      placeholder="例: 床號 / 代號 (請勿輸入姓名、病歷號)")
        c2.form_submit_button("新增", type="primary", use_container_width=True,
                              on_click=cb_add_patient)
    err = ss.pop("_acct_pat_err", None)
    if err:
        st.error(err)

    st.divider()
    plist = acct_list_patients(ss.user["id"])
    if not plist:
        st.caption("尚無病患")
    for p in plist:
        with st.container(border=True):
            c1, c2, c3 = st.columns([3, 1.3, 1.3], vertical_alignment="center")
            current = "　✅ 目前" if p["id"] == ss.get("patient_id") else ""
            c1.markdown(f"**{p['name']}**{current}  \n{p['n']} 筆紀錄")
            if ss.get("confirm_del_patient") == p["id"]:
                c2.button("確認刪除", type="primary", key=f"pdel_ok_{p['id']}",
                          use_container_width=True,
                          on_click=cb_delete_patient, args=(p["id"], p["name"]))
                c3.button("取消", key=f"pdel_no_{p['id']}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_patient", None))
                st.warning(f"將刪除「{p['name']}」及其 {p['n']} 筆紀錄,無法復原。")
            else:
                c2.button("選擇", key=f"psel_{p['id']}", use_container_width=True,
                          on_click=cb_choose_patient, args=(p["id"],))
                c3.button("🗑 刪除", key=f"pdel_{p['id']}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_patient", p["id"]))


@st.dialog("新增 / 管理床位")
def beds_dialog(unit: str) -> None:
    rerun_if_requested()
    show_db_error()
    with st.form("acct_new_bed_form", border=False):
        c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
        c1.text_input(f"新增 {unit} 床位", key="acct_new_bed", max_chars=BED_MAX_LEN,
                      placeholder="例: 加床-1 (請勿輸入姓名、病歷號)")
        c2.form_submit_button("新增", type="primary", use_container_width=True,
                              on_click=cb_add_bed)
    err = ss.pop("_acct_bed_err", None)
    if err:
        st.error(err)
    fixed = UNITS[unit]
    st.caption(f"預設床位 {fixed[0]} – {fixed[-1]} (共 {len(fixed)} 床) 固定,不可刪除;"
               "只有自訂床位可以刪除。")

    st.divider()
    beds = custom_beds(unit)
    if not beds:
        st.caption("尚無自訂床位")
    for bed in beds:
        n = len(target_records(("unit", unit, bed)))
        with st.container(border=True):
            c1, c2, c3 = st.columns([3, 1.3, 1.3], vertical_alignment="center")
            current = "　✅ 目前" if bed == ss.get("bed_sel") else ""
            c1.markdown(f"**{bed}**{current}  \n{n} 筆紀錄")
            if ss.get("confirm_del_bed") == bed:
                c2.button("確認刪除", type="primary", key=f"bdel_ok_{bed}",
                          use_container_width=True, on_click=cb_delete_bed, args=(bed,))
                c3.button("取消", key=f"bdel_no_{bed}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_bed", None))
                st.warning(f"將刪除床位「{bed}」" + (f"及其 {n} 筆紀錄" if n else "")
                           + ",無法復原。")
            else:
                c3.button("🗑 刪除", key=f"bdel_{bed}", use_container_width=True,
                          on_click=cb_set, args=("confirm_del_bed", bed))


@st.dialog("清空床位")
def clear_bed_dialog(unit: str, bed: str) -> None:
    rerun_if_requested()
    show_db_error()
    n = len(target_records(("unit", unit, bed)))
    st.warning(f"確定要清空 {unit} · {bed}?"
               + (f"將刪除 {n} 筆紀錄,無法復原。" if n else "此床尚無紀錄,只會重設畫面。"))
    c1, c2 = st.columns(2)
    c1.button("確認清空", type="primary", use_container_width=True,
              on_click=cb_clear_bed, args=(unit, bed))
    c2.button("取消", use_container_width=True, on_click=request_app_rerun)


@st.dialog("儲存紀錄")
def save_dialog(target: tuple, label: str, payload: dict) -> None:
    rerun_if_requested()
    show_db_error()
    sm = payload.get("summary", {})
    st.markdown(f"{'床位' if target[0] == 'unit' else '病患'}:**{label}**")
    st.caption(
        f"BW {fmt(payload['bw_g'])} g ｜ TF {sm.get('TF', '—')} ｜ Cal {sm.get('Cal', '—')} ｜ "
        f"SIR {sm.get('SIR', '—')} ｜ AA {sm.get('AA', '—')} ｜ PN {sm.get('PN', '—')} ｜ "
        f"Fat {sm.get('Fat', '—')}"
    )
    st.caption(f"Feeding:{sm.get('Feeding', '—')} ｜ IVF:{sm.get('IVF', '—')}")
    st.text_input("備註 (選填)", key="acct_note", max_chars=200, placeholder="例: Day 3 增量")
    st.button("💾 儲存", type="primary", use_container_width=True,
              on_click=cb_save_record, args=(target, payload))


def _text_width(text) -> int:
    """估計文字寬度:中文等全形字算 2,英數算 1。"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def _cell(text, header: bool = False) -> str:
    weight = "font-weight:600;" if header else ""
    height = "2rem" if header else "2.5rem"
    return (f"<div style='display:flex; align-items:center; justify-content:center; "
            f"min-height:{height}; text-align:center; {weight} font-size:0.9rem; "
            f"line-height:1.3; word-break:break-word;'>{html.escape(str(text))}</div>")


@st.dialog("歷史紀錄", width="large")
def history_dialog(target: tuple, label: str) -> None:
    rerun_if_requested()
    show_db_error()
    st.markdown(f"{'床位' if target[0] == 'unit' else '病患'}:**{label}**")
    recs = target_records(target)
    if not recs:
        st.caption("尚無紀錄。調整好資料後,按上方「💾 儲存」即可存檔。")
        return

    rows = []
    for rid, when, note, d in recs:
        rows.append({
            "_id": rid,
            "_data": d,
            "時間": when,
            "備註": note or d.get("bed") or "—",
            "BW (g)": fmt(d.get("bw_g", 0)),
            **_history_cols(d),
        })
    cols = [c for c in rows[0] if not c.startswith("_")]
    # 欄寬 = 欄名與內容中最寬者 (上限 28,過長自動換行;按鈕欄固定寬度)
    widths = [5, 5] + [
        min(28, max(_text_width(c), *(_text_width(r[c]) for r in rows)) + 2) for c in cols
    ]

    st.markdown(
        """
        <style>
        [class*="st-key-hrow_"] {
            border-bottom: 1px solid rgba(128, 128, 128, 0.25);
            padding: 0.3rem 0;
        }
        .st-key-hrow_header { border-bottom-width: 2px; }
        [class*="st-key-hrow_"] [data-testid="stMarkdown"],
        [class*="st-key-hrow_"] [data-testid="stMarkdownContainer"],
        [class*="st-key-hrow_"] .stMarkdown { margin: 0 !important; }
        [class*="st-key-hrow_"] p { margin: 0 !important; }
        [class*="st-key-hrow_"] button { min-height: 2.5rem; padding: 0 0.25rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.container(key="hrow_header"):
        hdr = st.columns(widths, gap="small", vertical_alignment="center")
        for c, name in zip(hdr, ["載入", "刪除"] + cols):
            c.markdown(_cell(name, header=True), unsafe_allow_html=True)

    for r in rows:
        rid = r["_id"]
        with st.container(key=f"hrow_{rid}"):
            cells = st.columns(widths, gap="small", vertical_alignment="center")
            cells[0].button("📥", key=f"rload_{rid}", use_container_width=True,
                            help="載入這筆資料到計算器",
                            on_click=cb_load_record, args=(r["_data"], r["時間"]))
            cells[1].button("🗑", key=f"rdel_{rid}", use_container_width=True,
                            help="刪除這筆紀錄",
                            on_click=cb_set, args=("confirm_del_record", rid))
            for c, name in zip(cells[2:], cols):
                c.markdown(_cell(r[name]), unsafe_allow_html=True)

    st.caption(HISTORY_UNITS_NOTE)

    rid = ss.get("confirm_del_record")
    ids = [r["_id"] for r in rows]
    if rid in ids:
        when = rows[ids.index(rid)]["時間"]
        with st.container(border=True):
            st.warning(f"確定要刪除 {when} 的紀錄?刪除後無法復原。")
            c1, c2, _ = st.columns([1, 1, 3])
            c1.button("確認刪除", type="primary", key=f"rdel_ok_{rid}",
                      use_container_width=True, on_click=cb_delete_record, args=(target, rid))
            c2.button("取消", key=f"rdel_no_{rid}", use_container_width=True,
                      on_click=cb_set, args=("confirm_del_record", None))


# --- 標題列 ---
_ICONS = {"success": "✅", "error": "⚠️", "info": "ℹ️"}
NOTE_SECONDS = 3  # 一般提示顯示秒數 (有指定秒數者依指定)


def render_notes(notes: list) -> None:
    """在目前位置顯示提示文字 (靠右對齊),數秒後淡出。"""
    parts = []
    for kind, msg, sec in notes:
        ss._note_n = ss.get("_note_n", 0) + 1  # 每次用新的動畫名稱,確保動畫重新播放
        anim = f"nicuFade{ss._note_n}"
        fg, bg = ("#b42318", "#fde4e1") if kind == "error" else ("#177233", "#dff3e4")
        parts.append(
            f"<style>@keyframes {anim} {{ 0%, 85% {{ opacity: 1; }} "
            f"100% {{ opacity: 0; visibility: hidden; }} }}</style>"
            f"<span style='display:inline-block; padding:0.35rem 0.8rem; "
            f"border-radius:0.5rem; font-size:0.9rem; font-weight:600; "
            f"white-space:normal; text-align:right; "
            f"color:{fg}; background:{bg}; animation:{anim} {sec or NOTE_SECONDS}s forwards;'>"
            f"{_ICONS.get(kind, '')} {html.escape(msg)}</span>"
        )
    st.markdown(
        "<div style='position:relative; min-height:2.5rem;'>"
        "<div style='position:absolute; right:0; top:50%; transform:translateY(-50%); "
        "z-index:1000; width:max-content; max-width:min(32rem, 45vw); "
        "display:flex; flex-direction:column; align-items:flex-end; gap:0.25rem;'>"
        + "".join(parts) + "</div></div>",
        unsafe_allow_html=True,
    )


def render_account_header(title_html: str):
    """標題列 (固定在上方): 左邊標題,右邊登入 / 床位或病患 / 儲存 / 紀錄 / 帳號。
    回傳 (要開啟的對話框, {病患 id: 名稱}, 目前病患 id, 單位各床最近紀錄)。"""
    try:
        dark = st.context.theme.type == "dark"  # Streamlit >= 1.46
    except Exception:
        dark = False
    st.markdown(
        f"""
        <style>
        .st-key-hdr_ctrl [data-testid="stHorizontalBlock"] {{
            flex-wrap: nowrap !important; justify-content: flex-end;
            align-items: center; gap: 0.5rem;
        }}
        .st-key-hdr_ctrl [data-testid="stColumn"],
        .st-key-hdr_ctrl [data-testid="column"] {{
            flex: 0 0 auto !important; width: auto !important; min-width: 0 !important;
        }}
        .st-key-hdr_ctrl [data-testid="stColumn"]:has(.st-key-notice),
        .st-key-hdr_ctrl [data-testid="column"]:has(.st-key-notice) {{
            flex: 1 1 auto !important; min-width: 0 !important; overflow: visible !important;
        }}
        .st-key-notice, .st-key-notice [data-testid="stElementContainer"],
        .st-key-notice [data-testid="stMarkdown"],
        .st-key-notice [data-testid="stMarkdownContainer"],
        .st-key-hdr_ctrl, .st-key-hdr_ctrl [data-testid="stHorizontalBlock"] {{
            overflow: visible !important;
        }}
        /* 病患選單約 10rem;床位選單含體重與時間,較寬 */
        .st-key-hdr_ctrl [data-testid="stColumn"]:has(.st-key-pat_sel),
        .st-key-hdr_ctrl [data-testid="column"]:has(.st-key-pat_sel) {{
            flex: 0 0 10rem !important; width: 10rem !important;
        }}
        .st-key-hdr_ctrl [data-testid="stColumn"]:has(.st-key-bed_box),
        .st-key-hdr_ctrl [data-testid="column"]:has(.st-key-bed_box) {{
            flex: 0 0 16rem !important; width: 16rem !important;
        }}
        .st-key-hdr_ctrl button {{ white-space: nowrap; }}
        .st-key-app_header {{
            position: sticky; top: 0; z-index: 999;
            background-color: {"#0e1117" if dark else "#ffffff"};
            padding-bottom: 0.25rem;
            border-bottom: 1px solid rgba(128, 128, 128, 0.25);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    notes = ss.get("acct_flash", [])
    ss.acct_flash = []

    unit = ss.get("unit")
    user = ss.get("user")
    latest = {}
    patients = {}
    load_err = None
    if unit:
        beds = UNITS[unit]
        try:
            latest = latest_by_bed(unit)
            beds = all_beds(unit)  # 預設床位 + 自訂床位
        except Exception as e:  # 試算表暫時連不上: 先顯示預設床位
            load_err = e
        if ss.get("bed_sel") not in beds:
            ss.bed_sel = UNITS[unit][0]
    if user:
        try:
            patients = {p["id"]: p["name"] for p in acct_list_patients(user["id"])}
        except Exception as e:
            load_err = e
        if ss.get("patient_id") not in patients:
            if patients:
                ss.patient_id = next(iter(patients))
            else:
                ss.pop("patient_id", None)
    pid = ss.get("patient_id") if user else None
    dialog_to_open = None

    def bed_label(bed):
        if bed in latest:
            t, d = latest[bed]
            return f"{bed} ● {d.get('bw_g', 0):g} g · {t[5:]}"
        return f"{bed} (空床)"

    with st.container(key="app_header"):
        h_title, h_ctrl = st.columns([5, 6], vertical_alignment="center")
        h_title.markdown(title_html, unsafe_allow_html=True)
        with h_ctrl.container(key="hdr_ctrl"):
            # 一排靠右: [提示文字] [床位 / 病患選單] [儲存] [紀錄] [帳號]
            if not (unit or user):
                c_note, c_login = st.columns(2, vertical_alignment="center")
                with c_note.container(key="notice"):
                    render_notes(notes)
                if c_login.button("🔑 登入 / 註冊",
                                  help="PICU / NBC 單位帳號或個人帳號:登入後可儲存與查看歷史紀錄"):
                    dialog_to_open = "auth"
            else:
                c_note, c_sel, c_save, c_hist, c_acct = st.columns(
                    5, vertical_alignment="center")
                with c_note.container(key="notice"):
                    render_notes(notes)
                if unit:
                    with c_sel.container(key="bed_box"):
                        st.selectbox("床位", beds, key="bed_sel", format_func=bed_label,
                                     on_change=cb_select_bed, label_visibility="collapsed")
                    can_save = True
                else:
                    with c_sel.container(key="pat_sel"):
                        if patients:
                            st.selectbox("目前病患", list(patients),
                                         format_func=lambda i: patients[i],
                                         key="patient_id", on_change=cb_select_patient,
                                         label_visibility="collapsed")
                        else:
                            st.caption("尚無病患")
                    can_save = bool(pid)
                if c_save.button("💾 儲存", disabled=not can_save,
                                 help="把目前輸入的資料存到這個" + ("床位" if unit else "病患")):
                    dialog_to_open = "save"
                if c_hist.button("🕘 紀錄", disabled=not can_save,
                                 help="查看、載入或刪除歷史紀錄"):
                    dialog_to_open = "history"
                with c_acct.popover(f"👤 {unit}" if unit else "👤 帳號"):
                    if unit:
                        st.caption(f"單位帳號:**{unit}**")
                        st.caption(f"儲存位置:{storage_name()}")
                        if st.button("🛏️ 新增 / 管理床位", use_container_width=True):
                            dialog_to_open = "beds"
                        if st.button(f"🗑️ 清空 {ss.bed_sel} 資料", use_container_width=True):
                            dialog_to_open = "clear"
                    else:
                        st.caption(f"登入身分:**{user['username']}**")
                        if st.button("👥 新增 / 管理病患", use_container_width=True):
                            dialog_to_open = "patients"
                    st.button("🚪 登出", on_click=cb_logout, use_container_width=True)
    show_db_error()
    if load_err is not None:
        st.error(f"讀取紀錄失敗 ({type(load_err).__name__}):{load_err}\n\n"
                 f"儲存位置:{_safe_storage_label()}")
    return dialog_to_open, patients, pid, latest


def open_account_dialog(which, patients: dict, pid, payload) -> None:
    """開啟標題列按下的對話框 (放在 main 最後,儲存時才拿得到計算結果)。"""
    if which == "auth":
        auth_dialog()
        return
    if which == "patients":
        patients_dialog()
        return
    unit = ss.get("unit")
    if unit:
        bed = ss.bed_sel
        target, label = ("unit", unit, bed), f"{unit} · {bed}"
        if which == "clear":
            clear_bed_dialog(unit, bed)
            return
        if which == "beds":
            beds_dialog(unit)
            return
    elif pid:
        target, label = ("acct", pid), patients[pid]
    else:
        return
    if which == "save" and payload:
        save_dialog(target, label, payload)
    elif which == "history":
        history_dialog(target, label)


def main():
    _logo = Path(__file__).parent / "logo.png"
    st.set_page_config(page_title="NICU 營養計算器",
                       page_icon=str(_logo) if _logo.exists() else "🍼", layout="wide",
                       initial_sidebar_state="expanded")
    # 個人帳號: 套用上一個動作排定的輸入值 (必須在任何輸入欄位建立之前)
    st.session_state.pop("_acct_rerun", None)
    if "_acct_pending" in st.session_state:
        apply_record(st.session_state.pop("_acct_pending"))
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
    ss = st.session_state
    for k, v in DEFAULTS.items():
        ss.setdefault(k, v)

    # 標題列 + 登入 (單位帳號 PICU / NBC 或個人帳號) / 床位或病患 / 儲存 / 紀錄
    acct_dialog, acct_patients, acct_pid, latest = render_account_header(
        f"<h1 style='margin-bottom:0; display:flex; align-items:center; flex-wrap:wrap;'>"
        f"{logo_html}NICU 營養計算器</h1>")
    acct_payload = None
    unit = ss.get("unit")

    sb = st.sidebar
    if not (unit or ss.get("user")):
        sb.caption("未登入也可直接計算,但不會儲存紀錄")
        sb.button("🗑️ 清空輸入 (重設為預設值)", on_click=do_reset_inputs,
                  use_container_width=True)

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
            for b_ in all_beds(unit):
                row = {"床位": b_, "上次儲存": "—", "備註": "", "BW (g)": ""}
                if b_ in latest:
                    t, d = latest[b_]
                    row.update({"上次儲存": t, "備註": d.get("note") or d.get("bed") or "",
                                "BW (g)": fmt(d.get("bw_g", 0)), **_history_cols(d)})
                rows.append(row)
            # 欄位順序與歷史紀錄相同
            cols = ["床位", "上次儲存", "備註", "BW (g)"] + list(_history_cols({}))
            st.caption(f"{unit} 各床最近一次紀錄 (於上方選擇床位即可載入)")
            st.dataframe(pd.DataFrame(rows, columns=cols).fillna(""), hide_index=True,
                         use_container_width=True, height=min(40 + 35 * len(rows), 950))
            st.caption(HISTORY_UNITS_NOTE)

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

        # 個人帳號儲存用的資料 (格式與床位紀錄相同)
        acct_payload = {k: (to_dec1(ss[k]) if type(DEFAULTS[k]) is float else ss[k])
                        for k in DEFAULTS}
        acct_payload.update(milk_comp=p["milk_comp"], fort_comp=p["fort_comp"], summary=s)

        # --- 頂部摘要 ---
        with top:
            if unit:
                st.markdown(f"#### 🛏️ {unit} · {ss['bed_sel']}"
                            + (f" — {ss['bed']}" if ss["bed"] else ""))
            if acct_pid:
                st.markdown(f"#### 👤 {acct_patients[acct_pid]}")
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

    # 開啟標題列按下的對話框 (放最後,儲存時才拿得到計算結果)
    open_account_dialog(acct_dialog, acct_patients, acct_pid, acct_payload)


if __name__ == "__main__":
    main()
