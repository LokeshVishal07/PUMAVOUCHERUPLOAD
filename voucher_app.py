import streamlit as st
import pandas as pd
import re
import io
import zipfile
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PUMA Voucher SKU Tool",
    page_icon="🏷️",
    layout="wide"
)

# ─────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────

REGION_MARKETPLACES = {
    "PH": ["Lazada", "Shopee", "Zalora"],
    "MY": ["Lazada", "Shopee", "Zalora", "TikTok"],
    "SG": ["Lazada", "Shopee", "Zalora"],
}

# All column indices are 0-based positional (iloc)
REGION_CONFIG = {
    "PH": {
        "zecom_sheet": "PH",
        "zecom_read": "ph",          # special double-header structure
        "article_col": "PIM Article#",
        "rrp_idx": 32,
        "srp_idx": 50,
        "threshold": 650,
        "currency": "PHP > 650",
        "mp_flags": {
            "Lazada": "LAZADA",
            "Shopee": "SHOPEE",
            "Zalora": "ZALORA",
        },
        "campaign_types": ["NON SBD", "SBD"],
        "excl_idx": {"NON SBD": 71, "SBD": 76},
    },
    "MY": {
        "zecom_sheet": "MY",
        "zecom_read": "header3",
        "article_col": "Style#",
        # Lazada/Zalora/TikTok block: RRP=48, SRP=49, Excl=51
        # Shopee BAU block:           RRP=54, SRP=55, Excl=66
        "rrp_idx":  {"Shopee": 54, "default": 48},
        "srp_idx":  {"Shopee": 55, "default": 49},
        "threshold": 36,
        "currency": "RM > 36",
        "mp_flags": {
            "Lazada":  "Lazada",
            "Shopee":  "Shopee",
            "Zalora":  "Zalora MP",
            "TikTok":  "TIKTOK",
        },
        "campaign_types": None,       # determined by marketplace
        "excl_idx": {"Shopee": 66, "default": 51},
    },
    "SG": {
        "zecom_sheet": "SG",
        "zecom_read": "header3",
        "article_col": "STYLE#",
        "rrp_idx": 26,
        "srp_idx": 50,
        "threshold": 16,
        "currency": "SGD > 16",
        "mp_flags": {
            "Lazada": "Lazada",
            "Shopee": "Shopee",
            "Zalora": "Zalora",
        },
        "campaign_types": ["BAU", "Campaign"],
        "excl_idx": {"BAU": 52, "Campaign": 56},
    },
}

# ─────────────────────────────────────────────────────────────────
# ZECOM READING & VALIDATION
# ─────────────────────────────────────────────────────────────────

def validate_zecom_region(file_bytes: bytes, selected_region: str):
    """Check that uploaded ZeCom file matches selected region."""
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
        sheets = wb.sheetnames
        wb.close()
    except Exception as e:
        return False, f"Cannot read ZeCom file: {e}"

    if selected_region == "PH":
        if "PH" not in sheets:
            extra = " (looks like MY/SG tracker)" if ("MY" in sheets or "SG" in sheets) else ""
            return False, f"⚠️ Wrong file! Selected **PH** but 'PH' sheet not found{extra}."
    else:
        if selected_region not in sheets:
            extra = " (looks like PH tracker)" if "PH" in sheets else ""
            return False, f"⚠️ Wrong file! Selected **{selected_region}** but '{selected_region}' sheet not found{extra}."

    return True, "OK"


def read_zecom(file_bytes: bytes, region: str) -> pd.DataFrame:
    """Read ZeCom tracker for given region."""
    cfg = REGION_CONFIG[region]
    sheet = cfg["zecom_sheet"]

    if cfg["zecom_read"] == "ph":
        # PH has a double-row header: real column names are in row 2 of the raw file
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="PH", header=1)
        df.columns = df.iloc[0]          # row 0 of df = actual column names
        df = df.iloc[1:].reset_index(drop=True)
    else:
        # MY and SG: header is at row index 3
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=3)

    return df


# ─────────────────────────────────────────────────────────────────
# INVENTORY READING
# ─────────────────────────────────────────────────────────────────

def _normalize_inv(df: pd.DataFrame, ean_cands: list, stock_cands: list):
    """Find EAN + stock cols and return clean DataFrame."""
    ean_col   = next((c for c in ean_cands   if c in df.columns), None)
    stock_col = next((c for c in stock_cands if c in df.columns), None)
    if ean_col is None or stock_col is None:
        return None
    out = df[[ean_col, stock_col]].copy()
    out.columns = ["EAN", "Stock"]
    out["EAN"]   = out["EAN"].astype(str).str.strip()
    out["Stock"] = pd.to_numeric(out["Stock"], errors="coerce").fillna(0)
    out = out[out["EAN"].str.match(r"^\d{13}$")]  # keep 13-digit EANs only
    return out


def read_inventory(file_bytes: bytes, filename: str):
    """Auto-detect EAN and stock columns from inventory file."""
    ean_cands   = ["EAN", "Sku", "PROD_CODE", "SellerSku"]
    stock_cands = ["Avail_Qty", "QtyAvailable", "QTY", "Quantity"]

    if filename.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
        return _normalize_inv(df, ean_cands, stock_cands)

    # Try standard header (row 0) — works for PH xlsx and MY
    df = pd.read_excel(io.BytesIO(file_bytes))
    result = _normalize_inv(df, ean_cands, stock_cands)
    if result is not None:
        return result

    # Try skipping 4 report header rows — works for SG
    df = pd.read_excel(io.BytesIO(file_bytes), header=4)
    return _normalize_inv(df, ean_cands, stock_cands)


# ─────────────────────────────────────────────────────────────────
# VOUCHER ELIGIBILITY LOGIC
# ─────────────────────────────────────────────────────────────────

def classify_remark(remark, voucher_pct: int, voucher_type: str, marketplace: str = None) -> str:
    """
    Returns: 'eligible' | 'ineligible' | 'no_remark'

    Bundle discount: only OPEN FOR ALL variants are eligible.
    Regular VC:
      - OPEN FOR ALL* → any %
      - OPEN for X days up to Y% → any % ≤ Y
      - N% VC ONLY / N% VC MAX / MAX N% VC / N% VC - * → exactly N%
      - N% VC ONLY - Shopee exclusive → only on Shopee at exactly N%
      - EXCLUDE* → never
    """
    if pd.isna(remark):
        return "no_remark"
    r = str(remark).strip()
    if not r or r.lower() == "nan":
        return "no_remark"

    rl = r.lower()

    # EXCLUDE wins first
    if rl.startswith("exclude"):
        return "ineligible"

    # OPEN FOR ALL → eligible for any voucher type and any %
    if rl.startswith("open for all"):
        return "eligible"

    # Bundle: only OPEN FOR ALL qualifies
    if voucher_type == "bundle":
        return "ineligible"

    # "OPEN for N days up to Y%"
    m = re.search(r"open for.*?up to\s*(\d+)%", rl)
    if m:
        return "eligible" if voucher_pct <= int(m.group(1)) else "ineligible"

    # Exact % match
    m = re.search(r"(\d+)%", r)
    if m:
        remark_pct = int(m.group(1))
        if remark_pct == voucher_pct:
            if "shopee exclusive" in rl:
                return "eligible" if (marketplace and "shopee" in marketplace.lower()) else "ineligible"
            return "eligible"
        return "ineligible"

    return "ineligible"


# ─────────────────────────────────────────────────────────────────
# ZECOM ARTICLE-LEVEL PROCESSING
# ─────────────────────────────────────────────────────────────────

def process_zecom(zecom_df: pd.DataFrame, region: str, marketplace: str,
                  campaign_type: str, voucher_pct: int, voucher_type: str) -> pd.DataFrame:
    """
    Returns DataFrame: [article, remark_status]
    remark_status: 'eligible' | 'ineligible' | 'no_remark'
    Already accounts for price threshold and MP flag filter.
    """
    cfg = REGION_CONFIG[region]
    df  = zecom_df.copy()

    # ── MP flag filter ────────────────────────────────────────────
    mp_col = cfg["mp_flags"].get(marketplace)
    if mp_col and mp_col in df.columns:
        df = df[df[mp_col].astype(str).str.strip().str.upper() == "YES"].copy()

    if df.empty:
        return pd.DataFrame(columns=["article", "remark_status"])

    # ── Price columns (positional) ────────────────────────────────
    if region == "MY":
        rrp_idx = cfg["rrp_idx"].get(marketplace, cfg["rrp_idx"]["default"])
        srp_idx = cfg["srp_idx"].get(marketplace, cfg["srp_idx"]["default"])
    else:
        rrp_idx = cfg["rrp_idx"]
        srp_idx = cfg["srp_idx"]

    threshold = cfg["threshold"]
    rrp = pd.to_numeric(df.iloc[:, rrp_idx], errors="coerce")
    srp = pd.to_numeric(df.iloc[:, srp_idx], errors="coerce")
    price_ok = (rrp > threshold) & (srp > threshold)

    # ── Exclusion column (positional) ─────────────────────────────
    if region == "MY":
        excl_idx = cfg["excl_idx"].get(marketplace, cfg["excl_idx"]["default"])
    else:
        excl_idx = cfg["excl_idx"][campaign_type]

    excl_vals = df.iloc[:, excl_idx]

    # ── Build result ──────────────────────────────────────────────
    article_col = cfg["article_col"]
    result = pd.DataFrame({
        "article":    df[article_col].astype(str).str.strip().values,
        "price_ok":   price_ok.values,
        "remark":     excl_vals.values,
    })

    # Price fail → ineligible regardless of remark
    result["remark_status"] = result.apply(
        lambda row: "ineligible" if not row["price_ok"]
        else classify_remark(row["remark"], voucher_pct, voucher_type, marketplace),
        axis=1
    )

    # Drop empty / NaN articles
    result = result[result["article"].str.match(r"^[\w_\-]+$", na=False)]
    result = result[result["article"].str.lower() != "nan"]

    return result[["article", "remark_status"]].drop_duplicates(subset=["article"])


# ─────────────────────────────────────────────────────────────────
# EAN ELIGIBILITY MAPPING
# ─────────────────────────────────────────────────────────────────

def map_to_eans(article_status: pd.DataFrame,
                content_df: pd.DataFrame,
                inventory_df: pd.DataFrame) -> pd.DataFrame:
    """
    Joins: article_status → content (article→EAN) → inventory (EAN→stock)
    Returns: [article, EAN, remark_status, has_stock]
    """
    merged = article_status.merge(
        content_df.rename(columns={"Color_No": "article"}),
        on="article", how="inner"
    )
    merged["EAN"] = merged["EAN"].astype(str).str.strip()

    merged = merged.merge(
        inventory_df.rename(columns={"Stock": "stock"}),
        on="EAN", how="left"
    )
    merged["has_stock"] = merged["stock"].fillna(0) > 0

    return merged[["article", "EAN", "remark_status", "has_stock"]]


def eligible_ean_set(ean_df: pd.DataFrame) -> set:
    """EANs that are eligible + in stock."""
    return set(ean_df[(ean_df["remark_status"] == "eligible") & ean_df["has_stock"]]["EAN"])


def excluded_ean_set(ean_df: pd.DataFrame) -> set:
    """EANs from articles flagged ineligible (for PID-level exclusion)."""
    return set(ean_df[ean_df["remark_status"] == "ineligible"]["EAN"])


def no_remark_ean_set(ean_df: pd.DataFrame) -> set:
    """EANs with no remark (flagged separately)."""
    return set(ean_df[ean_df["remark_status"] == "no_remark"]["EAN"])


# ─────────────────────────────────────────────────────────────────
# MARKETPLACE PROCESSORS
# ─────────────────────────────────────────────────────────────────

def process_lazada(ean_df: pd.DataFrame, lazada_bytes: bytes) -> list:
    """Return list of eligible Lazada Shop SKUs."""
    df = pd.read_excel(io.BytesIO(lazada_bytes), sheet_name="template", header=0)
    data = df.iloc[3:].copy()
    data.columns = df.columns

    active = data[data["status"].astype(str).str.lower() == "active"].copy()
    active["_ean"] = active["SellerSKU"].astype(str).str.strip()

    ok_eans = eligible_ean_set(ean_df)
    matched = active[active["_ean"].isin(ok_eans)]
    return matched["Shop SKU"].dropna().unique().tolist()


def _read_shopee_zip(zip_bytes: bytes) -> pd.DataFrame:
    """Read and combine all xlsx files from Shopee export zip."""
    dfs = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = sorted(n for n in zf.namelist() if n.endswith(".xlsx"))
        progress = st.progress(0, text="Reading Shopee export files…")
        for i, name in enumerate(names):
            with zf.open(name) as f:
                df = pd.read_excel(f, engine="calamine", header=2, skiprows=[3, 4])
                dfs.append(df)
            progress.progress((i + 1) / len(names), text=f"Reading Shopee file {i+1}/{len(names)}…")
        progress.empty()
    return pd.concat(dfs, ignore_index=True)


def _extract_shopee_ean(row) -> str | None:
    """Extract 13-digit EAN from SKU (float) or Parent SKU."""
    sku = row.get("SKU") if hasattr(row, "get") else row["SKU"]
    parent = str(row.get("Parent SKU", "") or "").strip()
    if pd.notna(sku):
        try:
            s = str(int(float(sku)))
            if re.match(r"^\d{13}$", s):
                return s
        except (ValueError, TypeError):
            pass
    if re.match(r"^\d{13}$", parent):
        return parent
    return None


def _pid_eligible(combined: pd.DataFrame, ok_eans: set, excl_eans: set) -> list:
    """
    PID-level eligibility:
      - Exclude entire PID if ANY variant EAN is in excluded set
      - Include PID if it has at least one EAN in eligible set
    """
    combined["_ean"] = combined.apply(_extract_shopee_ean, axis=1)
    combined["_pid"] = combined["Product ID"].astype(str).str.strip()

    result = []
    for pid, grp in combined.groupby("_pid"):
        g_eans = set(grp["_ean"].dropna())
        if g_eans & excl_eans:     # any ineligible variant → skip whole PID
            continue
        if g_eans & ok_eans:       # has at least one eligible variant
            result.append(pid)
    return result


def process_shopee(ean_df: pd.DataFrame, zip_bytes: bytes) -> list:
    """Return list of eligible Shopee Product IDs."""
    combined = _read_shopee_zip(zip_bytes)
    return _pid_eligible(combined, eligible_ean_set(ean_df), excluded_ean_set(ean_df))


def process_zalora(ean_df: pd.DataFrame, eligible_bytes: bytes,
                   content_df: pd.DataFrame) -> pd.DataFrame:
    """Annotate Zalora EligibleProducts file with Article No + Voucher Eligible column."""
    df = pd.read_excel(io.BytesIO(eligible_bytes), sheet_name="Eligible Products")

    ok_eans       = eligible_ean_set(ean_df)
    no_rmk_eans   = no_remark_ean_set(ean_df)

    ean_to_article = dict(zip(
        content_df["EAN"].astype(str).str.strip(),
        content_df["Color_No"].astype(str).str.strip()
    ))

    df["_ean"] = df["Seller SKU"].astype(str).str.strip()
    df["Article No"] = df["_ean"].map(ean_to_article)

    def label(ean):
        if ean in ok_eans:     return "Yes"
        if ean in no_rmk_eans: return "No Remark"
        return "No"

    df["Voucher Eligible"] = df["_ean"].apply(label)
    df.drop(columns=["_ean"], inplace=True)
    return df


def process_tiktok(ean_df: pd.DataFrame, tiktok_bytes: bytes) -> list:
    """Return list of eligible TikTok Product IDs (same PID logic as Shopee)."""
    df = pd.read_excel(io.BytesIO(tiktok_bytes), sheet_name="Template",
                       header=2, skiprows=[3, 4])

    def extract_ean(val):
        if pd.notna(val):
            try:
                s = str(int(float(val)))
                if re.match(r"^\d{13}$", s):
                    return s
            except (ValueError, TypeError):
                pass
        return None

    df["_ean"] = df["Seller SKU"].apply(extract_ean)
    df["_pid"] = df["Product ID"].astype(str).str.strip()

    ok_eans   = eligible_ean_set(ean_df)
    excl_eans = excluded_ean_set(ean_df)

    result = []
    for pid, grp in df.groupby("_pid"):
        g_eans = set(grp["_ean"].dropna())
        if g_eans & excl_eans:
            continue
        if g_eans & ok_eans:
            result.append(pid)
    return result


# ─────────────────────────────────────────────────────────────────
# OUTPUT GENERATION
# ─────────────────────────────────────────────────────────────────

def _to_excel(df: pd.DataFrame, sheet: str = "Sheet1") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    return buf.getvalue()


def make_lazada_output(shop_skus: list) -> bytes:
    return _to_excel(pd.DataFrame({"SHOP SKU": shop_skus}))


def make_shopee_output(product_ids: list) -> bytes:
    return _to_excel(pd.DataFrame({"Product ID": product_ids}))


def make_zalora_output(annotated: pd.DataFrame) -> bytes:
    return _to_excel(annotated, sheet="Eligible Products")


# ─────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────

def main():
    st.title("🏷️ PUMA Voucher Eligible SKU Tool")
    st.caption("Generate marketplace-ready voucher eligible SKU lists from ZeCom tracker data.")

    # ── STEP 1 ── REGION & MARKETPLACE ───────────────────────────
    st.markdown("---")
    st.subheader("① Region & Marketplace")

    c1, c2, c3 = st.columns(3)
    with c1:
        region = st.selectbox("Region", ["PH", "MY", "SG"])
    with c2:
        marketplace = st.selectbox("Marketplace", REGION_MARKETPLACES[region])
    with c3:
        cfg = REGION_CONFIG[region]
        if cfg["campaign_types"]:
            campaign_type = st.selectbox("Campaign Type", cfg["campaign_types"])
        else:
            campaign_type = None
            excl_label = "Shopee exclusion column" if marketplace == "Shopee" \
                         else "Lazada / Zalora / TikTok exclusion column"
            st.info(f"📋 Using: {excl_label}")

    # ── STEP 2 ── VOUCHER CONFIG ──────────────────────────────────
    st.markdown("---")
    st.subheader("② Voucher Configuration")

    cv1, cv2 = st.columns([1, 2])
    with cv1:
        voucher_type = st.radio("Voucher Type", ["Regular VC", "Bundle Discount"], horizontal=True)
    with cv2:
        pct_raw = st.text_input("Voucher % — comma-separated (e.g. 10, 20, 50)",
                                 placeholder="10, 20, 50")

    voucher_pcts = []
    for tok in re.split(r"[,\s%]+", pct_raw.strip()):
        if tok.isdigit():
            voucher_pcts.append(int(tok))

    if voucher_pcts:
        vt_label = "Bundle Discount" if voucher_type == "Bundle Discount" else "Regular VC"
        st.success(
            f"Will generate: **{', '.join(str(p)+'%' for p in voucher_pcts)} "
            f"{vt_label}** lists for **{region} — {marketplace}**"
        )

    # ── STEP 3 ── FILE UPLOADS ────────────────────────────────────
    st.markdown("---")
    st.subheader("③ Upload Files")

    cf1, cf2 = st.columns(2)

    with cf1:
        st.markdown("**Core files (always required)**")
        zecom_file = st.file_uploader(
            f"ZeCom Tracker ({'PH file' if region == 'PH' else 'MY + SG file'})",
            type=["xlsx"], key="zecom"
        )
        content_file = st.file_uploader("Content File", type=["xlsx"], key="content")
        inv_file = st.file_uploader(
            f"Inventory File ({region})", type=["xlsx", "csv"], key="inv"
        )

    with cf2:
        st.markdown(f"**{marketplace} marketplace export**")
        if marketplace == "Lazada":
            mp_file = st.file_uploader("Lazada Product Export (.xlsx)", type=["xlsx"], key="mp")
            st.caption("Price/stock export from Lazada Seller Center")
        elif marketplace == "Shopee":
            mp_file = st.file_uploader("Shopee Export ZIP (.zip)", type=["zip"], key="mp")
            st.caption("Mass update export — all batch files in one zip")
        elif marketplace == "Zalora":
            mp_file = st.file_uploader("Zalora EligibleProducts File (.xlsx)", type=["xlsx"], key="mp")
            st.caption("EligibleProducts export from the Zalora voucher campaign")
        elif marketplace == "TikTok":
            mp_file = st.file_uploader("TikTok Seller Center Export (.xlsx)", type=["xlsx"], key="mp")
            st.caption("Batch edit / sales information export from TikTok")

    # ── STEP 4 ── GENERATE ────────────────────────────────────────
    st.markdown("---")
    st.subheader("④ Generate")

    missing = []
    if not zecom_file:   missing.append("ZeCom Tracker")
    if not content_file: missing.append("Content File")
    if not inv_file:     missing.append("Inventory File")
    if not mp_file:      missing.append(f"{marketplace} Export")
    if not voucher_pcts: missing.append("Voucher %")

    if missing:
        st.info(f"Still needed: **{', '.join(missing)}**")

    ready = not missing

    if st.button("🚀 Generate Eligible SKU Lists", disabled=not ready, type="primary"):
        _run(zecom_file, content_file, inv_file, mp_file,
             region, marketplace, campaign_type, voucher_pcts, voucher_type)


# ─────────────────────────────────────────────────────────────────
# PROCESSING PIPELINE
# ─────────────────────────────────────────────────────────────────

def _run(zecom_file, content_file, inv_file, mp_file,
         region, marketplace, campaign_type, voucher_pcts, voucher_type):

    vtype = "bundle" if voucher_type == "Bundle Discount" else "regular"

    with st.status("Processing…", expanded=True) as status:

        # 1. Validate ZeCom region
        st.write("🔍 Validating ZeCom file region…")
        ok, msg = validate_zecom_region(zecom_file.getvalue(), region)
        if not ok:
            st.error(msg)
            status.update(label="❌ Validation failed", state="error")
            return

        # 2. Read ZeCom
        st.write(f"📊 Reading ZeCom — {region}…")
        try:
            zecom_df = read_zecom(zecom_file.getvalue(), region)
            st.write(f"   ✓ {len(zecom_df):,} rows loaded")
        except Exception as e:
            st.error(f"ZeCom read error: {e}")
            status.update(label="❌ Error", state="error")
            return

        # 3. Read Content
        st.write("📦 Reading content file…")
        try:
            content_df = pd.read_excel(io.BytesIO(content_file.getvalue()), sheet_name="content")
            content_df = content_df[["Color_No", "EAN"]].dropna()
            content_df["EAN"]      = content_df["EAN"].astype(str).str.strip()
            content_df["Color_No"] = content_df["Color_No"].astype(str).str.strip()
            st.write(f"   ✓ {len(content_df):,} EAN mappings")
        except Exception as e:
            st.error(f"Content file error: {e}")
            status.update(label="❌ Error", state="error")
            return

        # 4. Read Inventory
        st.write("🏭 Reading inventory…")
        try:
            inv_df = read_inventory(inv_file.getvalue(), inv_file.name)
            if inv_df is None:
                st.error("Could not detect EAN/Stock columns in inventory file. "
                         "Expected column names: EAN/Sku/PROD_CODE and Avail_Qty/QtyAvailable/QTY.")
                status.update(label="❌ Error", state="error")
                return
            in_stock = (inv_df["Stock"] > 0).sum()
            st.write(f"   ✓ {len(inv_df):,} EANs | {in_stock:,} in stock")
        except Exception as e:
            st.error(f"Inventory read error: {e}")
            status.update(label="❌ Error", state="error")
            return

        # 5. Process each voucher %
        results = {}

        for pct in voucher_pcts:
            st.write(f"⚙️  Processing **{pct}% {voucher_type}**…")

            art = process_zecom(zecom_df, region, marketplace, campaign_type, pct, vtype)
            n_elig  = (art["remark_status"] == "eligible").sum()
            n_nrmk  = (art["remark_status"] == "no_remark").sum()
            n_ineli = (art["remark_status"] == "ineligible").sum()
            st.write(f"   ZeCom articles → {n_elig} eligible, "
                     f"{n_nrmk} no-remark, {n_ineli} ineligible")

            if n_elig + n_nrmk == 0:
                st.warning(f"   No eligible / no-remark articles found for {pct}% — skipping.")
                continue

            ean_df = map_to_eans(art, content_df, inv_df)
            n_ok_eans = len(eligible_ean_set(ean_df))
            st.write(f"   EANs in stock & eligible → {n_ok_eans:,}")

            if n_ok_eans == 0:
                st.warning(f"   No in-stock eligible EANs for {pct}% — skipping.")
                continue

            try:
                if marketplace == "Lazada":
                    ids = process_lazada(ean_df, mp_file.getvalue())
                    st.write(f"   ✅ Lazada Shop SKUs → **{len(ids)}**")
                    results[pct] = {"mp": "Lazada", "ids": ids}

                elif marketplace == "Shopee":
                    ids = process_shopee(ean_df, mp_file.getvalue())
                    st.write(f"   ✅ Shopee Product IDs → **{len(ids)}**")
                    results[pct] = {"mp": "Shopee", "ids": ids}

                elif marketplace == "Zalora":
                    ann = process_zalora(ean_df, mp_file.getvalue(), content_df)
                    y   = (ann["Voucher Eligible"] == "Yes").sum()
                    nr  = (ann["Voucher Eligible"] == "No Remark").sum()
                    st.write(f"   ✅ Zalora eligible → **{y}** Yes, {nr} No Remark")
                    results[pct] = {"mp": "Zalora", "ann": ann, "yes_count": y}

                elif marketplace == "TikTok":
                    ids = process_tiktok(ean_df, mp_file.getvalue())
                    st.write(f"   ✅ TikTok Product IDs → **{len(ids)}**")
                    results[pct] = {"mp": "TikTok", "ids": ids}

            except Exception as e:
                st.error(f"   Error processing {marketplace} for {pct}%: {e}")

        if results:
            status.update(label="✅ Done!", state="complete")
        else:
            status.update(label="⚠️ No results generated", state="error")
            return

    # ── DOWNLOAD RESULTS ─────────────────────────────────────────
    st.markdown("---")
    st.subheader("⑤ Download Results")

    today = pd.Timestamp.now().strftime("%Y%m%d")
    vt_short = "Bundle" if vtype == "bundle" else "VC"

    for pct, res in results.items():
        mp = res["mp"]
        fname = f"{mp}_{region}_{pct}pct_{vt_short}_{today}.xlsx"

        col_m, col_d = st.columns([3, 2])

        with col_m:
            if mp == "Zalora":
                count = res["yes_count"]
                st.metric(f"{mp} — {pct}% {vt_short}", f"{count} eligible SKUs")
            else:
                count = len(res["ids"])
                id_label = "Shop SKUs" if mp == "Lazada" else "Product IDs"
                st.metric(f"{mp} — {pct}% {vt_short}", f"{count} {id_label}")

        with col_d:
            if mp == "Lazada":
                data = make_lazada_output(res["ids"])
            elif mp in ("Shopee", "TikTok"):
                data = make_shopee_output(res["ids"])
            elif mp == "Zalora":
                data = make_zalora_output(res["ann"])

            st.download_button(
                f"⬇️ Download {fname}",
                data=data,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{mp}_{pct}",
            )


# ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
