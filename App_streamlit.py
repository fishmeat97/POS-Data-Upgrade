"""
POS Monthly Data Mapping  -  Streamlit UI  (v2.0)

All business logic lives in pos_core.py (no Streamlit there). This file only contains:
upload widgets, the Start/Reset flow, KPI dashboard, submission tracking view, previews and downloads.

Run:   streamlit run app_streamlit.py          (pos_core.py must sit next to this file)
"""

from pathlib import Path

import pandas as pd
import streamlit as st

import Pos_core as core

# ============================================================
# Page config  (MUST be the first Streamlit call - v1 called st.write before it)
# ============================================================
st.set_page_config(page_title="POS Monthly Data Mapping", page_icon="🥃", layout="wide")


# ============================================================
# UI helpers
# ============================================================
def section_header(step_number, title, subtitle=""):
    st.markdown(
        f"""
<div style="background: linear-gradient(90deg, rgba(80,80,80,0.10), rgba(80,80,80,0.03));
            padding: 18px 22px; border-radius: 14px; border: 1px solid rgba(128,128,128,0.18);
            margin-top: 24px; margin-bottom: 16px;">
  <div style="display:flex; align-items:center; gap:14px;">
    <div style="width:44px; height:44px; min-width:44px; border-radius:50%;
                background: rgba(120,120,120,0.18); display:flex; align-items:center;
                justify-content:center; font-size:21px; font-weight:700;">{step_number}</div>
    <div>
      <div style="font-size:23px; font-weight:700; margin-bottom:2px;">{title}</div>
      <div style="font-size:14px; color:#888;">{subtitle}</div>
    </div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


st.markdown(
    """
<style>
div.stButton > button[kind="primary"] { background-color:#4A4A4A !important; color:white !important;
    border:1px solid #4A4A4A !important; font-weight:600 !important; }
div.stButton > button[kind="primary"]:hover { background-color:#3A3A3A !important; border-color:#3A3A3A !important; }
div.stButton > button[kind="primary"]:active { background-color:#2F2F2F !important; border-color:#2F2F2F !important; }
</style>
""",
    unsafe_allow_html=True,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def show_df(df, height=None, **kw):
    if height is None:
        height = "auto"  # explicit default; passing None is rejected by recent Streamlit versions
    st.dataframe(df, width="stretch", hide_index=True, height=height, **kw)


# ============================================================
# Header
# ============================================================
header_col1, header_col2 = st.columns([0.8, 6], vertical_alignment="center", gap="small")
with header_col1:
    if Path("logo.png").exists():  # v1 crashed when logo.png was missing
        st.image("logo.png", width=160)
with header_col2:
    st.markdown(
        """
<h1 style="margin-bottom:0; padding-bottom:0; margin-left:-20px;">POS Monthly Data Mapping</h1>
<p style="color:grey; font-size:16px; margin-top:4px; margin-left:-10px;">
Monthly POS Data Integration, Mapping &amp; Submission Dashboard</p>
""",
        unsafe_allow_html=True,
    )
st.divider()


# ============================================================
# STEP 1 - Upload monthly POS files
# ============================================================
@st.cache_data(show_spinner=False)
def cached_extract_zip(data: bytes):
    return core.extract_pos_zip(data)


section_header("1", "Upload monthly POS templates", "上傳當月份資料")

upload_mode = st.radio(
    "Upload Method",
    ["📦 ZIP File (批次上傳)", "📄 Excel / CSV (多檔上傳)"],
    horizontal=True,
)

template_files = []  # list[(filename, bytes)]

if upload_mode.startswith("📦"):
    zip_file = st.file_uploader("上傳當月份 Upload POS ZIP File", type=["zip"], key="pos_zip", on_change="ignore")
    if zip_file is not None:
        try:
            template_files, ignored_files = cached_extract_zip(zip_file.getvalue())
        except Exception as exc:
            st.error(f"ZIP 解壓縮失敗（檔案無法讀取或已損毀）：{exc}")
            st.stop()
        if template_files:
            total_mb = sum(len(b) for _, b in template_files) / 1024 / 1024
            st.success(f"ZIP 解壓縮成功：{len(template_files)} 個可處理檔案")
            c1, c2, c3 = st.columns(3)
            c1.metric("可處理檔案", len(template_files))
            c2.metric("總大小", f"{total_mb:.2f} MB")
            c3.metric("忽略檔案", len(ignored_files))
            with st.expander("📂 查看 ZIP 內檔案", expanded=False):
                show_df(pd.DataFrame([
                    {"檔案": n, "大小 KB": round(len(b) / 1024, 1), "格式": Path(n).suffix.lower()}
                    for n, b in template_files]))
            if ignored_files:
                with st.expander("⚠️ 被忽略的檔案", expanded=False):
                    for n in ignored_files:
                        st.write(n)
        else:
            st.warning("ZIP 裡沒有找到可處理的 Excel / CSV 檔案。")
else:
    uploaded = st.file_uploader(
        "可一次上傳多個 Excel / CSV Template", type=["xlsx", "xls", "csv"],
        accept_multiple_files=True, key="templates", on_change="ignore")
    if uploaded:
        template_files = [(f.name, f.getvalue()) for f in uploaded]
        st.caption(f"已選擇 {len(template_files)} 個 Template ｜ 總大小 "
                   f"{sum(len(b) for _, b in template_files) / 1024 / 1024:.2f} MB")
    else:
        st.caption("系統會自動辨識不同 Template 格式，不需要人工選擇格式。")


# ============================================================
# STEP 2 - Upload mapping files
# ============================================================
section_header("2", "Upload customer and SKU mapping files", "上傳 Mapping 資料")
map_col1, map_col2 = st.columns(2)
with map_col1:
    customer_file = st.file_uploader("Customer Mapping", type=["xlsx", "xls", "csv"], key="customer", on_change="ignore")
with map_col2:
    sku_files = st.file_uploader("SKU Mapping", type=["xlsx", "xls", "csv"], accept_multiple_files=True,
                                 key="sku", on_change="ignore")
    if sku_files:
        st.caption(f"已選擇 {len(sku_files)} 個 SKU Mapping 檔案")


# ============================================================
# STEP 3 - Processing options + Start / Reset
# ============================================================
section_header("3", "Process uploaded POS and mapping files", "開始處理資料")

with st.expander("⚙️ 處理選項 Options", expanded=False):
    o1, o2 = st.columns(2)
    with o1:
        duplicate_policy = st.selectbox(
            "重複提交（同月份 + 同店家編號）處理方式",
            ["flag", "keep_first", "keep_last"],
            format_func={"flag": "只標記，不排除 (flag)", "keep_first": "只保留第一個檔案 (keep_first)",
                         "keep_last": "只保留最後一個檔案 (keep_last)"}.get,
        )
        drop_nan_qty = st.checkbox("排除數量空白 / 無法解析的列", value=True)
    with o2:
        box_factor = st.number_input("「箱」換算成幾瓶（0 = 不換算，只標記 NEEDS_REVIEW）", min_value=0.0, value=0.0, step=1.0)
        qty_abs_max = st.number_input("數量合理上限（超過會標記 OUT_OF_RANGE）", min_value=1.0, value=100000.0, step=1000.0)

if "result" not in st.session_state:
    st.session_state["result"] = None
if "export_cache" not in st.session_state:
    st.session_state["export_cache"] = {}

col_start, col_reset = st.columns([4, 1])


def run_processing():
    cfg = core.PipelineConfig(
        duplicate_policy=duplicate_policy, drop_nan_qty=drop_nan_qty, qty_abs_max=qty_abs_max,
        unit_factors={"瓶": 1.0, "罐": 1.0, "支": 1.0, **({"箱": box_factor} if box_factor > 0 else {})},
        review_units=tuple(u for u in ("箱", "盒", "組", "包", "桶") if not (u == "箱" and box_factor > 0)),
    )
    progress = st.progress(0.0, text="準備中…")

    def on_progress(i, n, name):
        progress.progress(i / max(n, 1), text=f"處理中 {i}/{n}：{name}")

    try:
        result = core.run_transformation_pipeline(
            template_files,
            (customer_file.name, customer_file.getvalue()),
            [(f.name, f.getvalue()) for f in sku_files],
            customer_mapping_sheet="Customer data", sku_mapping_sheet="SKU data",
            config=cfg, progress_callback=on_progress,
        )
    except Exception as exc:  # mapping problems (missing columns etc.)
        progress.empty()
        st.error(f"處理失敗：{exc}")
        st.stop()
    progress.empty()
    st.session_state["result"] = result
    st.session_state["export_cache"] = {}


with col_start:
    if st.button("▶️ Start Processing data", type="primary", width="stretch"):
        if not template_files:
            st.error("請至少上傳一個當月份 Template。")
        elif customer_file is None:
            st.error("請上傳 Customer Mapping。")
        elif not sku_files:
            st.error("請至少上傳一個 SKU Mapping。")
        else:
            run_processing()  # runs ONCE; results are kept in session_state (v1 re-ran on every widget click)
with col_reset:
    if st.button("🔄 Reset", width="stretch"):
        st.session_state["result"] = None
        st.session_state["export_cache"] = {}
        st.rerun()

result = st.session_state["result"]
if result is None:
    st.info("Please make sure that all the mapping files have been uploaded before starting the process.")
    st.stop()

upload_log = result["upload_log"]

if result["final_report"] is None:
    st.error("所有 Template 都處理失敗。")
    show_df(upload_log)
    st.stop()

for w in result["warnings"]:
    st.warning(w)

detail = result["detail"]
final_report = result["final_report"]
trace = result["final_report_trace"]
store_rec = result["store_reconciliation"]

if not result["sku_upload_log"].empty:
    with st.expander("📦 SKU Mapping 上傳狀況", expanded=False):
        show_df(result["sku_upload_log"])


# ============================================================
# Dashboard KPIs
# ============================================================
source_count, output_count = len(result["raw_all"]), len(final_report)
customer_unmapped = int(detail["CUSTOMER_MAPPING_STATUS"].eq("UNMAPPED").sum())
sku_unmapped = int(detail["SKU_MAPPING_STATUS"].eq("UNMAPPED").sum())
sku_incomplete = int(detail["SKU_MAPPING_STATUS"].eq("INCOMPLETE").sum())
rsp_missing = int(detail["RSP_STATUS"].eq("MISSING").sum())
both_unmapped = int((detail["CUSTOMER_MAPPING_STATUS"].eq("UNMAPPED") & detail["SKU_MAPPING_STATUS"].eq("UNMAPPED")).sum())
unmapped_owner_files = (
    detail.loc[detail["CUSTOMER_OWNER_STATUS"].eq("UNMAPPED"), ["SOURCE_FILE", "FILE_OUTLET_NO", "FILE_RAWDATA_NAME"]]
    .drop_duplicates("SOURCE_FILE").reset_index(drop=True))
contract_mismatch = int(detail["CONTRACT_CHECK"].eq("MISMATCH").sum())
failed_files = int(upload_log["狀態"].eq("FAILED").sum())
successful_files = int(upload_log["狀態"].eq("SUCCESS").sum())

d1, d2, d3 = st.columns(3)
d4, d5, d6 = st.columns(3)
d1.metric("🐄 Source Count (來源資料)", f"{source_count:,}")
d2.metric("🧀 Output Count (輸出資料)", f"{output_count:,}", delta=f"{output_count - source_count:+,}")
d3.metric("🍪 Customer Unmapped", f"{customer_unmapped:,}")
d4.metric("🍾 SKU Unmapped", f"{sku_unmapped:,}")
d5.metric("✔️ Successful Files (成功檔案)", f"{successful_files:,}")
d6.metric("❌ Failed Files (失敗檔案)", f"{failed_files:,}")

# ---------------- store-level reconciliation
store_pass = int(store_rec["STATUS"].eq("PASS").sum())
store_check = int(store_rec["STATUS"].eq("CHECK").sum())
r1, r2, r3 = st.columns(3)
r1.metric("🏪 Total Store Count (核對店家)", f"{len(store_rec):,}")
r2.metric("✅ PASS", f"{store_pass:,}")
r3.metric("⚠️ CHECK", f"{store_check:,}")
if store_check == 0:
    st.success("✅ 所有店家銷量核對完成：來源資料與最終輸出總銷量一致。")
else:
    st.error(f"⚠️ 有 {store_check:,} 間店家來源與輸出銷量不一致。")
    show_df(store_rec[store_rec["STATUS"].eq("CHECK")])

# ---------------- additional quality metrics
q1, q2, q3 = st.columns(3)
q1.metric("Customer + SKU 同時 Unmapped", f"{both_unmapped:,}")
q2.metric("無法判定負責業務（檔案）", f"{len(unmapped_owner_files):,}")
q3.metric("合約編號不一致（Audit）", f"{contract_mismatch:,}")
if len(unmapped_owner_files):
    with st.expander("查看無法判定負責業務的檔案", expanded=False):
        show_df(unmapped_owner_files)

# ---------------- NEW v2 data-quality flags
st.markdown("#### 🧪 Data Quality Flags（v2 新增）")
n_dup_files = int(upload_log["重複提交"].eq("DUPLICATE").sum()) if "重複提交" in upload_log else 0
n_fallback = int(detail["CUSTOMER_MATCH_METHOD"].eq("RAWDATA_FALLBACK").sum())
n_unit = int(detail["UNIT_STATUS"].isin(["NEEDS_REVIEW", "UNKNOWN_UNIT"]).sum())
n_qtyflag = int(detail["QTY_FLAG"].ne("").sum())
n_warn_files = int(upload_log["警告"].astype(str).str.len().gt(0).sum())
f1, f2, f3, f4, f5, f6 = st.columns(6)
f1.metric("重複提交檔案", f"{n_dup_files:,}")
f2.metric("以店名配對(Fallback)", f"{n_fallback:,}")
f3.metric("單位待確認", f"{n_unit:,}")
f4.metric("數量異常(負/超大)", f"{n_qtyflag:,}")
f5.metric("SKU 不完整/無RSP", f"{sku_incomplete + rsp_missing:,}")
f6.metric("含警告的檔案", f"{n_warn_files:,}")
if n_warn_files:
    with st.expander("查看各檔案警告 Warnings", expanded=False):
        show_df(upload_log.loc[upload_log["警告"].astype(str).str.len().gt(0), ["檔案", "辨識格式", "警告"]])
if len(result["customer_duplicates"]) or len(result["sku_duplicates"]):
    with st.expander("Mapping 重複 / 衝突 Duplicates", expanded=False):
        if len(result["customer_duplicates"]):
            st.write("Customer Mapping（Outlet No / Rawdata Name 重複）")
            show_df(result["customer_duplicates"])
        if len(result["sku_duplicates"]):
            st.write("SKU Mapping（CONFLICT=True 表示重複項目的屬性不一致）")
            show_df(result["sku_duplicates"])


# ============================================================
# Monthly submission tracking
# ============================================================
st.markdown("---")
periods = core.available_periods(upload_log)
selected_period = ""
submission = None
if periods:
    selected_period = st.selectbox(
        "📅 選擇繳交月份", options=periods, index=0,
        format_func=lambda x: f"{x[:4]} / {x[4:]}", key="reporting_period_selector")
    submission = core.build_submission_tracking(upload_log, result["customer_map"], selected_period)

if submission is None:
    st.markdown("### 🔄 繳交進度")
    st.info("沒有可辨識的年月，無法計算繳交進度。")
else:
    st.markdown(f"### 🔄 {selected_period[:4]} / {selected_period[4:]} 繳交進度")
    s1, s2, s3, s4, s5 = st.columns(5)
    s1.metric("📋 本月應繳 (待確認FY27合約數!)", f"{submission['expected_count']:,}")
    s2.metric("✔️ Submission Successful (成功繳交)", f"{submission['success_count']:,}")
    s3.metric("❌ Submission Failed (上傳失敗)", f"{submission['failed_count']:,}")
    s4.metric("⏳ Missing Submission (尚未繳交)", f"{submission['missing_count']:,}")
    s5.metric("⚠️ Error (無法識別)", f"{submission['unmatched_count']:,}")

    st.write("**Overall Submission Progress**")
    st.progress(max(0.0, min(1.0, submission["completion"])))
    st.caption(f"{submission['success_count']:,} / {submission['expected_count']:,} ({submission['completion']:.1%})")

    st.markdown("####  👀業務繳交進度")
    sales = submission["sales_submission"]
    show_df(sales[["Sales ID", "Sales", "應繳", "成功", "失敗", "未繳", "完成率"]] if not sales.empty else sales)

    st.markdown("#### 🍩 店家繳交明細")
    sd = submission["submission_detail"]
    shop_detail = sd[["Sales ID", "Sales", "Contract JDE", "Contract NAME", "Outlet No", "Outlet NAME",
                      "Rawdata Name", "檔案", "辨識格式", "SUBMISSION_STATUS", "錯誤"]].copy()
    shop_detail.insert(0, "年月", selected_period)
    shop_detail = shop_detail.rename(columns={
        "Sales ID": "業務代號", "Sales": "業務員", "Contract JDE": "合約編號", "Contract NAME": "合約名稱",
        "Outlet No": "店家編號", "Outlet NAME": "店家名稱", "SUBMISSION_STATUS": "繳交狀態"})
    status_filter = st.multiselect("篩選繳交狀態", ["SUCCESS", "FAILED", "MISSING"],
                                   default=["SUCCESS", "FAILED", "MISSING"], key="submission_status_filter")
    show_df(shop_detail[shop_detail["繳交狀態"].isin(status_filter)], height=450)

    if submission["unmatched_count"] > 0:
        st.warning(f"⚠️ 有 {submission['unmatched_count']:,} 個上傳檔案無法與 Customer Mapping 對應。")
        with st.expander("查看無法識別的上傳檔案"):
            um = submission["unmatched_uploads"]
            show_df(um[[c for c in ["年月", "業務代號", "合約編號", "店家編號", "分店/Rawdata Name", "檔案",
                                    "辨識格式", "狀態", "錯誤", "UPLOAD_KEY"] if c in um.columns]])

st.markdown("#### 📁 檔案上傳狀況")
show_df(upload_log[[c for c in ["檔案", "辨識格式", "狀態", "警告", "原始筆數", "重複提交", "Detect 秒數",
                                "Transform 秒數", "總秒數", "業務代號", "合約編號", "年月", "店家編號",
                                "分店/Rawdata Name", "檔名格式", "錯誤"] if c in upload_log.columns]], height=400)
if failed_files:
    st.error(f"⚠️ 有 {failed_files:,} 個檔案處理失敗，請至「檔案上傳狀況」查看錯誤訊息。")

st.markdown("#### ⚠️ Data Quality Exceptions")
filename_error_count = int(upload_log["檔名格式"].ne("OK").sum())
e1, e2, e3, e4 = st.columns(4)
e1.metric("🍾 SKU Unmapped", f"{sku_unmapped:,}")
e2.metric("🍪 Customer Unmapped", f"{customer_unmapped:,}")
e3.metric("📄 檔名格式異常", f"{filename_error_count:,}")
e4.metric("⚠️ Submission Unmatched", f"{submission['unmatched_count'] if submission else 0:,}")


# ============================================================
# STEP 4 - Data preview
# ============================================================
section_header("4", "Data Preview", "Review final output, unmapped records and mapping exceptions")

PREVIEW_AUDIT_COLUMNS = core.existing_columns(detail, [
    "SOURCE_FILE", "SOURCE_SHEET", "SOURCE_ROW", "TEMPLATE_TYPE", "FILE_SALES_ID", "FILE_CONTRACT_JDE",
    "FILE_PERIOD", "FILE_OUTLET_NO", "FILE_RAWDATA_NAME", "RAW_ROW_CUSTOMER", "RAW_SKU", "QTY", "QTY_UNIT",
    "UNIT_STATUS", "QTY_FLAG", "CUSTOMER_MAPPING_STATUS", "CUSTOMER_MATCH_METHOD", "SKU_MAPPING_STATUS",
    "RSP_STATUS", "SALES_ID_CHECK", "CONTRACT_CHECK", "OUTLET_CHECK", "RAWDATA_NAME_CHECK"])


# NOTE: detail and final_report are sorted differently in v2, so they are no longer concatenated
# side by side (v1 did, which mis-aligned rows). Previews use the detail audit columns instead.
cust_un = detail["CUSTOMER_MAPPING_STATUS"].eq("UNMAPPED")
sku_un = detail["SKU_MAPPING_STATUS"].ne("MAPPED")

tabs = st.tabs(["最終報表", "Customer Unmapped", "SKU Unmapped", "全部 Unmapped", "數量/單位待確認", "上傳明細"])
with tabs[0]:
    show_df(final_report, height=550)
with tabs[1]:
    st.write(f"共 {int(cust_un.sum()):,} 筆")
    show_df(detail.loc[cust_un, PREVIEW_AUDIT_COLUMNS], height=520)
with tabs[2]:
    st.write(f"共 {int(sku_un.sum()):,} 筆（UNMAPPED + INCOMPLETE）")
    show_df(detail.loc[sku_un, PREVIEW_AUDIT_COLUMNS], height=520)
with tabs[3]:
    st.write(f"共 {int((cust_un | sku_un).sum()):,} 筆")
    show_df(detail.loc[cust_un | sku_un, PREVIEW_AUDIT_COLUMNS], height=520)
with tabs[4]:
    needs = detail["UNIT_STATUS"].isin(["NEEDS_REVIEW", "UNKNOWN_UNIT"]) | detail["QTY_FLAG"].ne("")
    st.write(f"共 {int(needs.sum()):,} 筆")
    show_df(detail.loc[needs, PREVIEW_AUDIT_COLUMNS], height=520)
with tabs[5]:
    show_df(upload_log, height=520)


# ============================================================
# STEP 5 - Export
# ============================================================
st.divider()
section_header("5", "Export", "Download final report, audit records and unmapped records")

export_period = selected_period or "ALL"
cache = st.session_state["export_cache"]
if export_period not in cache:  # build once per period (v1 rebuilt the workbook on every rerun)
    with st.spinner("正在產生匯出檔案…"):
        frames = core.build_export_frames(result, submission)
        cache[export_period] = {
            "frames": frames,
            "workbook": core.excel_bytes(frames["workbook_sheets"]),
            "unmapped": core.excel_bytes(frames["unmapped_workbook_sheets"]),
            "csv": final_report.to_csv(index=False).encode("utf-8-sig"),
        }
exp = cache[export_period]

x1, x2, x3 = st.columns(3)
x1.metric("Final Report", f"{len(final_report):,}")
x2.metric("Unmapped Rows", f"{len(exp['frames']['all_unmapped']):,}")
x3.metric("Owner Unmapped Files", f"{len(exp['frames']['owner_unmapped']):,}")

b1, b2, b3 = st.columns(3)
with b1:
    st.download_button("📥 下載完整 Excel", data=exp["workbook"],
                       file_name=f"POS_Monthly_Report_{export_period}.xlsx", mime=XLSX_MIME,
                       width="stretch", on_click="ignore")
with b2:
    st.download_button("📄 下載 Final CSV", data=exp["csv"],
                       file_name=f"POS_Monthly_Final_{export_period}.csv", mime="text/csv",
                       width="stretch", on_click="ignore")
with b3:
    st.download_button("⚠️ 下載 Unmapped Excel", data=exp["unmapped"],
                       file_name=f"POS_Unmapped_{export_period}.xlsx", mime=XLSX_MIME,
                       width="stretch", on_click="ignore")

with st.expander("🏪 All Store Reconciliation", expanded=False):
    show_df(store_rec)
