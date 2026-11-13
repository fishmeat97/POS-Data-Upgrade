"""
Tests for pos_core.py  -  run with:  python -m pytest -q test_pos_core.py
(or simply: python test_pos_core.py)

Each test maps to an item of the code review (A1..D7) - see CHANGE_RECORD.md.
.xls-only templates are tested by injecting the parsed sheet into SourceFile
so that no legacy Excel writer is required.
"""
import io

import numpy as np
import pandas as pd
import pytest

import pos_core as core


# ---------------------------------------------------------------- helpers
def xlsx_bytes(sheets, header=True):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False, header=header)
    return buf.getvalue()


def fake_src(name, sheets):
    """SourceFile whose sheets are injected (for .xls-only templates)."""
    s = core.SourceFile(name, b"")
    s._sheets = sheets
    return s


CUST = pd.DataFrame({
    "REGION": ["北", "南"], "Sales ID": ["2172", "2233"], "Sales": ["王", "李"],
    "CONTRACT TYPE": ["A", "B"], "Contract JDE": ["31032115", "31035675"],
    "Contract NAME": ["C1", "C2"], "Outlet No": ["31020549", "32014065"],
    "Outlet NAME": ["店一", "店二"], "Rawdata Name": ["加州三重店", "新店李系列"],
})
SKU = pd.DataFrame({
    "Mapping Name": ["黑牌12年", "紅牌"], "RSP": [1000, 500], "Price Band": ["H", "L"],
    "Manufacture": ["M", "M"], "Brand": ["JW", "JW"],
    "SKU Name": ["黑牌 12年 700cc", "紅牌700cc"], "SIZE": [700, 700], "CATEGORY": ["威士忌", "威士忌"],
})


# ---------------------------------------------------------------- numbers (A2, C2)
@pytest.mark.parametrize("value,expected", [
    ("26瓶", 26), ("2,408瓶", 2408), (12, 12), ("12.5", 12.5), ("(5)", -5), (" 7 ", 7),
])
def test_parse_number_ok(value, expected):
    assert core.parse_number(value) == expected


@pytest.mark.parametrize("value", [
    "4713268580974", "4713268580974.0", 4713268580974.0, "700cc", "2026-09", "約12", "", None, float("nan"),
])
def test_parse_number_rejects(value):
    assert np.isnan(core.parse_number(value))


def test_parse_unit():
    assert core.parse_unit("3箱") == "箱" and core.parse_unit("26瓶") == "瓶" and core.parse_unit(5) == ""


# ---------------------------------------------------------------- filename (C5)
def test_parse_filename_ok_and_roc():
    m = core.parse_filename("2233_31035675_202607_31020549_加州洋酒三重店.xlsx")
    assert (m["YEAR"], m["MONTH"], m["FILENAME_STATUS"]) == (2026, 7, "OK")
    m = core.parse_filename("2233_31035675_11509_31020549_店.xlsx")
    assert (m["YEAR"], m["MONTH"]) == (2026, 9)


def test_parse_filename_roc_sanity():
    assert core.parse_filename("1_2_20269_4_x.xlsx")["YEAR"] is None  # 202 -> invalid ROC year


def test_parse_filename_copy_suffix():
    assert core.parse_filename("1_2_202607_4_x(1).xlsx")["FILE_COPY_SUFFIX"] is True


# ---------------------------------------------------------------- column matching (C1)
def test_first_matching_column_keyword_priority_and_spaces():
    df = pd.DataFrame(columns=["訂購數量", "數 量", "x"])
    assert core.first_matching_column(df, ["數量"]) == "數 量"  # exact (norm_key) beats contains


# ---------------------------------------------------------------- T1 + policies
def t1_file(rows, name="2233_31035675_202607_31020549_加州三重店.xlsx"):
    df = pd.DataFrame(rows, columns=["品項名稱 SKU Name", "銷量 Quantity", "品牌名稱 Brand Name"])
    return (name, xlsx_bytes({"Sheet1": df}))


def test_t1_basic_and_real_source_rows_and_sheet_name():
    f = t1_file([["黑牌 12年 700cc", 5, "JW"], [None, None, None], ["紅牌700cc", "3瓶", "JW"]])
    data, label, code = core.transform_file(f)
    assert code == "TEMPLATE_1" or code == "TEMPLATE_7"
    assert list(data["QTY"]) == [5, 3]
    assert list(data["SOURCE_ROW"]) == [2, 4]  # B8: true Excel rows (blank row 3 skipped)
    assert (data["SOURCE_SHEET"] == "Sheet1").all()


def test_nan_qty_and_total_rows_dropped_with_note():  # B5 / B3
    f = t1_file([["黑牌 12年 700cc", 5, ""], ["合計", 99, ""], ["紅牌700cc", None, ""]])
    src = core.as_source_file(f)
    data, _ = core.auto_transform_template(src, core.detect_template_type(src))
    assert list(data["RAW_SKU"]) == ["黑牌 12年 700cc"]
    assert any("排除" in n for n in data.attrs["notes"])


def test_unit_flagged_not_converted_and_factor_applied():  # A4
    f = t1_file([["黑牌 12年 700cc", "2箱", ""], ["紅牌700cc", "3瓶", ""]])
    data, _, _ = core.transform_file(f)
    assert dict(zip(data["RAW_SKU"], data["UNIT_STATUS"])) == {
        "黑牌 12年 700cc": "NEEDS_REVIEW", "紅牌700cc": "OK"}
    assert data.loc[0, "QTY"] == 2  # not silently converted
    cfg = core.PipelineConfig(unit_factors={"瓶": 1, "箱": 6})
    data, _, _ = core.transform_file(f, cfg=cfg)
    assert data.loc[0, "QTY"] == 12 and data.loc[0, "UNIT_STATUS"] == "CONVERTED"


def test_qty_flag_out_of_range_and_negative():  # A2
    f = t1_file([["黑牌 12年 700cc", 200000, ""], ["紅牌700cc", -3, ""]])
    data, _, _ = core.transform_file(f)
    assert dict(zip(data["RAW_SKU"], data["QTY_FLAG"])) == {
        "黑牌 12年 700cc": "OUT_OF_RANGE", "紅牌700cc": "NEGATIVE"}


# ---------------------------------------------------------------- A1 : T14
def test_t14_keeps_names_starting_with_D():
    raw = pd.DataFrame([
        ["", "", "", "D001-00-070-001", "", 9],
        ["", "", "", "Dalmore 12年", "", 4],
        ["", "", "", "Dewar's 12年", "", 7],
        ["", "", "", "黑牌", "", 2],
    ])
    src = fake_src("2233_31035675_202607_31020549_美滿商行.xls", {"Sheet1": raw})
    data, _ = core.auto_transform_template(src, "TEMPLATE_14")
    assert set(data["RAW_SKU"]) == {"Dalmore 12年", "Dewar's 12年", "黑牌"}
    assert any("D 開頭商品代碼" in n for n in data.attrs["notes"])
    assert list(data["SOURCE_ROW"]) == [2, 3, 4]


# ---------------------------------------------------------------- A5 : T2 fallback warning
def t2_raw(month_label):
    return pd.DataFrame([
        [None, None, month_label, None],
        ["商品代號", "商品名稱", "實銷", "庫存"],
        ["a", "黑牌", 5, 1], ["b", "紅牌", 3, 1], ["c", "合計", 8, 2],
    ])


def test_t2_matching_month_no_warning_and_fallback_warns():
    ok = fake_src("1_2_202609_4_洋酒城.xlsx", {"S": t2_raw("2026-09-01")})
    data, _ = core.auto_transform_template(ok, "TEMPLATE_2")
    assert list(data["RAW_SKU"]) == ["黑牌", "紅牌"]  # 合計 removed
    assert not any("⚠" in n for n in data.attrs["notes"])
    bad = fake_src("1_2_202609_4_洋酒城.xlsx", {"S": t2_raw("2026-08-01")})
    data, _ = core.auto_transform_template(bad, "TEMPLATE_2")
    assert any("⚠" in n for n in data.attrs["notes"])


# ---------------------------------------------------------------- A6 : T3 fallback only on TemplateError
def test_t3_falls_back_to_3b_on_template_error():
    sheet1 = pd.DataFrame([["貨品名稱", "數 量"], ["黑牌", 3]])
    sheet2 = pd.DataFrame([["x", "y"], [1, 2]])
    src = fake_src("1_2_202609_4_國泰.xlsx", {"A": sheet1, "B": sheet2})
    data, label = core.auto_transform_template(src, "TEMPLATE_3")
    assert "3B" in label and list(data["RAW_SKU"]) == ["黑牌"]


# ---------------------------------------------------------------- A7 : summary sheets
def test_summary_sheet_skipped():
    a = pd.DataFrame([["品名規格", "數量"], ["黑牌", 3]])
    s = pd.DataFrame([["品名規格", "數量"], ["黑牌", 3]])
    src = fake_src("1_2_202609_4_國泰.xlsx", {"明細": a, "彙總": s})
    data, label = core.auto_transform_template(src, "TEMPLATE_3")
    assert len(data) == 1 and any("彙總" in n for n in data.attrs["notes"])


# ---------------------------------------------------------------- B4 : T16 store column
def test_t16_uses_store_column_not_neighbour():
    raw = pd.DataFrame([
        ["產品編號", "產品品名", "東門店", "合歡"],
        ["F1", "黑牌", 1, 10],
        ["F2", "紅牌", 2, 20],
    ])
    src = fake_src("1_2_202609_4_合歡.xlsx", {"S": raw})
    data, _ = core.auto_transform_template(src, "TEMPLATE_16")
    assert list(data["QTY"]) == [10, 20] and data["SOURCE_QTY_COLUMN"].iloc[0] == "合歡"


def test_t16_fallback_warns():
    raw = pd.DataFrame([["產品編號", "產品品名", "東門店"], ["F1", "黑牌", 1]])
    src = fake_src("1_2_202609_4_合歡.xlsx", {"S": raw})
    data, _ = core.auto_transform_template(src, "TEMPLATE_16")
    assert any("⚠" in n for n in data.attrs["notes"])


# ---------------------------------------------------------------- B1/B2 : CSV detect + transform consistent
def test_csv_t1_with_銷售量_header_detected_and_transformed():
    csv = "品名,銷售量\n黑牌 12年 700cc,5\n紅牌700cc,3\n".encode("utf-8-sig")
    f = ("2233_31035675_202607_31020549_x.csv", csv)
    data, _, code = core.transform_file(f)
    assert list(data["QTY"]) == [5, 3]


def test_csv_cp950_and_ragged_rows():
    text = "品名,數量\n黑牌,5\n紅牌,3,多出來的欄位\n"
    f = ("2233_31035675_202607_31020549_x.csv", text.encode("cp950"))
    data, _, _ = core.transform_file(f)
    assert list(data["QTY"]) == [5, 3]


# ---------------------------------------------------------------- B12 : file-like objects
def test_source_from_file_like(tmp_path):
    p = tmp_path / "2233_31035675_202607_31020549_x.csv"
    p.write_bytes("品名,數量\n黑牌,5\n".encode("utf-8"))
    with open(p, "rb") as fh:  # plain file object: has .name but no .getvalue()
        data, _, _ = core.transform_file(fh)
    assert list(data["QTY"]) == [5]


# ---------------------------------------------------------------- pipeline
def pipeline(files, **kw):
    return core.run_transformation_pipeline(
        files,
        ("Customer.xlsx", xlsx_bytes({"Customer data": CUST})),
        ("SKU.xlsx", xlsx_bytes({"SKU data": SKU})),
        customer_mapping_sheet="Customer data", sku_mapping_sheet="SKU data", **kw)


def test_pipeline_end_to_end_and_report_columns():
    res = pipeline([t1_file([["黑牌 12年 700cc", 5, "JW"], ["未知品", 2, "x"]])])
    rep = res["final_report"]
    assert list(rep.columns) == core.OUTPUT_COLUMNS
    row = rep[rep["品項"] == "黑牌 12年 700cc"].iloc[0]
    assert row["單價"] == 1000 and row["總價"] == 5000 and row["店家名稱"] == "店一"
    assert rep.set_index("品項").loc["未知品", "SKU Mapping Status"] == "UNMAPPED"
    assert (res["store_reconciliation"]["STATUS"] == "PASS").all()


def test_pipeline_all_failed_returns_log_not_exception():  # D5
    res = pipeline([("bad.xlsx", b"not an excel")])
    assert res["final_report"] is None and res["upload_log"].loc[0, "狀態"] == "FAILED"
    with pytest.raises(ValueError):
        pipeline([("bad.xlsx", b"not an excel")], raise_on_all_failed=True)


def test_pipeline_one_bad_file_does_not_stop_batch():
    res = pipeline([("bad.xlsx", b"zzz"), t1_file([["黑牌 12年 700cc", 5, "JW"]])])
    assert sorted(res["upload_log"]["狀態"]) == ["FAILED", "SUCCESS"]
    assert len(res["final_report"]) == 1


def test_duplicate_submission_flag_and_policy():  # A3
    f1 = t1_file([["黑牌 12年 700cc", 5, ""]])
    f2 = ("2233_31035675_202607_31020549_加州三重店(1).xlsx", t1_file([["黑牌 12年 700cc", 7, ""]])[1])
    res = pipeline([f1, f2])
    assert (res["upload_log"]["重複提交"] == "DUPLICATE").all()
    assert res["final_report"]["銷量(瓶)"].sum() == 12  # flag only
    res = pipeline([f1, f2], config=core.PipelineConfig(duplicate_policy="keep_last"))
    assert res["final_report"]["銷量(瓶)"].sum() == 7


# ---------------------------------------------------------------- A8 : customer fallback
def test_customer_fallback_is_flagged_and_report_store_consistent():
    name = "2233_99999999_202607_99999999_加州三重店.xlsx"  # outlet unknown, Rawdata Name known
    res = pipeline([t1_file([["黑牌 12年 700cc", 5, ""]], name=name)])
    d = res["detail"]
    assert d.loc[0, "CUSTOMER_MATCH_METHOD"] == "RAWDATA_FALLBACK"
    assert d.loc[0, "OUTLET_CHECK"] == "MISMATCH"
    rep = res["final_report"]
    assert rep.loc[0, "店家編號"] == "31020549" and rep.loc[0, "店家名稱"] == "店一"
    assert (res["store_reconciliation"]["STATUS"] == "PASS").all()


def test_ambiguous_rawdata_not_used_for_fallback():
    cust = pd.concat([CUST, CUST.iloc[[0]].assign(**{"Outlet No": "55555555"})], ignore_index=True)
    cmap, dups = core.prepare_customer_mapping(cust)
    assert set(dups["DUP_TYPE"]) == {"RAWDATA_NAME"}
    raw = pd.DataFrame({"FILE_OUTLET_NO": ["000"], "FILE_RAWDATA_NAME": ["加州三重店"],
                        "FILE_SALES_ID": ["x"], "FILE_CONTRACT_JDE": ["y"]})
    out = core.attach_customer_mapping(raw, cmap)
    assert out.loc[0, "CUSTOMER_MAPPING_STATUS"] == "UNMAPPED"
    out = core.attach_customer_mapping(raw, cmap, core.PipelineConfig(ambiguous_rawdata_fallback="first"))
    assert out.loc[0, "CUSTOMER_MAPPING_STATUS"] == "MAPPED"


# ---------------------------------------------------------------- A9 : SKU status
def test_sku_incomplete_and_rsp_missing_and_conflict():
    sku = pd.concat([SKU, pd.DataFrame({
        "Mapping Name": [""], "RSP": [None], "Price Band": ["x"], "Manufacture": ["x"],
        "Brand": ["x"], "SKU Name": ["沒名字品"], "SIZE": [1], "CATEGORY": ["c"]}),
        SKU.iloc[[0]].assign(RSP=999)], ignore_index=True)
    smap, dup, _ = core.load_sku_mapping(("SKU.xlsx", xlsx_bytes({"SKU data": sku})), "SKU data")
    assert bool(dup["CONFLICT"].all())
    raw = pd.DataFrame({"RAW_SKU": ["黑牌 12年 700cc", "沒名字品", "不存在"]})
    out = core.attach_sku_mapping(raw, smap)
    assert list(out["SKU_MAPPING_STATUS"]) == ["MAPPED", "INCOMPLETE", "UNMAPPED"]
    assert list(out["RSP_STATUS"]) == ["OK", "MISSING", "UNMAPPED"]


# ---------------------------------------------------------------- submission + export
def test_submission_tracking_and_export_workbook():
    res = pipeline([t1_file([["黑牌 12年 700cc", 5, ""]])])
    period = core.available_periods(res["upload_log"])[0]
    assert period == "202607"
    sub = core.build_submission_tracking(res["upload_log"], res["customer_map"], period)
    assert (sub["expected_count"], sub["success_count"], sub["missing_count"]) == (2, 1, 1)
    frames = core.build_export_frames(res, sub)
    data = core.excel_bytes(frames["workbook_sheets"])
    xl = pd.ExcelFile(io.BytesIO(data))
    assert "Final Report" in xl.sheet_names and "Submission Tracker" in xl.sheet_names


def test_excel_illegal_characters_removed():
    out = core.excel_bytes({"a:b": pd.DataFrame({"x": ["ok\x01bad"]})})
    assert pd.read_excel(io.BytesIO(out)).loc[0, "x"] == "okbad"


def test_zip_extract_ignores_junk():
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("a/1_2_202607_3_x.csv", "a,b\n1,2")
        z.writestr("__MACOSX/._x.csv", "junk")
        z.writestr("readme.txt", "hi")
    files, ignored = core.extract_pos_zip(buf.getvalue())
    assert [n for n, _ in files] == ["1_2_202607_3_x.csv"] and ignored == ["readme.txt"]


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
