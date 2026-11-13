"""Detection + transformation smoke tests for the remaining templates."""
import io
import pandas as pd
import pos_core as core
from test_pos_core import xlsx_bytes, fake_src


def run(name, sheets, header=False):
    data = xlsx_bytes(sheets, header=header)
    src = core.SourceFile(name, data)
    code = core.detect_template_type(src)
    out, label = core.auto_transform_template(src, code)
    return code, out


def test_t4_wrapped_name_and_footer():
    raw = pd.DataFrame([
        ["麥金酒 期間商品銷售統計表", None, None],
        ["商品貨號", "商品名稱", "QTY"],
        ["T001", "黑牌12年700cc", 5],
        ["T020-1", "蘇格登13年貴腐酒桶(亞", 23],
        ["版)", None, None],
        ["第1頁", None, None],
        ["商品貨號", "商品名稱", "QTY"],
        ["T002", "紅牌", 2],
        ["以下空白", None, None],
    ])
    code, out = run("1_2_202609_4_麥金酒.xlsx", {"S": raw})
    assert code == "TEMPLATE_4"
    assert list(out["RAW_SKU"]) == ["黑牌12年700cc", "蘇格登13年貴腐酒桶(亞版)", "紅牌"]


def test_t5_header_on_row2():
    raw = pd.DataFrame([
        ["POS 公版", None, None, None, None],
        ["品項名稱 SKU Name", "品牌名稱 Brand Name", "規格 Size (ml)", "品類 Category", "銷量 Quantity"],
        ["黑牌", "JW", 700, "威士忌", 5],
    ])
    code, out = run("1_2_202609_4_康尼克.xlsx", {"POS公版": raw})
    assert code == "TEMPLATE_5" and out["SOURCE_ROW"].iloc[0] == 3
    assert out["SOURCE_SHEET"].iloc[0] == "POS公版" and out["RAW_BRAND"].iloc[0] == "JW"


def test_t7_unit_filter_for_special_customers():
    raw = pd.DataFrame([["品名", "數量", "單位"], ["黑牌", 5, "瓶"], ["紅牌", 2, "箱"], ["白牌", 1, "盒"]])
    code, out = run("1_2_202609_4_駱北.xlsx", {"S": raw})
    assert code == "TEMPLATE_7" and list(out["RAW_SKU"]) == ["黑牌", "白牌"]
    assert out.loc[out.RAW_SKU == "白牌", "UNIT_STATUS"].iloc[0] == "NEEDS_REVIEW"


def test_t8_row_level_units():
    raw = pd.DataFrame([
        ["黑牌", "12年", None, "瓶", None, 19, None, 13110],
        ["紅牌", None, None, None, "瓶", 1, None, 350],
        ["合計", None, None, "瓶", None, 20, None, 99999],
    ])
    code, out = run("1_2_202609_4_唯君系列.xlsx", {"S": raw})
    assert code == "TEMPLATE_8"
    assert list(out["RAW_SKU"]) == ["黑牌 12年", "紅牌"] and list(out["QTY"]) == [19, 1]


def test_t9_wins_over_t3a_for_multisheet():
    s1 = pd.DataFrame([["品號", "品名規格", "淨銷數量"], ["a", "黑牌", 3]])
    code, out = run("1_2_202609_4_新店李.xlsx", {"A": s1, "B": pd.DataFrame([[1]])})
    assert code == "TEMPLATE_9" and list(out["QTY"]) == [3]


def test_t10_header_and_total_rows_removed():
    raw = pd.DataFrame([["品名", "數量"], ["黑牌", 5], ["紅牌", 3], ["合計", 8]])
    code, out = run("1_2_202609_4_名酒集.xlsx", {"S": raw})
    assert code == "TEMPLATE_10" and list(out["RAW_SKU"]) == ["黑牌", "紅牌"]


def test_t11_t12_t15():
    raw = pd.DataFrame([["國際條碼", "品牌", "數量"], ["4713268580974", "黑牌", 5]])
    code, out = run("1_2_202609_4_酒江村.xlsx", {"S": raw})
    assert code == "TEMPLATE_11" and list(out["QTY"]) == [5]

    cols = ["品項名稱 SKU Name", "品牌名稱 Brand Name", "規格 Size (ml)", "品類 Category"] + [f"c{i}" for i in range(5)] + [""]
    rows = [cols, ["黑牌", "JW", 700, "W"] + [None] * 5 + [4], ["紅牌", "JW", 700, "W"] + [None] * 5 + [None]]
    code, out = run("1_2_202609_4_廣茂隆.xlsx", {"S": pd.DataFrame(rows)})
    assert code == "TEMPLATE_12" and list(out["RAW_SKU"]) == ["黑牌"] and list(out["QTY"]) == [4]

    a = pd.DataFrame([["產品名稱", "銷售量"], ["黑牌", 3]])
    b = pd.DataFrame([["產品名稱", "銷售量"], ["紅牌", 2]])
    code, out = run("1_2_202609_4_精湛酒藏.xlsx", {"A": a, "B": b})
    assert code == "TEMPLATE_15" and list(out["QTY"]) == [3, 2]


def test_t13_month_column():
    raw = pd.DataFrame([
        ["NO.", "in", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP"],
        ["A1", "黑牌", 1, 1, 1, 1, 1, 1, 1, 1, 9],
        ["Bootles", "x", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["TOTAL", "y", 0, 0, 0, 0, 0, 0, 0, 0, 99],
    ])
    code, out = run("1_2_202609_4_橡木桶.xlsx", {"S": raw})
    assert code == "TEMPLATE_13" and list(out["QTY"]) == [9]


def test_t6_and_t14_fixed_columns_with_injected_sheets():
    raw = pd.DataFrame([["2026/9/1", "C1", "客戶", "P1", "黑牌", 100, 5, 500]])
    out, _ = core.auto_transform_template(fake_src("1_2_202609_4_松賀.xls", {"S": raw}), "TEMPLATE_6")
    assert list(out["QTY"]) == [5] and out["RAW_UNIT_PRICE"].iloc[0] == 100


def test_xls_engine_error_is_clean():
    src = core.SourceFile("x.xls", b"garbage")
    try:
        src.sheets()
        assert False
    except ValueError as e:
        assert "無法讀取" in str(e)
