import pandas as pd
import pytest

from buybox import labels as L
from buybox.core import Params, market_summary
from buybox.io import InputFileError, detect_columns, is_csv, read_kam_list, read_table_bytes, to_number
from buybox.pipeline import analyze

HEADER = "DKPC,DKP,Seller_ID,Weight,Selling_Price,Seller_Name\n"


def _analyze_csv(data: bytes, name: str = "a.csv"):
    raw = read_table_bytes(data, name)
    return analyze(raw, detect_columns(list(raw.columns)).mapping, None, Params(), name)


def test_csv_utf8_bom_with_thousands_and_persian_digits():
    data = (HEADER + '1,10,5,1,"1,000,000",الف\n2,10,6,1,۱۲۰۰۰۰۰,ب\n3,10,7,,900,ج\n').encode("utf-8-sig")
    raw = read_table_bytes(data, "a.csv")
    assert list(raw.columns) == HEADER.strip().split(",")
    assert raw["Selling_Price"].tolist()[:2] == [1_000_000, 1_200_000]
    a = _analyze_csv(data)
    r = a.rows.set_index("DKPC")
    assert r.loc[1, "Comment"] == L.COMMENT_WIN and r.loc[2, "Comment"] == L.COMMENT_LOSS
    assert r.loc[3, "Comment"] == L.COMMENT_NO_WEIGHT


def test_csv_semicolon_cp1256_decimal_comma():
    text = "DKPC;DKP;Seller ID;Weight;Selling Price;Seller Name\n1;10;5;0,5;100;فروشنده\n2;10;6;0,5;90;رقيب\n"
    raw = read_table_bytes(text.encode("cp1256"), "b.csv")
    assert raw["Weight"].tolist() == [0.5, 0.5]
    assert raw["Seller Name"].tolist() == ["فروشنده", "رقيب"]
    a = analyze(raw, detect_columns(list(raw.columns)).mapping, None, Params(), "b.csv")
    assert market_summary(a.rows)["comp_rows"] == 2


def test_csv_tab_delimited_and_trailing_delimiter():
    data = "DKPC\tDKP\tSeller_ID\tWeight\tSelling_Price\t\n1\t10\t5\t1\t100\t\n".encode()
    raw = read_table_bytes(data, "c.csv")
    assert list(raw.columns) == ["DKPC", "DKP", "Seller_ID", "Weight", "Selling_Price"]


def test_csv_gives_same_result_as_xlsx(tmp_path):
    df = pd.DataFrame({"DKPC": [1, 2, 3], "DKP": [7, 7, 8], "Seller_ID": [1, 2, 2], "Weight": [0.5, 0.5, 1.0],
                       "Selling_Price": [100, 101, 50]})
    xlsx, csv = tmp_path / "a.xlsx", tmp_path / "a.csv"
    df.to_excel(xlsx, index=False)
    df.to_csv(csv, index=False)
    a = _analyze_csv(csv.read_bytes(), "a.csv")
    raw = read_table_bytes(xlsx.read_bytes(), "a.xlsx")
    b = analyze(raw, detect_columns(list(raw.columns)).mapping, None, Params(), "a.xlsx")
    assert a.rows["Comment"].tolist() == b.rows["Comment"].tolist()
    assert a.rows["Near_Miss"].tolist() == b.rows["Near_Miss"].tolist()


def test_csv_errors():
    for bad in (b"", b"\n\n"):
        with pytest.raises(InputFileError, match="خالی"):
            read_table_bytes(bad, "x.csv")
    with pytest.raises(InputFileError, match="CSV"):
        read_table_bytes(b"a,b\n1,2\n3,4,5\n", "x.csv")


def test_kam_list_csv():
    data = "Seller Id,Seller Name,KAM\n5,X,علي\n5,X2,مريم\n".encode("cp1256")
    kam, warnings = read_kam_list(data, "kam.csv")
    assert warnings == [] and kam["Seller_ID"].tolist() == [5, 5]


def test_is_csv_detection():
    assert is_csv(b"a,b", "x.CSV")
    assert not is_csv(b"PK\x03\x04", "x.xlsx")
    assert is_csv(b"a,b\n1,2", "")          # no name: not a zip -> CSV
    assert not is_csv(b"PK\x03\x04", "")


def test_to_number_variants():
    s = pd.Series(["1,000,000", "۱۲۰۰", "0,5", "2.5", " 3 ", "", None, "abc", "۰٫۵", "1,234.5"])
    out = to_number(s).tolist()
    assert out[:5] == [1_000_000, 1200, 0.5, 2.5, 3]
    assert all(v != v for v in out[5:8])  # NaN
    assert out[8:] == [0.5, 1234.5]
