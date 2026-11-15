"""Prepare monthly sales inputs for the Excel sales performance tracker.

Source: "Online Retail II" (UCI Machine Learning Repository, CC BY 4.0).
Real transactions of a UK-based online gift retailer, 1 Dec 2009 - 9 Dec 2011.

Usage:
    python prepare_data.py <online_retail_II.xlsx or raw.pkl> <output_dir>

Outputs (CSV):
    sales_input.csv      Month, Market, ProductLine, CustomerType, NetSales, Units
    customers_input.csv  Month, Market, CustomerType, Customers, Orders
    control_totals.csv   Month, NetSales, Units, Lines (checks for the workbook)
    cleaning_log.csv     what was removed and why
    country_market.csv   which market each country belongs to
    product_rules.csv    keyword rules that assign product lines
"""
import re
import sys
from pathlib import Path

import pandas as pd

FIRST_MONTH = "2009-12"   # first full month in the source
LAST_MONTH = "2011-11"    # last full month (December 2011 only has 9 days)

# Stock codes that are not products (postage, fees, manual adjustments, tests, vouchers).
NON_PRODUCT_CODES = {
    "POST", "DOT", "M", "m", "C2", "D", "S", "BANK CHARGES", "ADJUST", "ADJUST2",
    "AMAZONFEE", "PADS", "CRUK", "TEST001", "TEST002", "B", "DCGSSGIRL", "DCGSSBOY",
}

EU_COUNTRIES = {  # EU member states in 2010-2011, as named in the source
    "Austria", "Belgium", "Bulgaria", "Cyprus", "Czech Republic", "Denmark", "EIRE",
    "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Italy", "Latvia",
    "Lithuania", "Luxembourg", "Malta", "Netherlands", "Poland", "Portugal", "Romania",
    "Slovakia", "Slovenia", "Spain", "Sweden", "European Community",
}

# Product lines from keywords in the product description (first match wins).
PRODUCT_RULES = [
    ("Christmas & seasonal", r"CHRISTMAS|XMAS|ADVENT|SANTA|REINDEER|EASTER|HALLOWEEN|SNOWFLAKE|NATIVITY|BAUBLE"),
    ("Gifts & accessories", r"HOT WATER BOTTLE|HAND WARMER|HOTTIE|PARASOL|UMBRELLA|\bFAN\b|PURSE|NECKLACE|BRACELET|EARRING|KEY ?RING|\bHAIR|BROOCH|SCARF|GLOVES|WALLET|PASSPORT COVER|LUGGAGE TAG"),
    ("Candles & lighting", r"T-LIGHT|TLIGHT|T LIGHT|CANDLE|LANTERN|NIGHT LIGHT|\bLIGHTS?\b|\bLAMP|\bLED\b|INCENSE|OIL BURNER|FAIRY LIGHT"),
    ("Party & gift wrap", r"PARTY|BUNTING|BIRTHDAY|PAPER CHAIN|BALLOON|RIBBON|\bWRAP|GIFT TAG|\bCARDS?\b(?! HOLDER)|GREETING|CONFETTI|TISSUE|PAPER PLATE|PAPER CUP|NAPKIN|INVITES|BANNER"),
    ("Kitchen & dining", r"CAKE|\bTEA\b|TEACUP|TEAPOT|TEA SET|\bMUGS?\b|\bCUPS?\b|SAUCER|PLATE|\bBOWLS?\b|\bJUG\b|COFFEE|BAKING|KITCHEN|LUNCH BOX|SNACK BOX|\bJARS?\b|\bJAM\b|RECIPE|POPCORN|EGG|CUTLERY|SPOON|APRON|OVEN|COASTER|\bTRAY\b|BISCUIT|SCALES|TOAST|BREAD|PANTRY|DRINK|GLASS(ES)?\b|BOTTLE|PICNIC|SUGAR|MILK"),
    ("Bags & storage", r"\bBAGS?\b|SHOPPER|BASKET|STORAGE|HAMPER|CRATE|TRUNK|SUITCASE|TOTE"),
    ("Toys, crafts & stationery", r"\bTOYS?\b|\bGAMES?\b|\bDOLL|PUZZLE|HARMONICA|BLOCKS\b|SPINNING TOP|SKIPPING|KIDS|CHILDS|CHILDRENS|CRAFT|PAINT SET|CRAYON|COLOURING|PENCIL|\bPENS?\b|NOTEBOOK|NOTE BOOK|STICKER|STATIONERY|SKETCHBOOK|ERASER|RUBBER|\bCHALK\b|STAMP|SEWING|KNITTING|\bFELT\b|MAGNET|JIGSAW|TEDDY|SOLDIER"),
    ("Home decor", r"SIGN|FRAME|CLOCK|DOORMAT|CUSHION|HEART|DECORATION|ORNAMENT|WICKER|HOOK|CHALKBOARD|BLACKBOARD|BLACK BOARD|MEMOBOARD|BLOCK WORD|CABINET|DRAWER|CHEST|COAT RACK|TRINKET|VASE|MIRROR|GARLAND|WREATH|CHIME|HANGING|PHOTO|WALL|DOORSTOP|HANGER|\bBIRD|SHELF|PLAQUE|DOOR|BELL|STAR\b|FLOWER|GARDEN|PLANT|POT\b"),
]
OTHER_LINE = "Other"
PRODUCT_LINES = [name for name, _ in PRODUCT_RULES] + [OTHER_LINE]
MARKETS = ["United Kingdom", "EU", "Rest of world"]
CUSTOMER_TYPES = ["New", "Returning", "Guest (no ID)"]


def product_line(description):
    text = str(description).upper()
    for name, pattern in PRODUCT_RULES:
        if re.search(pattern, text):
            return name
    return OTHER_LINE


def market(country):
    if country == "United Kingdom":
        return "United Kingdom"
    if country in EU_COUNTRIES:
        return "EU"
    return "Rest of world"


def load(path):
    path = Path(path)
    if path.suffix == ".pkl":
        return pd.read_pickle(path)
    sheets = pd.read_excel(path, sheet_name=None, dtype={"Invoice": str, "StockCode": str})
    return pd.concat(sheets.values(), ignore_index=True)


def main(src, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    raw = load(src)
    log = [("Rows in source (both sheets)", len(raw), "")]

    df = raw.drop_duplicates()
    log.append(("Exact duplicate rows removed", len(raw) - len(df),
                "includes the 1-9 Dec 2010 rows that appear in both sheets of the source"))

    code = df["StockCode"].astype(str)
    is_non_product = code.isin(NON_PRODUCT_CODES) | code.str.startswith("gift_0001")
    log.append(("Non-product lines removed", int(is_non_product.sum()),
                "postage, carriage, fees, manual adjustments, discounts, samples, tests, gift vouchers"))
    df = df[~is_non_product]

    zero_price = df["Price"] <= 0
    log.append(("Lines with price 0 or below removed", int(zero_price.sum()),
                "stock adjustments with no sales value"))
    df = df[~zero_price].copy()

    df["Month"] = df["InvoiceDate"].dt.to_period("M")
    in_range = (df["Month"] >= pd.Period(FIRST_MONTH)) & (df["Month"] <= pd.Period(LAST_MONTH))
    log.append(("Lines outside full months removed", int((~in_range).sum()),
                "1-9 Dec 2011 (partial month)"))
    df = df[in_range].copy()

    df["NetValue"] = df["Quantity"] * df["Price"]
    df["IsReturn"] = df["Invoice"].astype(str).str.startswith("C")
    df["Market"] = df["Country"].map(market)
    descriptions = df["Description"].fillna("").astype(str)
    line_lookup = {d: product_line(d) for d in descriptions.unique()}
    df["ProductLine"] = descriptions.map(line_lookup)

    # New = first purchase (not return) of that customer in the source data.
    cust = df["Customer ID"]
    first_buy = (df[~df["IsReturn"] & cust.notna()]
                 .groupby("Customer ID")["Month"].min())
    df["FirstMonth"] = cust.map(first_buy)
    df["CustomerType"] = "Returning"
    df.loc[cust.isna(), "CustomerType"] = "Guest (no ID)"
    df.loc[cust.notna() & (df["Month"] == df["FirstMonth"]), "CustomerType"] = "New"
    log.append(("Lines kept", len(df), "net sales = quantity x price; returns (invoices starting with C) are negative"))

    df["MonthStart"] = df["Month"].dt.to_timestamp()

    sales = (df.groupby(["MonthStart", "Market", "ProductLine", "CustomerType"], as_index=False)
               .agg(NetSales=("NetValue", "sum"), Units=("Quantity", "sum")))
    sales["NetSales"] = sales["NetSales"].round(2)
    order_types = {name: i for i, name in enumerate(MARKETS)}
    sales["_m"] = sales["Market"].map(order_types)
    sales["_p"] = sales["ProductLine"].map({n: i for i, n in enumerate(PRODUCT_LINES)})
    sales["_c"] = sales["CustomerType"].map({n: i for i, n in enumerate(CUSTOMER_TYPES)})
    sales = (sales.sort_values(["MonthStart", "_m", "_p", "_c"])
                  .drop(columns=["_m", "_p", "_c"])
                  .rename(columns={"MonthStart": "Month"}))

    buys = df[~df["IsReturn"]]
    customers = (buys.groupby(["MonthStart", "Market", "CustomerType"], as_index=False)
                     .agg(Customers=("Customer ID", "nunique"), Orders=("Invoice", "nunique")))
    customers.loc[customers["CustomerType"] == "Guest (no ID)", "Customers"] = 0
    customers["_m"] = customers["Market"].map(order_types)
    customers["_c"] = customers["CustomerType"].map({n: i for i, n in enumerate(CUSTOMER_TYPES)})
    customers = (customers.sort_values(["MonthStart", "_m", "_c"])
                          .drop(columns=["_m", "_c"])
                          .rename(columns={"MonthStart": "Month"}))

    control = (df.groupby("MonthStart", as_index=False)
                 .agg(NetSales=("NetValue", "sum"), Units=("Quantity", "sum"), Lines=("Invoice", "size"),
                      DistinctCustomers=("Customer ID", "nunique"))
                 .rename(columns={"MonthStart": "Month"}))
    control["NetSales"] = control["NetSales"].round(2)

    sales.to_csv(out / "sales_input.csv", index=False, date_format="%Y-%m-%d")
    customers.to_csv(out / "customers_input.csv", index=False, date_format="%Y-%m-%d")
    control.to_csv(out / "control_totals.csv", index=False, date_format="%Y-%m-%d")
    pd.DataFrame(log, columns=["Step", "Rows", "Note"]).to_csv(out / "cleaning_log.csv", index=False)

    countries = (df[["Country", "Market"]].drop_duplicates()
                   .assign(_m=lambda x: x["Market"].map(order_types))
                   .sort_values(["_m", "Country"]).drop(columns="_m"))
    countries.to_csv(out / "country_market.csv", index=False)
    rules = [(i + 1, name, pattern) for i, (name, pattern) in enumerate(PRODUCT_RULES)]
    rules.append((len(rules) + 1, OTHER_LINE, "(no rule above matched)"))
    pd.DataFrame(rules, columns=["Order", "ProductLine", "Keywords"]).to_csv(out / "product_rules.csv", index=False)

    multi_market = df[cust.notna()].groupby("Customer ID")["Market"].nunique()
    other_share = df.loc[df["ProductLine"] == OTHER_LINE, "NetValue"].sum() / df["NetValue"].sum()
    print("rows:", len(df), "| sales rows:", len(sales), "| customer rows:", len(customers))
    print("customers in more than one market:", int((multi_market > 1).sum()))
    print(f"net sales in 'Other' product line: {other_share:.1%}")
    return df, sales, customers, control


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
