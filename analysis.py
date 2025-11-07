import pandas as pd

# Load file
df = pd.read_excel(
    "CTF-S transport related (see notes).xlsx",
    sheet_name="Sheet1"
)


# Ensure normalized value exists — fallback to raw Value if missing
if "Value (USD)_normalized" not in df.columns:
    df["Value (USD)_normalized"] = df["Value (USD)"]

# Identify individual donors/recipients
df["is_indiv_donor"] = df["funding_economy_code"].notna()
df["is_indiv_recipient"] = df["recipient_economy_code"].notna()

years = [2021, 2022, "both"]

def format_money(x):
    return f"${x:,.2f}"

for year in years:
    if year == "both":
        sub = df.copy()
        year_label = "2021 + 2022"
    else:
        sub = df[df["Year"] == year]
        year_label = str(year)

    print("\n" + "="*80)
    print(f"TRANSPORT FINANCE SUMMARY — YEAR: {year_label}")
    print("="*80)

    # Total
    total_amount = sub["Value (USD)_normalized"].sum()
    print(f"Total Transport Finance: {format_money(total_amount)}\n")

    # ------------------------ TOP DONOR (INDIVIDUAL) ------------------------
    donor_indiv = (
        sub[sub["is_indiv_donor"]]
        .groupby("Funding economy")["Value (USD)_normalized"]
        .sum()
        .sort_values(ascending=False)
    ).head(3)

    print("Top 3 Donors (Individual Economies Only):")
    if donor_indiv.empty:
        print("  No individual donors found.")
    else:
        for name, val in donor_indiv.items():
            print(f"  {name:<20} {format_money(val)}")
    print()

    # ------------------------ TOP DONOR (ALL ENTITIES) ------------------------
    donor_all = (
        sub.groupby("Funding economy")["Value (USD)_normalized"]
        .sum()
        .sort_values(ascending=False)
    ).head(3)

    print("Top 3 Donors (All Entities):")
    for name, val in donor_all.items():
        print(f"  {name:<20} {format_money(val)}")
    print()

    # ------------------------ TOP RECIPIENT (INDIVIDUAL) ------------------------
    recipient_indiv = (
        sub[sub["is_indiv_recipient"]]
        .groupby("Recipient country or region")["Value (USD)_normalized"]
        .sum()
        .sort_values(ascending=False)
    ).head(3)

    print("Top 3 Recipients (Individual Economies Only):")
    if recipient_indiv.empty:
        print("  No individual recipients found.")
    else:
        for name, val in recipient_indiv.items():
            print(f"  {name:<25} {format_money(val)}")
    print()

    # ------------------------ TOP RECIPIENT (ALL ENTITIES) ------------------------
    recipient_all = (
        sub.groupby("Recipient country or region")["Value (USD)_normalized"]
        .sum()
        .sort_values(ascending=False)
    ).head(3)

    print("Top 3 Recipients (All Entities):")
    for name, val in recipient_all.items():
        print(f"  {name:<25} {format_money(val)}")
    print("\n")