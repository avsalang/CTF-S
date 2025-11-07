#!/usr/bin/env python3
"""
Enhanced Sankey diagram for transport-related funding (individual economies only)

- Fix: use rgba() strings for link transparency (Plotly doesn't accept 8-digit hex)
- Colors: separate palettes for donors vs recipients
- Styling: dark mode toggle, thicker nodes, spacing, title
- Edit parameters in USER SETTINGS
"""

import pandas as pd
import plotly.graph_objects as go
from pathlib import Path

# ============================================================
# ✅ USER SETTINGS — EDIT HERE
# ============================================================
INPUT_FILE = r"CTF-S transport related (see notes).xlsx"
OUTPUT_DIR = r"."

YEAR = "both"          # "2021", "2022", or "both"
TOP_EDGES = 50         # number of donor→recipient flows to include
MIN_USD = 0            # drop flows below this amount
USE_ISO_LABELS = True  # True = ISO3 codes; False = full names
DARK_MODE = False      # Toggle dark mode theme

NODE_THICKNESS = 18
NODE_PADDING   = 20
LINK_ALPHA     = 0.55  # 0..1 transparency for links

# ============================================================
# ✅ HELPERS
# ============================================================
def hex_to_rgba(hex_color: str, alpha: float = 0.5) -> str:
    """Convert '#RRGGBB' to 'rgba(r,g,b,a)'. Alpha in [0,1]."""
    hex_color = hex_color.strip()
    if hex_color.startswith("#"):
        hex_color = hex_color[1:]
    if len(hex_color) != 6:
        # Fallback to grey if malformed
        return f"rgba(127,127,127,{alpha})"
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"

# ============================================================
# ✅ LOAD DATA
# ============================================================
xls = pd.ExcelFile(INPUT_FILE)
sheet = xls.sheet_names[0]  # use first sheet safely
df = pd.read_excel(INPUT_FILE, sheet_name=sheet)

# Ensure normalized value exists (fallback to raw)
val_col = "Value (USD)_normalized"
if val_col not in df.columns:
    if "Value (USD)" in df.columns:
        df[val_col] = df["Value (USD)"]
    else:
        raise ValueError("Neither 'Value (USD)_normalized' nor 'Value (USD)' found in the sheet.")

# Cast to numeric just in case
df[val_col] = pd.to_numeric(df[val_col], errors="coerce")

# Filter individual economies only
df["is_indiv_donor"] = df["funding_economy_code"].notna()
df["is_indiv_recipient"] = df["recipient_economy_code"].notna()
df = df[df["is_indiv_donor"] & df["is_indiv_recipient"]]

# Year filter
if YEAR != "both":
    df = df[df["Year"] == int(YEAR)]

# Positive flows only (ignore corrections/reversals)
df = df[df[val_col] > 0]

# Label choice
if USE_ISO_LABELS:
    donor_label = "funding_economy_code"
    recip_label = "recipient_economy_code"
else:
    donor_label = "Funding economy"
    recip_label = "Recipient country or region"

# ============================================================
# ✅ AGGREGATE FLOWS
# ============================================================
agg = (
    df.groupby([donor_label, recip_label], dropna=False)[val_col]
      .sum()
      .reset_index()
      .sort_values(val_col, ascending=False)
)

if MIN_USD > 0:
    agg = agg[agg[val_col] >= MIN_USD]

if TOP_EDGES > 0:
    agg = agg.head(TOP_EDGES)

if agg.empty:
    raise SystemExit("No edges to plot after filtering. Try lowering TOP_EDGES or MIN_USD.")

# ============================================================
# ✅ BUILD NODE & LINK LISTS
# ============================================================
donors = agg[donor_label].astype(str).tolist()
recips = agg[recip_label].astype(str).tolist()

all_nodes = pd.Index(donors + recips).unique().tolist()
node_index = {name: i for i, name in enumerate(all_nodes)}

sources = [node_index[d] for d in donors]
targets = [node_index[r] for r in recips]
values  = agg[val_col].tolist()

# ============================================================
# ✅ COLOR STYLING
# ============================================================
# Palettes (distinct ranges for donors vs recipients)
donor_palette = [
    "#1f77b4", "#2ca02c", "#d62728", "#9467bd", "#8c564b",
    "#17becf", "#7f7f7f", "#bcbd22", "#e377c2"
]
recipient_palette = [
    "#ff7f0e", "#aec7e8", "#ffbb78", "#98df8a", "#ff9896",
    "#c5b0d5", "#c49c94", "#f7b6d2", "#dbdb8d"
]

# Assign colors
donor_unique = list(pd.Index(donors).unique())
recipient_unique = list(pd.Index(recips).unique())

color_map = {}
for i, d in enumerate(donor_unique):
    color_map[d] = donor_palette[i % len(donor_palette)]
for i, r in enumerate(recipient_unique):
    color_map[r] = recipient_palette[i % len(recipient_palette)]

node_colors = [color_map.get(name, "#7f7f7f") for name in all_nodes]

# Link colors follow donor color with alpha (rgba)
link_colors = [hex_to_rgba(color_map.get(donors[i], "#7f7f7f"), LINK_ALPHA) for i in range(len(donors))]

# ============================================================
# ✅ LAYOUT THEME
# ============================================================
bgcolor  = "#111111" if DARK_MODE else "white"
fontcol  = "white" if DARK_MODE else "black"
title_yr = "2021 + 2022" if YEAR == "both" else str(YEAR)
title    = f"Transport Funding — Individual Economies — Sankey ({title_yr})"

outfile = Path(OUTPUT_DIR) / f"sankey_styled_{YEAR}_{'iso' if USE_ISO_LABELS else 'names'}.html"

# ============================================================
# ✅ PLOT SANKEY
# ============================================================
fig = go.Figure(data=[go.Sankey(
    arrangement="snap",
    node=dict(
        pad=NODE_PADDING,
        thickness=NODE_THICKNESS,
        label=all_nodes,
        color=node_colors,
        line=dict(color="black", width=0.3 if not DARK_MODE else 0.2),
        hovertemplate="%{label}<extra></extra>",
    ),
    link=dict(
        source=sources,
        target=targets,
        value=values,
        color=link_colors,
        hovertemplate="Donor → Recipient<br>%{value:,.0f} USD<extra></extra>",
    )
)])

fig.update_layout(
    title=dict(text=title, font=dict(size=20, color=fontcol), x=0.5),
    font=dict(color=fontcol, size=12),
    plot_bgcolor=bgcolor,
    paper_bgcolor=bgcolor,
    margin=dict(l=20, r=20, t=60, b=20)
)

fig.write_html(str(outfile), include_plotlyjs="cdn")

print("=====================================================")
print(f"[OK] Styled Sankey saved → {outfile.resolve()}")
print(f"Nodes: {len(all_nodes)} | Edges: {len(values)}")
print(f"Total USD included: {agg[val_col].sum():,.2f}")
print("=====================================================")
