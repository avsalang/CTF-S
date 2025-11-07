# app.py
# Streamlit dashboard for transport-related support (CTF-S)
# - NaN-safe, ordered Sankey (donors by outgoing, recipients by incoming)
# - Stylized colors (donor/recipient palettes + rgba link transparency)
# - Year filter (2021 / 2022 / both)
# - Scope filter (Individual economies only / All entities)
# - Top N donors & recipients (bar plots, configurable)
# - Top 3 donors & recipients (text summary)
# - Word map from Title field (TW Cen MT font, default colors)
# - Histograms for Sector_text1 / Subsector_text1 (black text, TW Cen MT font)
# - Raw table tab + CSV download
# - Readability controls for Sankey labels (ISO3 vs names, wrap/ellipsis, height)

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path
from wordcloud import WordCloud, STOPWORDS
import matplotlib.pyplot as plt
import re
import textwrap
import plotly.express as px

st.set_page_config(page_title="Transport Funding (CTF-S) — Dashboard", layout="wide")

# =========================
# CONFIG
# =========================
DEFAULT_FILE = "CTF-S transport related (see notes).xlsx"
VAL_COL_NORM = "Value (USD)_normalized"
VAL_COL_RAW = "Value (USD)"
TITLE_COL = "Title of the project programme, activity or other"
FONT_FAMILY = "TW Cen MT"  # <-- Set desired font


# =========================
# HELPERS
# =========================
def money(x: float) -> str:
    """Format float as currency string."""
    try:
        # Use 'M' for millions, 'K' for thousands for bar plot labels
        if x >= 1_000_000_000:
            return f"${x / 1_000_000_000:,.2f}B"
        if x >= 1_000_000:
            return f"${x / 1_000_000:,.2f}M"
        if x >= 1000:
            return f"${x / 1000:,.2f}K"
        return f"${x:,.2f}"
    except Exception:
        return "-"

def money_full(x: float) -> str:
    """Format float as full currency string for tooltips/summaries."""
    try:
        return f"${x:,.2f}"
    except Exception:
        return "-"


import textwrap


def wrap_label(text, width=22):
    """Wrap labels to multiple lines for plotting."""
    if not isinstance(text, str):
        return text
    return "<br>".join(textwrap.wrap(text, width=width))


def hex_to_rgba(hex_color: str, alpha: float = 0.55) -> str:
    """Convert '#RRGGBB' to 'rgba(r,g,b,a)' with given alpha."""
    hc = hex_color.strip().lstrip("#")
    if len(hc) != 6:
        return f"rgba(127,127,127,{alpha})"
    r = int(hc[0:2], 16)
    g = int(hc[2:4], 16)
    b = int(hc[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def shorten_label(s: str, max_len: int = 14, wrap: bool = True) -> str:
    """Shorten label for readability. If wrap=True, insert line breaks; else ellipsize."""
    s = str(s)
    if len(s) <= max_len:
        return s
    if wrap:
        return "\n".join(textwrap.wrap(s, width=max_len, break_long_words=False))
    else:
        return s[: max_len - 1] + "…"


@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    sheet = xls.sheet_names[0]  # first sheet (safe)
    df = pd.read_excel(path, sheet_name=sheet)

    # normalized value
    if VAL_COL_NORM not in df.columns:
        if VAL_COL_RAW in df.columns:
            df[VAL_COL_NORM] = df[VAL_COL_RAW]
        else:
            raise ValueError("Neither 'Value (USD)_normalized' nor 'Value (USD)' found.")

    df[VAL_COL_NORM] = pd.to_numeric(df[VAL_COL_NORM], errors="coerce")

    # flags for individual economies (ISO present)
    df["is_indiv_donor"] = df["funding_economy_code"].notna()
    df["is_indiv_recipient"] = df["recipient_economy_code"].notna()
    return df


def filter_scope(df: pd.DataFrame, year_sel: str, individual_only: bool) -> pd.DataFrame:
    sub = df.copy()
    if year_sel != "both":
        sub = sub[sub["Year"] == int(year_sel)]
    # keep only positive flows (ignore corrections)
    sub = sub[sub[VAL_COL_NORM] > 0]
    if individual_only:
        sub = sub[sub["is_indiv_donor"] & sub["is_indiv_recipient"]]
    return sub


def sankey_stylized(
        df_flows: pd.DataFrame,
        donor_label_col: str,  # "funding_economy_code" or "Funding economy"
        recip_label_col: str,  # "recipient_economy_code" or "Recipient country or region"
        value_col: str,
        title: str,
        *,
        top_edges: int = 60,
        min_usd: float = 1_000_000,
        compact_labels: bool = True,
        max_label_len: int = 14,
        wrap_labels: bool = True,
        height_px: int = 750
):
    """
    Ordered Sankey with 'Others' buckets:
      - Donors sorted by total outgoing (desc)
      - Recipients sorted by total incoming (desc)
      - Edges beyond Top-N are grouped into:
          * Others (donors) -> each top recipient
          * each top donor -> Others (recipients)
          * Others (donors) -> Others (recipients)
    """
    OTH_DON = "(Other sources)"
    OTH_REC = "(Other recipients)"
    DONOR_MISS = "(Unspecified source)"
    RECIP_MISS = "(Unspecified recipient)"

    # -------- 1) Sanitize labels --------
    dfw = df_flows.copy()
    dfw[donor_label_col] = dfw[donor_label_col].fillna(DONOR_MISS)
    dfw[recip_label_col] = dfw[recip_label_col].fillna(RECIP_MISS)

    # -------- 2) Aggregate ALL edges, then prune by min_usd and split Top vs Rest --------
    agg_all = (
        dfw.groupby([donor_label_col, recip_label_col], dropna=False)[value_col]
        .sum()
        .reset_index()
        .sort_values(value_col, ascending=False)
    )

    if min_usd > 0:
        agg_all = agg_all[agg_all[value_col] >= min_usd]
    if agg_all.empty:
        return None, 0.0
    agg_top = agg_all.head(top_edges) if top_edges > 0 else agg_all.copy()
    agg_rest = agg_all.iloc[len(agg_top):] if top_edges > 0 else agg_all.iloc[0:0]
    donors_top = set(agg_top[donor_label_col].tolist())
    recips_top = set(agg_top[recip_label_col].tolist())

    # -------- 3) Build 'Others' aggregates from the REST edges --------
    others_chunks = []
    if not agg_rest.empty:
        a = agg_rest[~agg_rest[donor_label_col].isin(donors_top) & agg_rest[recip_label_col].isin(recips_top)]
        if not a.empty:
            a2 = (a.groupby(recip_label_col)[value_col].sum().reset_index())
            a2[donor_label_col] = OTH_DON
            others_chunks.append(a2[[donor_label_col, recip_label_col, value_col]])

        b = agg_rest[agg_rest[donor_label_col].isin(donors_top) & ~agg_rest[recip_label_col].isin(recips_top)]
        if not b.empty:
            b2 = (b.groupby(donor_label_col)[value_col].sum().reset_index())
            b2[recip_label_col] = OTH_REC
            others_chunks.append(b2[[donor_label_col, recip_label_col, value_col]])

        c = agg_rest[~agg_rest[donor_label_col].isin(donors_top) & ~agg_rest[recip_label_col].isin(recips_top)]
        if not c.empty:
            c_sum = c[value_col].sum()
            others_chunks.append(pd.DataFrame({
                donor_label_col: [OTH_DON],
                recip_label_col: [OTH_REC],
                value_col: [c_sum]
            }))

    if others_chunks:
        agg_final = pd.concat([agg_top] + others_chunks, ignore_index=True)
    else:
        agg_final = agg_top.copy()

    # -------- 4) Prepare lists for Sankey links --------
    donors = agg_final[donor_label_col].tolist()
    recips = agg_final[recip_label_col].tolist()
    values = agg_final[value_col].tolist()

    # -------- 5) ORDER NODES by size (descending) within each side --------
    donor_totals = (
        agg_final.groupby(donor_label_col)[value_col]
        .sum()
        .sort_values(ascending=False)
    )
    recip_totals = (
        agg_final.groupby(recip_label_col)[value_col]
        .sum()
        .sort_values(ascending=False)
    )
    donor_list_unique = donor_totals.index.tolist()
    recipient_list_unique = recip_totals.index.tolist()

    def move_to_end(lst, name):
        if name in lst:
            lst = [x for x in lst if x != name] + [name]
        return lst

    donor_list_unique = move_to_end(donor_list_unique, OTH_DON)
    recipient_list_unique = move_to_end(recipient_list_unique, OTH_REC)
    all_nodes_full = donor_list_unique + recipient_list_unique

    # -------- 6) Compact display labels (keep full for hover) --------
    if compact_labels:
        all_nodes_shown = [shorten_label(s, max_label_len, wrap_labels) for s in all_nodes_full]
    else:
        all_nodes_shown = all_nodes_full

    node_index = {name: i for i, name in enumerate(all_nodes_full)}
    try:
        sources = [node_index[d] for d in donors]
        targets = [node_index[r] for r in recips]
    except KeyError as e:
        raise KeyError(f"Label not found in node_index: {e}.")

    # -------- 7) Colors (donor vs recipient palettes; Others in gray) --------
    donor_palette = [
        "#1f77b4", "#2ca02c", "#d62728", "#9467bd", "#8c564b",
        "#17becf", "#7f7f7f", "#bcbd22", "#e377c2"
    ]
    recipient_palette = [
        "#ff7f0e", "#aec7e8", "#ffbb78", "#98df8a", "#ff9896",
        "#c5b0d5", "#c49c94", "#f7b6d2", "#dbdb8d"
    ]
    OTHERS_COLOR = "#9e9e9e"
    color_map = {}
    for i, d in enumerate([x for x in donor_list_unique if x != OTH_DON]):
        color_map[d] = donor_palette[i % len(donor_palette)]
    for i, r in enumerate([x for x in recipient_list_unique if x != OTH_REC]):
        color_map[r] = recipient_palette[i % len(recipient_palette)]
    color_map[OTH_DON] = OTHERS_COLOR
    color_map[OTH_REC] = OTHERS_COLOR
    node_colors = [color_map.get(n, "#7f7f7f") for n in all_nodes_full]

    def hex_to_rgba(hex_color: str, alpha: float = 0.55) -> str:
        hc = hex_color.strip().lstrip("#")
        if len(hc) != 6:
            return f"rgba(127,127,127,{alpha})"
        r = int(hc[0:2], 16);
        g = int(hc[2:4], 16);
        b = int(hc[4:6], 16)
        return f"rgba({r},{g},{b},{alpha})"

    link_colors = [hex_to_rgba(color_map.get(donors[i], "#7f7f7f"), 0.55) for i in range(len(donors))]

    # -------- 8) Figure (styled) --------
    bgcolor = "white"
    fontcol = "black"  # <-- SET TO BLACK
    node_customdata = [[full] for full in all_nodes_full]

    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node=dict(
            pad=45, thickness=22,  # <-- Increased padding
            label=all_nodes_shown,
            color=node_colors,
            line=dict(color="black", width=0.4),
            customdata=node_customdata,
            hovertemplate="%{customdata[0]}<extra></extra>",
        ),
        link=dict(
            source=sources, target=targets, value=values,
            color=link_colors,
            hovertemplate="Donor → Recipient<br>%{value:,.0f} USD<extra></extra>",
        )
    )])

    fig.update_layout(
        # <-- FONT FAMILY AND COLOR SET
        title=dict(text=title, font=dict(size=22, color=fontcol, family=FONT_FAMILY), x=0.5),
        font=dict(color=fontcol, size=14, family=FONT_FAMILY),
        # -->
        plot_bgcolor=bgcolor, paper_bgcolor=bgcolor,
        margin=dict(l=20, r=20, t=65, b=20),
        height=height_px
    )

    # <-- ADDED
    # Explicitly set the node label text font to ensure it's simple black
    fig.update_traces(textfont=dict(color="black", family=FONT_FAMILY, size=14))
    # -->

    return fig, float(agg_final[value_col].sum())


def make_wordcloud(text_series: pd.Series, width=1200, height=480, bg="white"):
    text = " ".join([str(t) for t in text_series.dropna().tolist()])
    text = re.sub(r"[^\w\s-]", " ", text)
    stop = set(STOPWORDS) | {"project", "programme", "program", "activity", "support", "undp", "national"}

    font_path_win = "C:/Windows/Fonts/TWCenMT.ttf"  # Standard Windows path

    try:
        # Try with specific font, but default colors
        wc = WordCloud(
            width=width, height=height,
            background_color=bg,
            stopwords=stop,
            collocations=False,
            font_path=font_path_win
            # <-- color_func removed to restore default colors
        )
        return wc.generate(text)
    except Exception as e:
        # Fallback to default font if "TW Cen MT" is not found
        st.warning(
            f"Font '{FONT_FAMILY}' not found at {font_path_win} ({e}). Falling back to default font for word cloud.")
        wc = WordCloud(
            width=width, height=height,
            background_color=bg,
            stopwords=stop,
            collocations=False
            # <-- color_func removed to restore default colors
        )
        return wc.generate(text)
    # -->


def clean_text_counts(s: pd.Series, top_n=30):
    s2 = s.fillna("").astype(str).str.strip()
    mask = (s2 != "") & (~s2.str.upper().isin(["NA", "NR"]))
    counts = s2[mask].value_counts().head(top_n)
    return counts


# =========================
# SIDEBAR (with readability controls)
# =========================
st.sidebar.header("Controls")

data_path = st.sidebar.text_input("Excel file path", value=DEFAULT_FILE)
year_sel = st.sidebar.selectbox("Year", ["both", "2021", "2022"], index=0)
scope_sel = st.sidebar.selectbox("Scope", ["Individual economies", "All entities"], index=0)

st.sidebar.subheader("Sankey Diagram")
# Label readability options
use_iso = st.sidebar.checkbox("Use ISO3 labels in Sankey", value=True)
compact_labels = st.sidebar.checkbox("Compact Sankey labels (wrap/ellipsis)", value=True)
max_label_len = st.sidebar.slider("Max label length", 8, 30, 14, 1)
wrap_labels = st.sidebar.checkbox("Wrap long labels to multiple lines", value=True)
top_edges = st.sidebar.slider("Sankey: Top N edges", 10, 200, 30, 10)
min_usd = st.sidebar.number_input("Sankey: Min USD per edge", min_value=0.0, value=1_000_000.0, step=100_000.0,
                                  format="%.0f")
sankey_height = st.sidebar.slider("Sankey height (px)", 400, 1200, 800, 50)

# <-- ADDED: Control for Top N Bar Plots
st.sidebar.subheader("Top N Bar Plots")
top_n_bar = st.sidebar.number_input("Top N Donors/Recipients", min_value=1, max_value=50, value=5, step=1)
# -->

st.sidebar.subheader("Histograms")
histo_label_width = st.sidebar.slider("Histogram Label Wrap Width", 10, 50, 50, 1)


# =========================
# LOAD + FILTER
# =========================
try:
    df_all = load_data(data_path)
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

individual_only = (scope_sel == "Individual economies")
df = filter_scope(df_all, year_sel, individual_only)

st.title("Funding for Transport-related Activities (CTF-S 2021-2022)")

# =========================
# METRICS
# =========================
m1, m2, m3, m4 = st.columns(4)
m1.metric("Records", f"{len(df):,}")
m2.metric("Total amount in USD (filtered)", money_full(df[VAL_COL_NORM].sum()))
m3.metric("Unique sources", f"{df['Funding economy'].nunique():,}")
m4.metric("Unique recipients", f"{df['Recipient country or region'].nunique():,}")

# =========================
# TABS
# =========================
tab1, tab2, tab3, tab4 = st.tabs(["Funding", "Word Map", "Sector Histograms", "Raw Table"])

# ---- TAB 1: Sankey + Rankings
with tab1:
    st.subheader("Flow of transport funding")
    if use_iso and scope_sel == "Individual economies":
        donor_label = "funding_economy_code"
        recip_label = "recipient_economy_code"
    else:
        donor_label = "Funding economy"
        recip_label = "Recipient country or region"

    fig, total_in_sankey = sankey_stylized(
        df,
        donor_label_col=donor_label,
        recip_label_col=recip_label,
        value_col=VAL_COL_NORM,
        title=f"Transport Funding ({'2021+2022' if year_sel == 'both' else year_sel}) {scope_sel}",
        top_edges=top_edges,
        min_usd=min_usd,
        compact_labels=compact_labels,
        max_label_len=max_label_len,
        wrap_labels=wrap_labels,
        height_px=sankey_height
    )
    if fig is None:
        st.warning("No flows to plot with the current filters. Adjust Top N or Min USD.")
    else:
        st.plotly_chart(fig, use_container_width=True)
        st.caption(f"Total amoutn (USD) represented in figure: **{money_full(total_in_sankey)}**")

    # <-- ADDED: Top N Bar Plots
    st.markdown("---")
    st.subheader(f"Top {top_n_bar} Sources & Recipients")

    # --- Calculate Data for Bar Plots (and reuse for Top 3 summary) ---
    if individual_only:
        donor_group = df[df["is_indiv_donor"]].groupby("Funding economy")
    else:
        donor_group = df.groupby("Funding economy")
    donor_series_all = donor_group[VAL_COL_NORM].sum().sort_values(ascending=False)
    donor_series_bar = donor_series_all.head(top_n_bar)

    if individual_only:
        recip_group = df[df["is_indiv_recipient"]].groupby("Recipient country or region")
    else:
        recip_group = df.groupby("Recipient country or region")
    recip_series_all = recip_group[VAL_COL_NORM].sum().sort_values(ascending=False)
    recip_series_bar = recip_series_all.head(top_n_bar)

    # --- Plot Bar Charts ---
    c1_bar, c2_bar = st.columns(2)

    with c1_bar:
        st.write(f"**Top {top_n_bar} Sources**")
        if donor_series_bar.empty:
            st.info("No donors found for current scope/filters.")
        else:
            # Sort ascending for horizontal bar plot (biggest at top)
            data_to_plot = donor_series_bar.sort_values(ascending=True)
            fig_donor_bar = px.bar(
                data_to_plot,
                orientation="h",
                labels={"value": "Total Value (USD)", "index": "Donor"},
                text=data_to_plot.apply(money) # Use abbreviated money format
            )
            fig_donor_bar.update_layout(
                yaxis_title=None,
                xaxis_title="Total Value (USD)",
                yaxis=dict(automargin=True),
                font_family=FONT_FAMILY,
                font_color="black",
                showlegend=False,
                plot_bgcolor="white"
            )
            # Add full value to hover template
            fig_donor_bar.update_traces(
                textposition='outside',
                hovertemplate='<b>%{label}</b><br>Value: %{customdata}<extra></extra>',
                customdata=data_to_plot.apply(money_full)
            )
            st.plotly_chart(fig_donor_bar, use_container_width=True)

    with c2_bar:
        st.write(f"**Top {top_n_bar} Recipients**")
        if recip_series_bar.empty:
            st.info("No recipients found for current scope/filters.")
        else:
            # Sort ascending for horizontal bar plot (biggest at top)
            data_to_plot = recip_series_bar.sort_values(ascending=True)
            fig_recip_bar = px.bar(
                data_to_plot,
                orientation="h",
                labels={"value": "Total Value (USD)", "index": "Recipient"},
                text=data_to_plot.apply(money) # Use abbreviated money format
            )
            fig_recip_bar.update_layout(
                yaxis_title=None,
                xaxis_title="Total Value (USD)",
                yaxis=dict(automargin=True),
                font_family=FONT_FAMILY,
                font_color="black",
                showlegend=False,
                plot_bgcolor="white"
            )
            # Add full value to hover template
            fig_recip_bar.update_traces(
                textposition='outside',
                hovertemplate='<b>%{label}</b><br>Value: %{customdata}<extra></st.caption(f"Total amoutn (USD) represented in figure: **{money_full(total_in_sankey)}**")ra>',
                customdata=data_to_plot.apply(money_full)
            )
            st.plotly_chart(fig_recip_bar, use_container_width=True)
    # -->


# ---- TAB 2: Word Map (from Titles)
with tab2:
    st.subheader("Project keywords")
    if TITLE_COL not in df.columns or df[TITLE_COL].dropna().empty:
        st.info("No titles available.")
    else:
        wc = make_wordcloud(df[TITLE_COL], width=1200, height=480, bg="white")
        fig_wc, ax = plt.subplots(figsize=(14, 5))
        ax.imshow(wc, interpolation="bilinear")
        ax.axis("off")
        st.pyplot(fig_wc, clear_figure=True)

# ---- TAB 3: Sector / Subsector histograms
with tab3:
    st.subheader("Sector distribution")


    def wrap_label(text, width=25):
        import textwrap
        if not isinstance(text, str):
            return text
        return "<br>".join(textwrap.wrap(text, width=width))


    def get_clean_counts(series, top_n=30, label_width=25):
        if series.isna().all():
            return pd.Series(dtype=int)
        s = (
            series.dropna()
            .astype(str)
            .str.strip()
        )
        s = s[(s != "") & (s.str.upper() != "NR")]
        if s.empty:
            return pd.Series(dtype=int)
        counts = s.value_counts().head(top_n)
        counts.index = [wrap_label(x, width=label_width) for x in counts.index]
        return counts


    # ---------------------------
    # Sector_text1 Horizontal Bar
    # ---------------------------
    if "Sector_text1" in df.columns:
        counts1 = get_clean_counts(df["Sector_text1"], top_n=30, label_width=histo_label_width)
        if counts1.empty:
            st.info("No valid Sector_text1 values.")
        else:
            fig_sec = px.bar(
                counts1.sort_values(ascending=True),
                orientation="h",
                labels={"value": "Count", "index": "Sector"},
                title="Top 30 sectors",
            )
            fig_sec.update_layout(
                height=1200,
                margin=dict(l=20, r=20, t=50, b=20),
                xaxis_title="Count",
                yaxis_title="Sector",
                yaxis=dict(automargin=True),  # Avoid label overlap
                font_family=FONT_FAMILY,  # <-- SET FONT
                font_color="black",  # <-- SET COLOR
                plot_bgcolor="white"
            )
            st.plotly_chart(fig_sec, use_container_width=True)
    else:
        st.info("Column 'Sector_text1' not found.")

    # --------------------------------
    # Subsector_text1 Horizontal Bar
    # --------------------------------
    st.subheader("Subsector distribution")

    if "Subsector_text1" in df.columns:
        counts2 = get_clean_counts(df["Subsector_text1"], top_n=30, label_width=histo_label_width)
        if counts2.empty:
            st.info("No valid Subsector_text1 values.")
        else:
            fig_sub = px.bar(
                counts2.sort_values(ascending=True),
                orientation="h",
                labels={"value": "Count", "index": "Subsector"},
                title="Top 30 subsectors",
            )
            fig_sub.update_layout(
                height=1200,
                margin=dict(l=20, r=20, t=50, b=20),
                xaxis_title="Count",
                yaxis_title="Subsector",
                yaxis=dict(automargin=True),  # Avoid label overlap
                font_family=FONT_FAMILY,  # <-- SET FONT
                font_color="black",  # <-- SET COLOR
                plot_bgcolor="white"
            )
            st.plotly_chart(fig_sub, use_container_width=True)
    else:
        st.info("Column 'Subsector_text1' not found.")

# ---- TAB 4: Raw Table
with tab4:
    st.subheader("Raw Table (Filtered)")
    st.dataframe(df, use_container_width=True)
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered CSV", csv, file_name="ctfs_transport_filtered.csv", mime="text/csv")