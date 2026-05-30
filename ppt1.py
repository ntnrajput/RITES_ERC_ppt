import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import os
import copy
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
import numpy as np
import re
from difflib import get_close_matches
import sys



MONTHS = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december"
]

def normalize_month(val):
    if pd.isna(val):
        return val

    text = str(val).strip()

    # Lowercase + clean junk characters
    text_clean = re.sub(r"[^a-zA-Z0-9 ]", " ", text).lower()

    # Try to correct month spelling using fuzzy match
    for word in text_clean.split():
        match = get_close_matches(word, MONTHS, n=1, cutoff=0.7)
        if match:
            text_clean = text_clean.replace(word, match[0])

    # Try parsing again
    try:
        dt = pd.to_datetime(text_clean, errors="raise")
        return dt.strftime("%b %y")   # Jan 26
    except Exception:
        return text   # fallback if still bad

# ================== PATHS ==================
excel_file = sys.argv[1]
sheet_name = "Copy of Final Month Wise Defect"

template_ppt = "RITES_Template.pptx"
output_ppt = "ERC_Defect_Analysis_RITES_FINAL.pptx"

# ================== READ EXCEL ==================


df = pd.read_excel(excel_file, sheet_name=sheet_name, header=1)

# Clean column names properly
df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.replace("\n", " ", regex=True)
)

print("Columns found:", df.columns.tolist())

df.rename(columns={
    "Firm Name": "Plant",
    "Month": "Month"
}, inplace=True)

df["Month"] = df["Month"].apply(normalize_month)  
df["_Month_dt"] = pd.to_datetime(df["Month"], format="%b %y", errors="coerce")


rio_prefixes = ["NRIO", "SRIO", "ERIO", "WRIO", "CRIO"]

def clean_plant_name(name):
    if not isinstance(name, str):
        return name
    for prefix in rio_prefixes:
        if name.strip().startswith(prefix):
            return name.replace(prefix, "", 1).strip()
    return name.strip()

df["Plant"] = df["Plant"].apply(clean_plant_name)


exclude_keywords = [
    "%", "Inspected", "Acceptance", "Rejection",
    "PO", "Zonal", "Consignee", "Unnamed"
]

defect_cols = [
    c for c in df.columns
    if not any(k in c for k in exclude_keywords)
    and c not in ["Plant", "Month", "_Month_dt"]
    and pd.api.types.is_numeric_dtype(df[c])
]

print("Non-numeric columns in defect_cols:")
for c in defect_cols:
    print(c, df[c].dtype)

df_diag = df.copy()
df_diag["Total_Defects"] = df_diag[defect_cols].fillna(0).sum(axis=1)

# Plants where inspection happened but no defect breakup exists
zero_defect_plants = df_diag[
    (df_diag["Inspected (Nos.)"] > 0) &
    (df_diag["Total_Defects"] == 0)
]["Plant"].unique()

print("\n Plants with inspection but NO defect breakup (charts skipped):")
for p in zero_defect_plants:
    print(" -", p)

print("\nTotal such plants:", len(zero_defect_plants))



df_long = df.melt(
    id_vars=["Plant", "Month"],
    value_vars=defect_cols,
    var_name="Defect",
    value_name="Rejected"
)

nan_count = df_long["Rejected"].isna().sum()
print(f"Number of NaN data points replaced with 0: {nan_count}")

df_long["Rejected"] = df_long["Rejected"].fillna(0)

df_long = df_long[df_long["Rejected"] > 0]

# ================== PIE FUNCTION ==================
def plot_best_worst_performers(df, img_path, mode="best", top_n=8):
    """
    mode = "best"  -> lowest rejection %
    mode = "worst" -> highest rejection %
    """

    summary = get_plant_rejection_summary(df)

    # Remove zero inspection plants
    summary = summary[summary["Inspected (Nos.)"] > 0]

    if mode == "best":
        summary = summary.sort_values("Rejection %", ascending=True).head(top_n)
        title = "Top Performing Manufacturers (Lowest Rejection %)"
        color = "#2e7d32"   # green
    else:
        summary = summary.sort_values("Rejection %", ascending=False).head(top_n)
        title = "Worst Performing Manufacturers (Highest Rejection %)"
        color = "#c62828"   # red

    fig, ax = plt.subplots(figsize=(12, 5))

    bars = ax.bar(
        summary["Plant"],
        summary["Rejection %"],
        color=color
    )

    ax.set_ylabel("Rejection %")
    ax.set_title(title, fontsize=14, fontweight="bold")

    ax.set_xticks(range(len(summary)))
    ax.set_xticklabels(
        summary["Plant"],
        rotation=45,
        ha="right",
        fontsize=8
    )

    # Add value labels
    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width()/2,
            height,
            f"{height:.2f}%",
            ha="center",
            va="bottom",
            fontsize=8
        )

    plt.subplots_adjust(bottom=0.30)
    plt.savefig(img_path, dpi=300, bbox_inches="tight")
    plt.close()



def add_best_worst_slides(prs, base_slide, df):

    # -------- BEST PERFORMERS --------
    best_img = "charts/best_performers.png"
    plot_best_worst_performers(df, best_img, mode="best", top_n=8)

    slide1 = duplicate_slide(prs, base_slide)
    clear_old_charts(slide1)

    title_box = slide1.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(11), Inches(0.6)
    )

    p = title_box.text_frame.paragraphs[0]
    p.text = "Top Performing Manufacturers (Lowest Rejection %)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.alignment = 1

    slide1.shapes.add_picture(
        best_img,
        Inches(0.5),
        Inches(1.1),
        width=Inches(11)
    )

    # -------- WORST PERFORMERS --------
    worst_img = "charts/worst_performers.png"
    plot_best_worst_performers(df, worst_img, mode="worst", top_n=8)

    slide2 = duplicate_slide(prs, base_slide)
    clear_old_charts(slide2)

    title_box = slide2.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(11), Inches(0.6)
    )

    p = title_box.text_frame.paragraphs[0]
    p.text = "Worst Performing Manufacturers (Highest Rejection %)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.alignment = 1

    slide2.shapes.add_picture(
        worst_img,
        Inches(0.5),
        Inches(1.1),
        width=Inches(11)
    )





def get_plant_rejection_summary(df):
    """
    Returns plant-wise overall rejection percentage
    combining all PO numbers and months
    """

    summary = (
        df.groupby("Plant", as_index=False)
        .agg({
            "Inspected (Nos.)": "sum",
            "Rejection (Nos.)": "sum"
        })
    )

    summary["Rejection %"] = (
        summary["Rejection (Nos.)"] /
        summary["Inspected (Nos.)"] * 100
    ).fillna(0)

    return summary



def plot_clean_donut(summary, plant_name, img_path, color_map=None):
    top_n = 6
    top = summary.head(top_n).copy()
    others = summary.iloc[top_n:].sum()

    if others > 0:
        top["Others"] = others

    labels = list(top.index)
    values = top.values
    colors = [
        get_defect_color(defect, color_map=color_map, index=i)
        for i, defect in enumerate(labels)
    ]

    # 🔽 Smaller canvas + tight control
    plt.figure(figsize=(5, 5))

    plt.pie(
        values,
        labels=None,
        autopct='%1.0f%%',
        startangle=140,
        pctdistance=0.78,
        colors=colors,
        wedgeprops=dict(width=0.42, edgecolor="white")
    )

    # 🔽 REMOVE ALL OUTER MARGINS
    plt.axis("equal")
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

    plt.savefig(img_path, dpi=300, bbox_inches="tight", pad_inches=0)
    plt.close()


def get_donut_legend_items(summary, top_n=6):
    top = summary.head(top_n).copy()
    others = summary.iloc[top_n:].sum()
    if others > 0:
        top["Others"] = others
    return list(top.index)



# ================== LOAD TEMPLATE ==================
prs = Presentation(template_ppt)

# ✅ SAFE BASE SLIDE (always exists)
base_slide = prs.slides[0]

plants = sorted(
    df_long["Plant"]
    .dropna()          # remove NaN
    .astype(str)       # ensure all strings
    .unique()
)
os.makedirs("charts", exist_ok=True)

# ================== DUPLICATE SLIDE ==================

DEFECT_COLORS = {
    "Embossing Die Impression / Failure": "#1f4e79",
    "Dimensional": "#4f81bd",
    "Flat Bearing Area": "#9bbb59",
    "Cut bar length": "#8064a2",
    "Turning length": "#c0504d",
    "Hardness": "#f79646",
    "Others": "#bfbfbf"
}
DEFECT_FALLBACK_COLORS = [
    "#2f5597", "#70ad47", "#ed7d31", "#a64d79", "#5b9bd5",
    "#c00000", "#00b0f0", "#948a54", "#ffc000", "#4472c4",
    "#7030a0", "#00b050", "#9e480e", "#264478", "#636363",
    "#ff6699", "#66cc99", "#cc9900", "#3399ff", "#993366"
]

def build_slide_color_map(legend_items):
    color_map = {}
    used_colors = set()

    for defect in legend_items:
        if defect in DEFECT_COLORS:
            color = DEFECT_COLORS[defect]
            color_map[defect] = color
            used_colors.add(color)

    fallback_iter = (c for c in DEFECT_FALLBACK_COLORS if c not in used_colors)
    for defect in legend_items:
        if defect not in color_map:
            color = next(fallback_iter, DEFECT_FALLBACK_COLORS[len(color_map) % len(DEFECT_FALLBACK_COLORS)])
            color_map[defect] = color
            used_colors.add(color)

    return color_map

def get_defect_color(defect, color_map=None, index=0):
    if color_map and defect in color_map:
        return color_map[defect]
    if defect in DEFECT_COLORS:
        return DEFECT_COLORS[defect]
    return DEFECT_FALLBACK_COLORS[index % len(DEFECT_FALLBACK_COLORS)]


def add_summary_table(slide, left, top, inspected, rejected, rej_pct):
    rows, cols = 2, 3
    table = slide.shapes.add_table(
        rows, cols,
        left,
        top,
        Inches(4.8),
        Inches(0.9)
    ).table

    # Column headings
    headers = [
        "Cumulative Inspected (Nos.)",
        "Cumulative Rejection (Nos.)",
        "Rejection (%)"
    ]

    values = [
        f"{int(inspected)}",
        f"{int(rejected)}",
        f"{rej_pct:.2f}"
    ]

    # Header row
    for col in range(cols):
        cell = table.cell(0, col)
        cell.text = headers[col]
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(9)
        p.alignment = 1  # center

    # Value row
    for col in range(cols):
        cell = table.cell(1, col)
        cell.text = values[col]
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(9)
        p.alignment = 1  # center


def add_common_legend(slide, legend_items, color_map=None):
    """
    Adds a vertical legend on the right side of slide.
    legend_items = list of defect names (ordered)
    """

    start_left = Inches(11)
    start_top = Inches(1.4)
    box_size = Inches(0.18)
    spacing = Inches(0.35)

    for i, defect in enumerate(legend_items):
        top = start_top + i * spacing

        # Color box
        rect = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            start_left,
            top,
            box_size,
            box_size
        )
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor.from_string(
            get_defect_color(defect, color_map=color_map, index=i).replace("#", "")
        )
        rect.line.fill.background()

        # Text
        textbox = slide.shapes.add_textbox(
            start_left + Inches(0.25),
            top - Inches(0.02),
            Inches(2.4),
            Inches(0.3)
        )

        tf = textbox.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = defect
        p.font.size = Pt(9)

def duplicate_slide(prs, slide):
    new_slide = prs.slides.add_slide(slide.slide_layout)
    for shape in slide.shapes:
        el = shape.element
        new_slide.shapes._spTree.insert_element_before(
            copy.deepcopy(el),
            'p:extLst'
        )
    return new_slide

def get_cumulative_stats(df, plant):
    plant_df = df[df["Plant"] == plant]

    inspected = plant_df["Inspected (Nos.)"].sum()
    rejected = plant_df["Rejection (Nos.)"].sum()
    rejection_pct = (rejected / inspected * 100) if inspected > 0 else 0

    return inspected, rejected, rejection_pct


def clear_old_charts(slide):
    """
    Remove pictures and charts from a slide
    while keeping background, text boxes, shapes, footer intact.
    Compatible with all python-pptx versions.
    """
    for shape in list(slide.shapes):
        # Remove pictures
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            slide.shapes._spTree.remove(shape.element)

        # Remove charts safely (no enums!)
        elif hasattr(shape, "has_chart") and shape.has_chart:
            slide.shapes._spTree.remove(shape.element)

def add_vertical_divider(slide):
    line = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(5.0),     # X position (middle of slide)
        Inches(0.6),     # Y start
        Inches(0.02),    # thin vertical line
        Inches(6.75)      # height
    )

    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(180, 180, 180)  # light grey
    line.line.fill.background()

def add_month_wise_table(slide, left, top, plant, df, max_months=12):
    """
    Adds Month-wise consolidated table:
    Combines all PO numbers for same Month & Manufacturer
    """

    plant_df = df[df["Plant"] == plant].copy()

    # ---- GROUP BY MONTH (combine all POs) ----
    month_summary = (
        plant_df
        .groupby("Month", as_index=False)
        .agg({
            "Inspected (Nos.)": "sum",
            "Rejection (Nos.)": "sum"
        })
    )

    # ---- Calculate rejection percentage AFTER aggregation ----
    month_summary["% Rej"] = (
        month_summary["Rejection (Nos.)"] /
        month_summary["Inspected (Nos.)"] * 100
    ).fillna(0)

    # ---- Sort & limit rows ----
    
    # ---- Attach sortable month column ----
    month_summary = month_summary.merge(
        plant_df[["Month", "_Month_dt"]].drop_duplicates(),
        on="Month",
        how="left"
    )

    # ---- Sort so MOST RECENT month is on top ----
    month_summary = (
        month_summary
        .sort_values("_Month_dt", ascending=False)
        .head(max_months)
        .drop(columns="_Month_dt")
    )






    rows = len(month_summary) + 1
    cols = 4

    table = slide.shapes.add_table(
        rows,
        cols,
        left,
        top,
        Inches(4.8),
        Inches(1.2)
    ).table

    headers = ["Month", "Inspected (Nos.)", "Rejection (Nos.)", "% Rej"]

    # ---- Header row ----
    for col in range(cols):
        cell = table.cell(0, col)
        cell.text = headers[col]
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(8)
        p.alignment = 1

    # ---- Data rows ----
    for r, (_, row) in enumerate(month_summary.iterrows(), start=1):
        table.cell(r, 0).text = str(row["Month"])
        table.cell(r, 1).text = str(int(row["Inspected (Nos.)"]))
        table.cell(r, 2).text = str(int(row["Rejection (Nos.)"]))
        table.cell(r, 3).text = f'{row["% Rej"]:.2f}'

        for c in range(cols):
            p = table.cell(r, c).text_frame.paragraphs[0]
            p.font.size = Pt(8)
            p.alignment = 1



def plot_spider_chart(plant_summary, img_path):


    # ---- Filter plants with rejection > 1% ----
    data = plant_summary[plant_summary["Rejection %"] > 1].copy()

    if data.empty:
        print("No plants with rejection > 1%")
        return False

    # ---- Sort for clean spiral (low → high) ----
    data = data.sort_values("Rejection %").reset_index(drop=True)

    # ---- Create Plant IDs ----
    data["Plant_ID"] = [f"Plant {i+1}" for i in range(len(data))]

    plant_ids = data["Plant_ID"].tolist()
    plant_names = data["Plant"].tolist()
    values = data["Rejection %"].tolist()

    # ---- Polar coordinates ----
    theta = np.linspace(0, 2 * np.pi, len(values), endpoint=False)
    r = values

    # ---- Large figure to fit legend ----
    fig = plt.figure(figsize=(13, 8))
    ax = plt.subplot(111, polar=True)

    # ---- Spiral curve ----
    ax.plot(theta, r, linewidth=2.5, color="#f79646")
    ax.fill(theta, r, alpha=0.25, color="#f79646")

    # ---- 1% reference circle ----
    ax.plot(
        np.linspace(0, 2 * np.pi, 360),
        [1] * 360,
        linestyle="--",
        linewidth=1.5,
        color="grey"
    )

    # ---- Labels INSIDE chart (Plant 1, Plant 2 + %) ----
    for angle, radius, pid in zip(theta, r, plant_ids):
        ax.text(
            angle,
            radius + 0.08,
            f"{pid}\n{radius:.2f}%",
            ha="center",
            va="center",
            fontsize=10,
            fontweight="bold"
        )

    # ---- Clean axes ----
    ax.set_xticks([])
    ax.set_yticklabels([])
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_ylim(0, max(r) * 1.25)

  
    legend_labels = [
        wrap_text(f"{pid} - {name}", max_chars=35)
        for pid, name in zip(plant_ids, plant_names)
    ]


    from matplotlib.lines import Line2D

# ---- Create dummy legend handles (ONE PER PLANT) ----
    legend_handles = [
        Line2D(
            [0], [0],
            color="#f79646",
            lw=2
        )
        for _ in legend_labels
    ]

    ax.legend(
        handles=legend_handles,
        labels=legend_labels,
        loc="upper left",
        bbox_to_anchor=(1.02, 1),
        fontsize=10,
        frameon=False
    )


    # ---- Layout tuning ----
    plt.subplots_adjust(left=0.05, right=0.78, top=0.950, bottom=0.15)

    plt.savefig(img_path, dpi=300)
    plt.close()

    return True


def wrap_text(text, max_chars=35):
    """
    Inserts line breaks into long legend text
    """
    words = text.split()
    lines = []
    current = ""

    for w in words:
        if len(current) + len(w) + 1 <= max_chars:
            current += (" " if current else "") + w
        else:
            lines.append(current)
            current = w

    if current:
        lines.append(current)

    return "\n".join(lines)


def add_spider_chart_slide(prs, df):
    """
    Adds one slide with spider chart showing plants with rejection > 1%
    """

    plant_summary = get_plant_rejection_summary(df)

    img_path = "charts/spider_rejection_gt_1.png"
    success = plot_spider_chart(plant_summary, img_path)

    if not success:
        return

    slide = duplicate_slide(prs, base_slide)
    clear_old_charts(slide)

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.5),
        Inches(0.3),
        Inches(11),
        Inches(0)
    )
    
    p = title_box.text_frame.paragraphs[0]
    p.text = "Plant-wise Rejection Trend (Rejection > 1%)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.alignment = 1

    # Chart image
    slide.shapes.add_picture(
        img_path,
        Inches(0.5),
        Inches(1.1),
        width=Inches(11)
    )


def plot_pareto_chart(df_long, img_path, top_n=None):
    """
    Creates Pareto chart:
    - Bars: No. of defects
    - Line: Cumulative percentage
    """

    # ---- Aggregate defect counts ----
    pareto = (
        df_long.groupby("Defect")["Rejected"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    if top_n:
        pareto = pareto.head(top_n)

    pareto["Cum %"] = pareto["Rejected"].cumsum() / pareto["Rejected"].sum() * 100

    # ---- Plot ----
    fig, ax1 = plt.subplots(figsize=(11, 5))

    # Bars
    ax1.bar(
        pareto["Defect"],
        pareto["Rejected"],
        color="#1f4e79"
    )
    ax1.set_ylabel("Nos. of Defects")
    ax1.set_xlabel("Defect Name")
    ax1.tick_params(axis="x", rotation=90)

    # Secondary axis for cumulative %
    ax2 = ax1.twinx()
    ax2.plot(
        pareto["Defect"],
        pareto["Cum %"],
        color="#f79646",
        marker="o",
        linewidth=2
    )
    ax2.set_ylabel("Cumulative Percentage Rejection")
    ax2.set_ylim(0, 100)

    # 80% reference line
    ax2.axhline(80, color="grey", linestyle="--", linewidth=1)

    # ---- Layout ----
    plt.title("Pareto Analysis for ERC Defects", fontsize=14, fontweight="bold")
    plt.tight_layout()

    plt.savefig(img_path, dpi=300)
    plt.close()

def add_pareto_slide(prs, base_slide, df_long):

    img_path = "charts/pareto_defects.png"
    plot_pareto_chart(df_long, img_path)

    # ✅ DUPLICATE TEMPLATE (SAFE HERE)
    slide = duplicate_slide(prs, base_slide)
    clear_old_charts(slide)

    # ---- Title ----
    title_box = slide.shapes.add_textbox(
        Inches(0.5),
        Inches(0.3),
        Inches(11),
        Inches(0.6)
    )
    p = title_box.text_frame.paragraphs[0]
    p.text = "Pareto Analysis for ERC Defects"
    p.font.size = Pt(20)
    p.font.bold = True
    p.alignment = 1

    # ---- Pareto Image ----
    slide.shapes.add_picture(
        img_path,
        Inches(0.5),
        Inches(1.2),
        width=Inches(11)
    )

    # ---- Highlight strip ----
    note_box = slide.shapes.add_textbox(
        Inches(2),
        Inches(6.3),
        Inches(8),
        Inches(0.5)
    )
    p = note_box.text_frame.paragraphs[0]
    p.text = (
        "Highlights: Top defects contribute to ~80% of total rejections, "
        "enabling focused process control."
    )
    p.font.size = Pt(11)
    p.font.bold = True


def add_plant_title(slide, text, left, top, width):
    textbox = slide.shapes.add_textbox(left, top, width, Inches(0.4))
    tf = textbox.text_frame
    tf.clear()
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.name = "Calibri"   # will auto-map to RITES font if theme enforces
    p.alignment = 1  # center


def plot_top_performing_units(df, img_path, top_n=12):
    """
    Vertical bar chart with long X-axis labels (no cutting)
    """

    summary = (
        df.groupby("Plant", as_index=False)
        .agg({"Inspected (Nos.)": "sum"})
        .sort_values("Inspected (Nos.)", ascending=False)
        .head(top_n)
    )

    # Convert to thousands
    summary["Inspected (000s)"] = summary["Inspected (Nos.)"] / 1000

    # 🔑 Bigger width + controlled height
    fig, ax = plt.subplots(figsize=(15, 5))

    bars = ax.bar(
        summary["Plant"],
        summary["Inspected (000s)"],
        color="#ed7d31"
    )

    ax.set_ylabel("Inspected Nos (in Thousands)")
    ax.set_title(
        "Manufacturing Units: High Production Units",
        fontsize=14,
        fontweight="bold"
    )

    # 🔑 CRITICAL FIXES FOR X-AXIS
    ax.set_xticks(range(len(summary)))
    ax.set_xticklabels(
        summary["Plant"],
        rotation=45,          # NOT 90
        ha="right",           # right aligned
        fontsize=8
    )

    # 🔑 Force bottom margin for long names
    plt.subplots_adjust(bottom=0.32)

    # Value labels on bars
    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            height,
            f"{height:.1f}",
            ha="center",
            va="bottom",
            fontsize=8
        )

    plt.savefig(img_path, dpi=300, bbox_inches="tight")
    plt.close()


def plot_overall_defect_pie(df_long, img_path):

    summary = (
        df_long.groupby("Defect")["Rejected"]
        .sum()
        .sort_values(ascending=False)
    )

    total = summary.sum()

    legend_labels = [
        f"{d} – {int(v)} ({v/total*100:.0f}%)"
        for d, v in summary.items()
    ]

    # 🔑 Large canvas only for pie
    fig, ax = plt.subplots(figsize=(8.5, 8.5))

    wedges, _ = ax.pie(
        summary.values,
        startangle=140,
        wedgeprops=dict(edgecolor="white")
    )

    ax.set_title(
        "Defect Analysis of all ERC Units",
        fontsize=14,
        fontweight="bold",
        pad=20
    )

    ax.axis("equal")

    # ---- Separate LEGEND ----
    ax.legend(
        wedges,
        legend_labels,
        title="Defect Type",
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        fontsize=10,
        title_fontsize=11,
        frameon=False
    )

    plt.tight_layout()
    plt.savefig(img_path, dpi=300)
    plt.close()


def add_erc_overview_slide(prs, base_slide, df, df_long):

    # ---------- DATA ----------
    inspected = int(df["Inspected (Nos.)"].sum())
    rejected = int(df["Rejection (Nos.)"].sum())
    rej_pct = (rejected / inspected * 100) if inspected else 0

    # Month-wise consolidated
    month_tbl = (
        df.groupby("Month", as_index=False)
        .agg({
            "Inspected (Nos.)": "sum",
            "Rejection (Nos.)": "sum"
        })
    )

    month_tbl = month_tbl.merge(
        df[["Month", "_Month_dt"]].drop_duplicates(),
        on="Month",
        how="left"
    )

    month_tbl = (
        month_tbl
        .sort_values("_Month_dt", ascending=False)
        .drop(columns="_Month_dt")
    )

    month_tbl["% Rej"] = (
        month_tbl["Rejection (Nos.)"] /
        month_tbl["Inspected (Nos.)"] * 100
    ).fillna(0)

    # ---------- PIE ----------
    pie_path = "charts/overall_defect_pie.png"
    plot_overall_defect_pie(df_long, pie_path)

    # ---------- SLIDE ----------
    slide = duplicate_slide(prs, base_slide)
    clear_old_charts(slide)

    # ---------- TITLE ----------
    title = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.3), Inches(7), Inches(0.6)
    )
    p = title.text_frame.paragraphs[0]
    p.text = "ERC Inspection – Since April 2025"
    p.font.size = Pt(20)
    p.font.bold = True

    # ---------- KPI TABLE ----------
    kpi = slide.shapes.add_table(
        2, 3,
        Inches(0.6), Inches(1.1),
        Inches(6.5), Inches(1.1)
    ).table

    headers = [
        "Cumulative Inspected (Nos.)",
        "Cumulative Rejection (Nos.)",
        "Rejection (%)"
    ]
    values = [f"{inspected}", f"{rejected}", f"{rej_pct:.2f}%"]

    for i in range(3):
        kpi.cell(0, i).text = headers[i]
        kpi.cell(1, i).text = values[i]
        for r in range(2):
            p = kpi.cell(r, i).text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.font.bold = (r == 0)
            p.alignment = 1

    # ---------- MONTH-WISE TABLE ----------
    rows = len(month_tbl) + 1
    table = slide.shapes.add_table(
        rows, 4,
        Inches(0.6), Inches(2.4),
        Inches(6.5), Inches(3.8)
    ).table

    headers = ["Month", "Inspected (Nos.)", "Rejected (Nos.)", "% Rej"]
    for c, h in enumerate(headers):
        table.cell(0, c).text = h
        p = table.cell(0, c).text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(9)
        p.alignment = 1

    for r, (_, row) in enumerate(month_tbl.iterrows(), start=1):
        table.cell(r, 0).text = str(row["Month"])
        table.cell(r, 1).text = str(int(row["Inspected (Nos.)"]))
        table.cell(r, 2).text = str(int(row["Rejection (Nos.)"]))
        table.cell(r, 3).text = f'{row["% Rej"]:.2f}'
        for c in range(4):
            p = table.cell(r, c).text_frame.paragraphs[0]
            p.font.size = Pt(9)
            p.alignment = 1

    # ---------- PIE IMAGE ----------
    slide.shapes.add_picture(
        pie_path,
        Inches(6.8),     # move left
        Inches(0.9),     # move up
        width=Inches(6.3)
    )



def add_top_performing_units_slide(prs, base_slide, df):

    img_path = "charts/top_performing_units.png"
    plot_top_performing_units(df, img_path)

    # ✅ Duplicate RITES template
    slide = duplicate_slide(prs, base_slide)
    clear_old_charts(slide)

    # ---- Title ----
    title_box = slide.shapes.add_textbox(
        Inches(0.5),
        Inches(0.3),
        Inches(11),
        Inches(0.6)
    )
    p = title_box.text_frame.paragraphs[0]
    p.text = "Manufacturing Units: High Production Units"
    p.font.size = Pt(20)
    p.font.bold = True
    p.alignment = 1

    # ---- Chart Image ----
    slide.shapes.add_picture(
        img_path,
        Inches(0.5),
        Inches(1.1),
        width=Inches(11)
    )


# ================== CREATE SLIDES ==================
for i in range(0, len(plants), 2):

    slide = duplicate_slide(prs, base_slide)
    clear_old_charts(slide)
    # add_vertical_divider(slide)

    # -------- BUILD COMMON LEGEND FROM BOTH PLANTS ON SLIDE --------
    slide_plants = [plants[k] for k in range(i, min(i + 2, len(plants)))]
    legend_defects = []

    for plant_name in slide_plants:
        plant_df_legend = df_long[df_long["Plant"] == plant_name]
        summary_legend = (
            plant_df_legend.groupby("Defect")["Rejected"]
            .sum()
            .sort_values(ascending=False)
        )

        for defect in get_donut_legend_items(summary_legend):
            if defect not in legend_defects:
                legend_defects.append(defect)

    slide_color_map = build_slide_color_map(legend_defects)
    add_common_legend(slide, legend_defects, color_map=slide_color_map)

        

    for j in range(2):
        if i + j >= len(plants):
            break

        plant = plants[i + j]
        plant_df = df_long[df_long["Plant"] == plant]

        summary = (
            plant_df.groupby("Defect")["Rejected"]
            .sum().sort_values(ascending=False)
        )

        img_path = f"charts/{plant.replace('/', '_').replace(',', '')}.png"
        plot_clean_donut(summary, plant, img_path, color_map=slide_color_map)
        
        # shifting right or second chart to right
        left = Inches(1) if j == 0 else Inches(6)
        top = Inches(1.2)

        add_plant_title(
            slide,
            plant,
            left,
            Inches(0.6),
            Inches(4.8)
        )

        # Add chart image
        slide.shapes.add_picture(
            img_path,
            left,
            Inches(1.1),
            width=Inches(4.8),
            height=Inches(3.8)   # ✅ HARD HEIGHT LIMIT
        )

        # --- Calculate cumulative stats ---
        inspected, rejected, rej_pct = get_cumulative_stats(df, plant)

        # --- Add summary table below chart ---
        add_summary_table(
            slide,
            left,
            Inches(5),   # below chart
            inspected,
            rejected,
            rej_pct
        )

        add_month_wise_table(
            slide,
            left,
            Inches(6),   # below cumulative table
            plant,
            df
        )



# ================== ADD SUMMARY SLIDES ==================

add_erc_overview_slide(prs, base_slide, df, df_long)
add_spider_chart_slide(prs, df)
add_best_worst_slides(prs, base_slide, df)
add_pareto_slide(prs, base_slide, df_long)

add_top_performing_units_slide(prs, base_slide, df)

# ================== REMOVE TEMPLATE SLIDE (ONCE) ==================
prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

# ================== SAVE (ONCE) ==================
prs.save(output_ppt)

print(" FINAL RITES PPT created:", output_ppt)



