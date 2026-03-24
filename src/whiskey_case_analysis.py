import csv
import math
import re
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

import matplotlib.pyplot as plt
import numpy as np
from scipy.cluster.vq import kmeans2
from scipy.spatial.distance import pdist, squareform


# ── File paths ─────────────────────────────────────────────────────────────────
ROOT      = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "supporting_documents" / "scotch (1).xlsx"
EXPORTS   = ROOT / "exports"
FIGURES   = EXPORTS / "figures"
TABLES    = EXPORTS / "tables"

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


# ── Cluster labels and descriptions ───────────────────────────────────────────
CLUSTER_NAMES = {
    0: "Classic Smoky-Sweet",
    1: "Sherried & Sweet",
    2: "Light & Approachable",
    3: "Herbal & Crisp",
    4: "Maritime Peat",
}

CLUSTER_DESCRIPTIONS = {
    0: "Medium-bodied pours that balance smoke and sweetness.",
    1: "Fruit-forward, sherried, smoother malts suited to gifting and premium browsing.",
    2: "Lighter, smoother entry points that help new customers trade up into single malts.",
    3: "Leaner grassy-fruity whiskies for curious customers and rotating tasting flights.",
    4: "Coastal, salty, peated bottles that create identity and tasting-event buzz.",
}


# ── Recommended assortment mixes ──────────────────────────────────────────────
LAUNCH_MIX_24 = {
    "Sherried & Sweet":     6,
    "Classic Smoky-Sweet":  5,
    "Light & Approachable": 5,
    "Herbal & Crisp":       4,
    "Maritime Peat":        4,
}

LEAN_MIX_18 = {
    "Sherried & Sweet":     5,
    "Classic Smoky-Sweet":  4,
    "Light & Approachable": 3,
    "Herbal & Crisp":       3,
    "Maritime Peat":        3,
}


# ══════════════════════════════════════════════════════════════════════════════
# Step 1 – Read the Excel file
# Excel (.xlsx) files are ZIP archives containing XML inside.
# We read the XML directly so we don't need pandas or openpyxl.
# ══════════════════════════════════════════════════════════════════════════════

def col_letter_to_number(col):
    """Convert a spreadsheet column letter (A, B … Z, AA …) to an integer."""
    value = 0
    for ch in col:
        value = value * 26 + ord(ch) - 64
    return value


def make_field_name(group, sub, col):
    """Build a clean Python-friendly field name from the two header rows."""
    if col == "A":
        return "name"
    if col == "B":
        return "dist_short"
    cleaned = (
        sub.lower()
        .replace("%", "pct")
        .replace(" ", "_")
        .replace(".", "")
        .replace("/", "_")
    )
    if group:
        return f"{group.lower()}_{cleaned}"
    return cleaned


def load_workbook_records():
    """
    Open the Excel workbook and return one dict per whiskey row.
    Also returns a small metadata dict for printing at the end.
    """
    with ZipFile(DATA_PATH) as zf:

        # Step 1a – Load Excel's shared string table.
        # Excel doesn't store repeated strings inline; it indexes them here.
        shared_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        shared_strings = []
        for si in shared_root.findall("main:si", NS):
            text = "".join(t.text or "" for t in si.iterfind(".//main:t", NS))
            shared_strings.append(text)

        # Step 1b – Get the name of the first worksheet.
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet_name = workbook.find("main:sheets/main:sheet", NS).attrib["name"]

        # Step 1c – Read every cell from the first sheet into a list of dicts.
        sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        rows = []
        for row in sheet.find("main:sheetData", NS).findall("main:row", NS):
            row_values = {}
            for cell in row.findall("main:c", NS):
                ref   = cell.attrib["r"]
                col   = re.match(r"[A-Z]+", ref).group(0)
                v_tag = cell.find("main:v", NS)
                value = "" if v_tag is None else (v_tag.text or "")
                # If the cell holds a shared-string index, look up the real text
                if cell.attrib.get("t") == "s" and value:
                    value = shared_strings[int(value)]
                row_values[col] = value
            rows.append(row_values)

    # Step 1d – Build a (column_letter → field_name) mapping from the two header rows.
    all_cols = sorted(
        set().union(*[set(r.keys()) for r in rows]),
        key=col_letter_to_number,
    )
    header_group = rows[0]   # Row 1: broad category (e.g. "NOSE", "BODY")
    header_sub   = rows[1]   # Row 2: specific descriptor (e.g. "Smoky", "Honey")
    fields = []
    for col in all_cols:
        group = header_group.get(col, "").strip()
        sub   = header_sub.get(col, "").strip()
        fields.append((col, make_field_name(group, sub, col)))

    # Step 1e – Build one dict per whiskey, skipping blank rows.
    raw_records = []
    blank_count = 0
    for row in rows[2:]:
        if not row.get("A", "").strip():
            blank_count += 1
            continue
        record = {name: row.get(col, "") for col, name in fields}
        raw_records.append(record)

    metadata = {
        "sheet_name":         sheet_name,
        "field_names":        [name for _, name in fields],
        "row_count_raw":      len(rows) - 2,
        "blank_rows_removed": blank_count,
    }
    return raw_records, metadata


# ══════════════════════════════════════════════════════════════════════════════
# Step 2 – Convert string values to the right Python types
# ══════════════════════════════════════════════════════════════════════════════

def convert_records(raw_records):
    """
    Turn the raw string dicts into properly typed dicts:
    - Sensory descriptor columns → int  (0 or 1 flags)
    - Numeric columns (age, score, pct, dist) → float
    - Everything else stays as a string
    """
    numeric_fields = {"age", "dist", "score", "pct"}
    region_flags   = {"islay", "midland", "spey", "east", "west",
                      "north", "lowland", "campbell", "islands"}

    cleaned = []
    for record in raw_records:
        item = {}
        for key, value in record.items():
            value = value.strip() if isinstance(value, str) else value

            if key in numeric_fields:
                if value == "":
                    item[key] = math.nan
                else:
                    number = float(value)
                    # A negative age is a data error — treat as missing
                    item[key] = math.nan if key == "age" and number < 0 else number

            elif key.startswith(("color_", "nose_", "body_", "pal_", "fin_")) or key in region_flags:
                item[key] = int(value or 0)

            else:
                item[key] = value

        cleaned.append(item)
    return cleaned


# ══════════════════════════════════════════════════════════════════════════════
# Step 3 – Cluster the 109 whiskies into 5 flavour groups
# ══════════════════════════════════════════════════════════════════════════════

def silhouette_score(distance_matrix, labels):
    """
    Measure how well-separated the clusters are.
    Score ranges from -1 (bad) to 1 (perfect). Above 0.2 is usable.

    For each whiskey we compute:
      a = average distance to others in the same cluster
      b = average distance to the nearest different cluster
      score = (b - a) / max(a, b)
    The overall score is the mean across all whiskies.
    """
    labels        = np.asarray(labels)
    unique_labels = np.unique(labels)
    scores        = []

    for idx in range(len(labels)):
        # Average distance to other members of the same cluster
        same       = labels == labels[idx]
        same[idx]  = False
        a_i        = distance_matrix[idx, same].mean() if same.sum() else 0.0

        # Average distance to each other cluster; keep the nearest one
        between = []
        for other in unique_labels:
            if other == labels[idx]:
                continue
            mask = labels == other
            if mask.sum():
                between.append(distance_matrix[idx, mask].mean())

        if not between:
            scores.append(0.0)
            continue

        b_i   = min(between)
        denom = max(a_i, b_i)
        scores.append(0.0 if denom == 0 else (b_i - a_i) / denom)

    return float(np.mean(scores))


def run_cluster_analysis(records):
    """
    Run k-means (k=5) on the sensory descriptor columns.
    Try 50 different random starts and keep the run with the lowest SSE
    (sum of squared distances), provided every cluster has at least 5 members.
    """
    # Grab only the sensory descriptor columns
    sensory_fields = [
        key for key in records[0].keys()
        if key.startswith(("color_", "nose_", "body_", "pal_", "fin_"))
    ]

    # Build a 2-D numeric array: rows = whiskies, cols = sensory features
    x = np.array(
        [[record[field] for field in sensory_fields] for record in records],
        dtype=float,
    )

    best_sse    = None
    best_labels = None

    for seed in range(50):
        np.random.seed(seed)
        centroids, labels = kmeans2(x, 5, minit="points", iter=100)

        # Skip if any cluster ended up too small
        if min(Counter(labels).values()) < 5:
            continue

        sse = float(((x - centroids[labels]) ** 2).sum())
        if best_sse is None or sse < best_sse:
            best_sse    = sse
            best_labels = labels.copy()

    if best_labels is None:
        raise RuntimeError("Could not find a stable 5-cluster solution.")

    # Compute silhouette score on the winning solution
    distance_matrix = squareform(pdist(x, metric="euclidean"))
    score = silhouette_score(distance_matrix, best_labels)

    return x, best_labels.astype(int), sensory_fields, score


# ══════════════════════════════════════════════════════════════════════════════
# Step 4 – Reduce to 2D with PCA so we can draw a scatter plot
# ══════════════════════════════════════════════════════════════════════════════

def pca_2d(x):
    """
    Project the high-dimensional sensory data down to 2 dimensions using PCA.
    We centre the data, then use SVD to find the two directions of most spread.
    """
    centered    = x - x.mean(axis=0)
    u, s, _     = np.linalg.svd(centered, full_matrices=False)
    return u[:, :2] * s[:2]


# ══════════════════════════════════════════════════════════════════════════════
# Step 5 – Compute summary statistics for each cluster
# ══════════════════════════════════════════════════════════════════════════════

def summarize_clusters(records, x, labels, sensory_fields):
    """Build one summary dict per cluster with counts, averages, and top features."""
    summaries = []
    for cluster_id in sorted(set(labels.tolist())):
        idx             = np.where(labels == cluster_id)[0]
        cluster_records = [records[i] for i in idx]
        cluster_x       = x[idx]

        # Top 6 sensory features by mean value in this cluster
        mean_vector  = cluster_x.mean(axis=0)
        top_idx      = np.argsort(mean_vector)[::-1][:6]
        top_features = [sensory_fields[i] for i in top_idx]

        ages = [r["age"] for r in cluster_records if not math.isnan(r["age"])]

        summaries.append({
            "cluster_id":    cluster_id,
            "cluster_name":  CLUSTER_NAMES[cluster_id],
            "description":   CLUSTER_DESCRIPTIONS[cluster_id],
            "count":         len(cluster_records),
            "avg_score":     round(float(np.mean([r["score"] for r in cluster_records])), 2),
            "avg_pct":       round(float(np.mean([r["pct"]   for r in cluster_records])), 2),
            "avg_age_known": round(float(np.mean(ages)), 2) if ages else "",
            "high_share":    sum(1 for r in cluster_records if r["region"].strip() == "HIGH"),
            "low_share":     sum(1 for r in cluster_records if r["region"].strip() == "LOW"),
            "islay_share":   sum(1 for r in cluster_records if r["region"].strip() == "ISLAY"),
            "top_features":  ", ".join(top_features),
        })
    return summaries


# ══════════════════════════════════════════════════════════════════════════════
# Step 6 – Attach cluster info to each whiskey record
# ══════════════════════════════════════════════════════════════════════════════

def add_cluster_fields(records, labels, coords):
    """
    Add cluster_id, cluster_name, and 2-D PCA coordinates (pc1, pc2)
    to every whiskey record so they show up in the exported CSV.
    """
    enriched = []
    for record, cluster_id, (pc1, pc2) in zip(records, labels, coords):
        item = dict(record)
        item["cluster_id"]   = int(cluster_id)
        item["cluster_name"] = CLUSTER_NAMES[int(cluster_id)]
        item["pc1"]          = round(float(pc1), 4)
        item["pc2"]          = round(float(pc2), 4)
        enriched.append(item)
    return enriched


# ══════════════════════════════════════════════════════════════════════════════
# Step 7 – Save CSV tables
# ══════════════════════════════════════════════════════════════════════════════

def write_csv(path, rows, fieldnames):
    """Write a list of dicts to a CSV file, creating parent folders if needed."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def save_tables(records, summaries):
    """
    Export four CSV tables to exports/tables/:
      scotch_clustered.csv       – every whiskey with its cluster assignment
      cluster_summary.csv        – one row per cluster with stats
      recommended_mix.csv        – 24-SKU and 18-SKU assortment options
      cluster_representatives.csv – top 5 bottles per cluster by score
    """
    # Table 1: Full clustered dataset
    write_csv(TABLES / "scotch_clustered.csv", records, list(records[0].keys()))

    # Table 2: Cluster summary statistics
    summary_fields = [
        "cluster_id", "cluster_name", "description", "count",
        "avg_score", "avg_pct", "avg_age_known",
        "high_share", "low_share", "islay_share", "top_features",
    ]
    write_csv(TABLES / "cluster_summary.csv", summaries, summary_fields)

    # Table 3: Recommended mix (both size options in one file)
    mix_rows = []
    for name, count in LAUNCH_MIX_24.items():
        mix_rows.append({"option": "Launch assortment (24 SKUs)", "cluster_name": name, "recommended_skus": count})
    for name, count in LEAN_MIX_18.items():
        mix_rows.append({"option": "Lean assortment (18 SKUs)", "cluster_name": name, "recommended_skus": count})
    write_csv(TABLES / "recommended_mix.csv", mix_rows, ["option", "cluster_name", "recommended_skus"])

    # Table 4: Top 5 representative bottles per cluster, ranked by score
    rep_rows = []
    for cluster_name in CLUSTER_NAMES.values():
        candidates = [r for r in records if r["cluster_name"] == cluster_name]
        top_5      = sorted(candidates, key=lambda r: r["score"], reverse=True)[:5]
        for r in top_5:
            rep_rows.append({
                "cluster_name": cluster_name,
                "name":         r["name"],
                "score":        r["score"],
                "region":       r["region"],
                "district":     r["district"],
                "pct":          r["pct"],
            })
    write_csv(
        TABLES / "cluster_representatives.csv",
        rep_rows,
        ["cluster_name", "name", "score", "region", "district", "pct"],
    )


# ══════════════════════════════════════════════════════════════════════════════
# Step 8 – Save PNG charts
# ══════════════════════════════════════════════════════════════════════════════

def save_figures(records, summaries):
    """
    Export three PNG charts to exports/figures/:
      cluster_scatter.png    – PCA scatter plot coloured by cluster
      cluster_size_score.png – bar chart of cluster sizes with score overlay
      recommended_mix.png    – horizontal bar of recommended SKU counts
    """
    FIGURES.mkdir(parents=True, exist_ok=True)

    # One colour per cluster — used in all three charts
    colors = {
        "Classic Smoky-Sweet":  "#7a4f01",
        "Sherried & Sweet":     "#9e2a2b",
        "Light & Approachable": "#d4a017",
        "Herbal & Crisp":       "#4f772d",
        "Maritime Peat":        "#1d3557",
    }

    plt.style.use("ggplot")

    # Chart 1: PCA scatter — each dot is one whiskey, coloured by cluster
    fig, ax = plt.subplots(figsize=(10, 6))
    for name in CLUSTER_NAMES.values():
        cluster_rows = [r for r in records if r["cluster_name"] == name]
        ax.scatter(
            [r["pc1"] for r in cluster_rows],
            [r["pc2"] for r in cluster_rows],
            s=55, alpha=0.8, label=name, color=colors[name],
        )
        # Label each cluster at its centroid
        ax.text(
            float(np.mean([r["pc1"] for r in cluster_rows])),
            float(np.mean([r["pc2"] for r in cluster_rows])),
            name, fontsize=9, weight="bold", ha="center", va="center",
            bbox={"boxstyle": "round,pad=0.2", "facecolor": "white", "alpha": 0.7, "edgecolor": "none"},
        )
    ax.set_title("Whiskey Styles Clustered on Sensory Descriptors")
    ax.set_xlabel("Principal Component 1")
    ax.set_ylabel("Principal Component 2")
    ax.legend(frameon=True, fontsize=8)
    fig.tight_layout()
    fig.savefig(FIGURES / "cluster_scatter.png", dpi=200)
    plt.close(fig)

    # Chart 2: Cluster size (bars) + average quality score (line overlay)
    fig, ax1 = plt.subplots(figsize=(10, 6))
    ordered = sorted(summaries, key=lambda r: r["count"], reverse=True)
    names   = [r["cluster_name"] for r in ordered]
    counts  = [r["count"]        for r in ordered]
    scores  = [r["avg_score"]    for r in ordered]
    ax1.bar(names, counts, color=[colors[n] for n in names], alpha=0.85)
    ax1.set_ylabel("Number of whiskies in dataset")
    ax1.set_title("Cluster Sizes and Average Quality Score")
    ax1.tick_params(axis="x", rotation=20)
    ax2 = ax1.twinx()   # Second y-axis for the score line
    ax2.plot(names, scores, color="#111111", marker="o", linewidth=2)
    ax2.set_ylabel("Average score")
    fig.tight_layout()
    fig.savefig(FIGURES / "cluster_size_score.png", dpi=200)
    plt.close(fig)

    # Chart 3: Recommended opening assortment — horizontal bars
    fig, ax = plt.subplots(figsize=(10, 6))
    launch_names  = list(LAUNCH_MIX_24.keys())
    launch_counts = list(LAUNCH_MIX_24.values())
    ax.barh(launch_names, launch_counts, color=[colors[n] for n in launch_names], alpha=0.9)
    for idx, value in enumerate(launch_counts):
        ax.text(value + 0.1, idx, str(value), va="center", fontsize=9)
    ax.set_xlabel("Recommended launch SKUs")
    ax.set_title("Recommended Initial Scotch Assortment: 24 SKUs")
    fig.tight_layout()
    fig.savefig(FIGURES / "recommended_mix.png", dpi=200)
    plt.close(fig)


# ══════════════════════════════════════════════════════════════════════════════
# Main – run all steps in order
# ══════════════════════════════════════════════════════════════════════════════

def main():
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Expected workbook at {DATA_PATH}")

    # Step 1: Load the Excel file
    raw_records, metadata = load_workbook_records()

    # Step 2: Convert string values to numbers
    records = convert_records(raw_records)

    # Step 3: Cluster the whiskies into 5 groups
    x, labels, sensory_fields, silhouette = run_cluster_analysis(records)

    # Step 4: Reduce to 2D for scatter plotting
    coords = pca_2d(x)

    # Step 5: Compute per-cluster summary statistics
    summaries = summarize_clusters(records, x, labels, sensory_fields)

    # Step 6: Attach cluster info to each whiskey record
    enriched = add_cluster_fields(records, labels, coords)

    # Step 7: Save CSV tables
    save_tables(enriched, summaries)

    # Step 8: Save PNG charts
    save_figures(enriched, summaries)

    print("Sheet:",               metadata["sheet_name"])
    print("Rows loaded:",         len(raw_records))
    print("Blank rows removed:",  metadata["blank_rows_removed"])
    print("Fields:",              len(metadata["field_names"]))
    print("Silhouette:",          round(silhouette, 4))
    print("Cluster counts:",      dict(sorted(Counter(labels.tolist()).items())))
    print("Exports written to:",  EXPORTS)


if __name__ == "__main__":
    main()
