from __future__ import annotations

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


ROOT = Path(__file__).resolve().parents[1]
SUPPORTING_DOCS = ROOT / "supporting_documents"
DATA_PATH = SUPPORTING_DOCS / "scotch (1).xlsx"
EXPORTS = ROOT / "exports"
FIGURES = EXPORTS / "figures"
TABLES = EXPORTS / "tables"
NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

CLUSTER_NAME_MAP = {
    0: "Classic Smoky-Sweet",
    1: "Sherried & Sweet",
    2: "Light & Approachable",
    3: "Herbal & Crisp",
    4: "Maritime Peat",
}

CLUSTER_DESCRIPTION_MAP = {
    0: "Medium-bodied pours that balance smoke and sweetness.",
    1: "Fruit-forward, sherried, smoother malts suited to gifting and premium browsing.",
    2: "Lighter, smoother entry points that help new customers trade up into single malts.",
    3: "Leaner grassy-fruity whiskies for curious customers and rotating tasting flights.",
    4: "Coastal, salty, peated bottles that create identity and tasting-event buzz.",
}

LAUNCH_MIX_24 = {
    "Sherried & Sweet": 6,
    "Classic Smoky-Sweet": 5,
    "Light & Approachable": 5,
    "Herbal & Crisp": 4,
    "Maritime Peat": 4,
}

LEAN_MIX_18 = {
    "Sherried & Sweet": 5,
    "Classic Smoky-Sweet": 4,
    "Light & Approachable": 3,
    "Herbal & Crisp": 3,
    "Maritime Peat": 3,
}


def col_to_num(col: str) -> int:
    value = 0
    for ch in col:
        value = value * 26 + ord(ch) - 64
    return value


def clean_field_name(group: str, sub: str, col: str) -> str:
    if col == "A":
        return "name"
    if col == "B":
        return "dist_short"
    cleaned_sub = (
        sub.lower()
        .replace("%", "pct")
        .replace(" ", "_")
        .replace(".", "")
        .replace("/", "_")
    )
    if group:
        return f"{group.lower()}_{cleaned_sub}"
    return cleaned_sub


def load_workbook_records() -> tuple[list[dict[str, str]], dict[str, object]]:
    with ZipFile(DATA_PATH) as zf:
        shared_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        shared_strings = [
            "".join(t.text or "" for t in si.iterfind(".//main:t", NS))
            for si in shared_root.findall("main:si", NS)
        ]

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet_name = workbook.find("main:sheets/main:sheet", NS).attrib["name"]

        sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        rows: list[dict[str, str]] = []
        for row in sheet.find("main:sheetData", NS).findall("main:row", NS):
            row_values: dict[str, str] = {}
            for cell in row.findall("main:c", NS):
                ref = cell.attrib["r"]
                col = re.match(r"[A-Z]+", ref).group(0)
                value_tag = cell.find("main:v", NS)
                value = "" if value_tag is None else value_tag.text or ""
                if cell.attrib.get("t") == "s" and value:
                    value = shared_strings[int(value)]
                row_values[col] = value
            rows.append(row_values)

    all_cols = sorted(set().union(*[set(row.keys()) for row in rows]), key=col_to_num)
    header_group = rows[0]
    header_sub = rows[1]
    fields = [
        (col, clean_field_name(header_group.get(col, "").strip(), header_sub.get(col, "").strip(), col))
        for col in all_cols
    ]

    raw_records = []
    blank_rows_removed = 0
    for row in rows[2:]:
        if not row.get("A", "").strip():
            blank_rows_removed += 1
            continue
        raw_records.append({name: row.get(col, "") for col, name in fields})

    metadata = {
        "sheet_name": sheet_name,
        "field_names": [name for _, name in fields],
        "row_count_raw": len(rows) - 2,
        "blank_rows_removed": blank_rows_removed,
    }
    return raw_records, metadata


def convert_records(raw_records: list[dict[str, str]]) -> list[dict[str, object]]:
    numeric_fields = {"age", "dist", "score", "pct"}
    region_flags = {"islay", "midland", "spey", "east", "west", "north", "lowland", "campbell", "islands"}

    cleaned: list[dict[str, object]] = []
    for record in raw_records:
        item: dict[str, object] = {}
        for key, value in record.items():
            value = value.strip() if isinstance(value, str) else value
            if key in numeric_fields:
                if value == "":
                    item[key] = math.nan
                else:
                    number = float(value)
                    item[key] = math.nan if key == "age" and number < 0 else number
            elif key.startswith(("color_", "nose_", "body_", "pal_", "fin_")) or key in region_flags:
                item[key] = int(value or 0)
            else:
                item[key] = value
        cleaned.append(item)
    return cleaned


def silhouette_score_from_distance(distance_matrix: np.ndarray, labels: np.ndarray) -> float:
    labels = np.asarray(labels)
    unique_labels = np.unique(labels)
    scores = []
    for idx in range(len(labels)):
        same_cluster = labels == labels[idx]
        same_cluster[idx] = False
        a_i = distance_matrix[idx, same_cluster].mean() if same_cluster.sum() else 0.0
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
        b_i = min(between)
        denom = max(a_i, b_i)
        scores.append(0.0 if denom == 0 else (b_i - a_i) / denom)
    return float(np.mean(scores))


def run_cluster_analysis(records: list[dict[str, object]]) -> tuple[np.ndarray, np.ndarray, list[str], float]:
    sensory_fields = [
        key
        for key in records[0].keys()
        if key.startswith(("color_", "nose_", "body_", "pal_", "fin_"))
    ]
    x = np.array([[record[field] for field in sensory_fields] for record in records], dtype=float)

    best_solution: tuple[float, np.ndarray, np.ndarray] | None = None
    for seed in range(50):
        np.random.seed(seed)
        centroids, labels = kmeans2(x, 5, minit="points", iter=100)
        cluster_sizes = Counter(labels)
        if min(cluster_sizes.values()) < 5:
            continue
        sse = float(((x - centroids[labels]) ** 2).sum())
        if best_solution is None or sse < best_solution[0]:
            best_solution = (sse, centroids, labels)

    if best_solution is None:
        raise RuntimeError("Unable to find a stable 5-cluster solution.")

    _, _, labels = best_solution
    distance_matrix = squareform(pdist(x, metric="euclidean"))
    silhouette = silhouette_score_from_distance(distance_matrix, labels)
    return x, labels.astype(int), sensory_fields, silhouette


def pca_coordinates(x: np.ndarray) -> np.ndarray:
    centered = x - x.mean(axis=0)
    u, s, _ = np.linalg.svd(centered, full_matrices=False)
    return u[:, :2] * s[:2]


def cluster_summary(
    records: list[dict[str, object]],
    x: np.ndarray,
    labels: np.ndarray,
    sensory_fields: list[str],
) -> list[dict[str, object]]:
    summaries: list[dict[str, object]] = []
    for cluster_id in sorted(set(labels.tolist())):
        idx = np.where(labels == cluster_id)[0]
        cluster_records = [records[i] for i in idx]
        cluster_x = x[idx]
        mean_vector = cluster_x.mean(axis=0)
        top_feature_idx = np.argsort(mean_vector)[::-1][:6]
        top_features = [sensory_fields[i] for i in top_feature_idx]
        ages = [record["age"] for record in cluster_records if not math.isnan(record["age"])]
        summaries.append(
            {
                "cluster_id": cluster_id,
                "cluster_name": CLUSTER_NAME_MAP[cluster_id],
                "description": CLUSTER_DESCRIPTION_MAP[cluster_id],
                "count": len(cluster_records),
                "avg_score": round(float(np.mean([record["score"] for record in cluster_records])), 2),
                "avg_pct": round(float(np.mean([record["pct"] for record in cluster_records])), 2),
                "avg_age_known": round(float(np.mean(ages)), 2) if ages else "",
                "high_share": sum(1 for record in cluster_records if record["region"].strip() == "HIGH"),
                "low_share": sum(1 for record in cluster_records if record["region"].strip() == "LOW"),
                "islay_share": sum(1 for record in cluster_records if record["region"].strip() == "ISLAY"),
                "top_features": ", ".join(top_features),
            }
        )
    return summaries


def add_cluster_fields(records: list[dict[str, object]], labels: np.ndarray, coords: np.ndarray) -> list[dict[str, object]]:
    enriched = []
    for record, cluster_id, (pc1, pc2) in zip(records, labels, coords):
        item = dict(record)
        item["cluster_id"] = int(cluster_id)
        item["cluster_name"] = CLUSTER_NAME_MAP[int(cluster_id)]
        item["pc1"] = round(float(pc1), 4)
        item["pc2"] = round(float(pc2), 4)
        enriched.append(item)
    return enriched


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def save_tables(records: list[dict[str, object]], summaries: list[dict[str, object]]) -> None:
    write_csv(TABLES / "scotch_clustered.csv", records, list(records[0].keys()))

    summary_fields = [
        "cluster_id",
        "cluster_name",
        "description",
        "count",
        "avg_score",
        "avg_pct",
        "avg_age_known",
        "high_share",
        "low_share",
        "islay_share",
        "top_features",
    ]
    write_csv(TABLES / "cluster_summary.csv", summaries, summary_fields)

    recommendation_rows = []
    for cluster_name, count in LAUNCH_MIX_24.items():
        recommendation_rows.append({"option": "Launch assortment (24 SKUs)", "cluster_name": cluster_name, "recommended_skus": count})
    for cluster_name, count in LEAN_MIX_18.items():
        recommendation_rows.append({"option": "Lean assortment (18 SKUs)", "cluster_name": cluster_name, "recommended_skus": count})
    write_csv(TABLES / "recommended_mix.csv", recommendation_rows, ["option", "cluster_name", "recommended_skus"])

    representative_rows = []
    for cluster_name in CLUSTER_NAME_MAP.values():
        candidates = [row for row in records if row["cluster_name"] == cluster_name]
        top_candidates = sorted(candidates, key=lambda row: row["score"], reverse=True)[:5]
        for row in top_candidates:
            representative_rows.append(
                {
                    "cluster_name": cluster_name,
                    "name": row["name"],
                    "score": row["score"],
                    "region": row["region"],
                    "district": row["district"],
                    "pct": row["pct"],
                }
            )
    write_csv(TABLES / "cluster_representatives.csv", representative_rows, ["cluster_name", "name", "score", "region", "district", "pct"])


def save_figures(records: list[dict[str, object]], summaries: list[dict[str, object]]) -> None:
    FIGURES.mkdir(parents=True, exist_ok=True)
    colors = {
        "Classic Smoky-Sweet": "#7a4f01",
        "Sherried & Sweet": "#9e2a2b",
        "Light & Approachable": "#d4a017",
        "Herbal & Crisp": "#4f772d",
        "Maritime Peat": "#1d3557",
    }

    plt.style.use("ggplot")

    fig, ax = plt.subplots(figsize=(10, 6))
    for cluster_name in CLUSTER_NAME_MAP.values():
        cluster_rows = [row for row in records if row["cluster_name"] == cluster_name]
        ax.scatter(
            [row["pc1"] for row in cluster_rows],
            [row["pc2"] for row in cluster_rows],
            s=55,
            alpha=0.8,
            label=cluster_name,
            color=colors[cluster_name],
        )
        ax.text(
            float(np.mean([row["pc1"] for row in cluster_rows])),
            float(np.mean([row["pc2"] for row in cluster_rows])),
            cluster_name,
            fontsize=9,
            weight="bold",
            ha="center",
            va="center",
            bbox={"boxstyle": "round,pad=0.2", "facecolor": "white", "alpha": 0.7, "edgecolor": "none"},
        )
    ax.set_title("Whiskey Styles Clustered on Sensory Descriptors")
    ax.set_xlabel("Principal Component 1")
    ax.set_ylabel("Principal Component 2")
    ax.legend(frameon=True, fontsize=8)
    fig.tight_layout()
    fig.savefig(FIGURES / "cluster_scatter.png", dpi=200)
    plt.close(fig)

    fig, ax1 = plt.subplots(figsize=(10, 6))
    ordered = sorted(summaries, key=lambda row: row["count"], reverse=True)
    names = [row["cluster_name"] for row in ordered]
    counts = [row["count"] for row in ordered]
    scores = [row["avg_score"] for row in ordered]
    ax1.bar(names, counts, color=[colors[name] for name in names], alpha=0.85)
    ax1.set_ylabel("Number of whiskies in dataset")
    ax1.set_title("Cluster Sizes and Average Quality Score")
    ax1.tick_params(axis="x", rotation=20)
    ax2 = ax1.twinx()
    ax2.plot(names, scores, color="#111111", marker="o", linewidth=2)
    ax2.set_ylabel("Average score")
    fig.tight_layout()
    fig.savefig(FIGURES / "cluster_size_score.png", dpi=200)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 6))
    launch_names = list(LAUNCH_MIX_24.keys())
    launch_counts = list(LAUNCH_MIX_24.values())
    ax.barh(launch_names, launch_counts, color=[colors[name] for name in launch_names], alpha=0.9)
    for idx, value in enumerate(launch_counts):
        ax.text(value + 0.1, idx, str(value), va="center", fontsize=9)
    ax.set_xlabel("Recommended launch SKUs")
    ax.set_title("Recommended Initial Scotch Assortment: 24 SKUs")
    fig.tight_layout()
    fig.savefig(FIGURES / "recommended_mix.png", dpi=200)
    plt.close(fig)


def main() -> None:
    if not DATA_PATH.exists():
        raise FileNotFoundError(f"Expected workbook at {DATA_PATH}")

    raw_records, metadata = load_workbook_records()
    records = convert_records(raw_records)
    x, labels, sensory_fields, silhouette = run_cluster_analysis(records)
    coords = pca_coordinates(x)
    enriched = add_cluster_fields(records, labels, coords)
    summaries = cluster_summary(records, x, labels, sensory_fields)
    save_tables(enriched, summaries)
    save_figures(enriched, summaries)

    print("Sheet:", metadata["sheet_name"])
    print("Rows loaded:", len(raw_records))
    print("Blank rows removed:", metadata["blank_rows_removed"])
    print("Fields:", len(metadata["field_names"]))
    print("Silhouette:", round(silhouette, 4))
    print("Cluster counts:", dict(sorted(Counter(labels.tolist()).items())))
    print("Exports written to:", EXPORTS)


if __name__ == "__main__":
    main()
