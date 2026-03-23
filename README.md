# The Whiskey Case

Business analytics case project focused on developing an opening scotch assortment recommendation for Kathryn Finnegan's boutique wine and spirits shop near Harvard Square in Cambridge, Massachusetts.

## Project Objective

The project uses the `scotch.xlsx` dataset from the supplied `ScotchData` archive to:

- analyze whiskey style differences using cluster analysis
- recommend a practical opening assortment
- identify the next analytics task after initial stocking
- prepare presentation-ready outputs for a short business presentation

## Data Used

Primary analysis dataset:
- `supporting_documents/scotch (1).xlsx`

Supporting source documents:
- `supporting_documents/whiskey_case (1).pdf`
- `supporting_documents/ScotchData 4 (1).zip`
- `supporting_documents/ScotchData_unzipped/...`

The analysis model uses the workbook as the primary input. The PDF and readme files are included as supporting assignment and data-reference materials.

## Repository Structure

- `src/`
  - analysis script
- `docs/`
  - project notes, data summary, method summary, recommendations, and slide outline
- `presentation/`
  - slide-ready presentation content
- `exports/figures/`
  - presentation graphics
- `exports/tables/`
  - cleaned and clustered output tables
- `supporting_documents/`
  - source workbook, case instructions, archive, and extracted reference files

## How to Run

One-click option:

```powershell
run_me.bat
```

Direct script option:

```powershell
python src/whiskey_case_analysis.py
```

## Main Outputs

Presentation content:
- `presentation/whiskey_case_slide_content.md`

Key figures:
- `exports/figures/recommended_mix.png`
- `exports/figures/cluster_scatter.png`
- `exports/figures/cluster_size_score.png`

Key tables:
- `exports/tables/cluster_summary.csv`
- `exports/tables/recommended_mix.csv`
- `exports/tables/cluster_representatives.csv`
- `exports/tables/scotch_clustered.csv`

Supporting writeups:
- `docs/project_plan.md`
- `docs/data_summary.md`
- `docs/clustering_method.md`
- `docs/recommendations.md`
- `docs/additional_data_sources.md`
- `docs/slide_outline.md`

## Method Summary

The analysis script parses the workbook directly, cleans the source records, and clusters the whiskies using the sensory descriptor fields from the dataset. The outputs are designed to support a concise presentation rather than a large technical workflow.

## Notes

- The project is intentionally lightweight and presentation-oriented.
- Generated figures and tables can be refreshed at any time by rerunning the analysis script.
