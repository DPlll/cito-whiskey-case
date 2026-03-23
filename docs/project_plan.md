# Whiskey Case Project Plan

## Objective
Create a 5-minute, presentation-ready business analytics recommendation for Kathryn Finnegan as she launches a boutique wine and spirits shop near Harvard Square.

## Scope
- Unzip `ScotchData 4 (1).zip` and use the included readme files.
- Analyze `scotch.xlsx` directly as the primary dataset.
- Build an explainable cluster analysis to segment the whiskies into usable store-assortment groups.
- Translate the clustering output into a stocking recommendation and a next-step analytics plan.
- Produce slide-ready content, exported charts, and supporting analysis files.

## Deliverables
- `docs/data_summary.md`
- `docs/clustering_method.md`
- `docs/recommendations.md`
- `docs/additional_data_sources.md`
- `docs/slide_outline.md`
- `presentation/whiskey_case_slide_content.md`
- `src/whiskey_case_analysis.py`
- `exports/figures/*`
- `exports/tables/*`

## Work Approach
1. Parse the workbook directly from the `.xlsx` XML structure.
2. Clean the data, treating blank rows as noise and `-9` age values as missing.
3. Cluster on the 68 sensory descriptors so the grouping stays easy to explain.
4. Interpret clusters in customer-facing business language.
5. Recommend a practical launch assortment and next analytics tasks for Kathryn.
