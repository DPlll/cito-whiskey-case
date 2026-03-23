# Clustering Method

## Business Objective
Segment the whiskies into a small number of style groups that Kathryn can use to build a balanced opening assortment and explain that assortment to customers.

## Method Chosen
- Method: `k`-means clustering
- Input variables: 68 binary sensory descriptors from color, nose, body, palate, and finish
- Distance logic: Euclidean distance on the binary sensory matrix
- Validation used: compared small cluster-count options and selected the option that produced the most usable segmentation for a retail assortment

## Why This Method
- It is easy to explain in class: whiskies were grouped by shared tasting-profile attributes.
- It avoids overfitting a small dataset.
- It uses the most relevant variables for an opening assortment: flavor and style.
- It keeps the recommendation centered on customer-facing product differences instead of geography alone.

## Why Five Clusters
- Tested cluster counts between 3 and 7.
- A 5-cluster solution gave the best balance between:
  - interpretability
  - enough variety for stocking decisions
  - avoiding a single oversized cluster
- The silhouette value is modest because the dataset is sparse and many whiskies overlap stylistically, but five clusters were still the most managerially useful segmentation.

## Cleaning Decisions
- Removed one blank row.
- Treated `age = -9` as missing.
- Did not standardize the binary sensory columns because they are already on the same 0/1 scale.
- Kept `score` and `pct` out of the clustering model so they could be used later to interpret cluster quality and price-positioning potential.

## Cluster Interpretation Framework
- `Classic Smoky-Sweet`: familiar medium-bodied malts with both smoke and sweetness.
- `Sherried & Sweet`: richer, fruit-forward, smoother malts.
- `Light & Approachable`: lighter, smoother entry-point whiskies.
- `Herbal & Crisp`: leaner grassy and fruit-led styles.
- `Maritime Peat`: coastal, salty, peated whiskies that create a strong store identity.

## Managerial Use
The clusters are not just academic groups. They support three retail decisions:
1. How broad the opening shelf should be.
2. How many bottles Kathryn should carry in each style family.
3. How to build Wednesday tasting flights that show meaningful contrast.
