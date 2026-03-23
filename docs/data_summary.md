# Data Summary

## Files Used
- `whiskey_case (1).pdf`: assignment instructions.
- `ScotchData 4 (1).zip`: source archive.
- `ScotchData_unzipped/ScotchData 4/README.TXT`: package-level readme.
- `ScotchData_unzipped/ScotchData 4/Scotch data (Windows)/readme.txt`: detailed data readme.
- `ScotchData_unzipped/ScotchData 4/Scotch data (Windows)/scotch.xlsx`: primary analysis dataset.

## Workbook Structure
- Workbook sheet count: 1
- Sheet name: `scotch.xls`
- Header format: 2 header rows
- Data records used: 109 whiskies
- Blank rows removed: 1

## Variables
- Identifier fields: 2
  - `name`
  - `dist_short`
- Sensory binary descriptors: 68
  - Color: 14
  - Nose: 12
  - Body: 8
  - Palate: 15
  - Finish: 19
- Numeric fields: 4
  - `age`
  - `dist`
  - `score`
  - `pct`
- Text categorical fields: 2
  - `region`
  - `district`
- Region one-hot fields: 9

## Data Types
- Binary fields are coded as `0`/`1`.
- `score` and `pct` are numeric and fully populated.
- `age` includes a sentinel value of `-9` for unknown age. Those values were treated as missing for summary purposes.
- `region` has three populated values in this dataset: `HIGH`, `LOW`, and `ISLAY`.

## Missingness and Quality Notes
- Blank record removed: 1
- Missing `age` after cleaning sentinel values: 37 of 109 whiskies
- Missing `score`: 0
- Missing `pct`: 0
- Missing `region`: 0
- Missing `district`: 0
- `dist` exists in the workbook but is not clearly defined in the bundled readmes, so it was excluded from the clustering model.

## High-Level Descriptives
- Average score: 75.59
- Median score: 76
- Average ABV (`pct`): 41.14
- ABV distribution:
  - 40%: 83 whiskies
  - 43%: 17 whiskies
  - 46%: 6 whiskies
  - Higher special values: 45.8%, 54.8%, 57.1%
- Region distribution:
  - Highland (`HIGH`): 89
  - Lowland (`LOW`): 12
  - Islay (`ISLAY`): 8

## Fields Used for Clustering
- Included: the 68 sensory descriptors
- Excluded:
  - `name`, `dist_short`: identifiers
  - `age`: too much missingness
  - `dist`: unclear meaning in the source documentation
  - `score`, `pct`: used for interpretation, not for clustering
  - `region`, `district`, region flags: used for interpretation, not for clustering
