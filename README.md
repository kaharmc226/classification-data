# House Price Data Preparation

This repository contains the original `DATA RUMAH.xlsx` spreadsheet with Indonesian house listings and a reproducible data-cleaning script that converts the workbook into a model-ready CSV file.

## Cleaned dataset

Run the cleaning script to regenerate the cleaned CSV:

```bash
python clean_data.py
```

The command reads `DATA RUMAH.xlsx` and writes `cleaned_house_data.csv` with the following columns:

| Column | Description |
| --- | --- |
| `name` | Listing name / short description. |
| `price` | Listing price in Indonesian Rupiah (integer). |
| `building_area` | Building area (m²). |
| `land_area` | Land area (m²). |
| `bedrooms` | Number of bedrooms. |
| `bathrooms` | Number of bathrooms. |
| `garage` | Number of available garage spaces (0 = none). |

Duplicate listings (same description and identical numeric attributes) and rows with invalid numeric values are removed during the cleaning step. The resulting dataset contains 1,008 clean records suitable for training a regression model to predict house prices from numeric attributes supplied by users.
