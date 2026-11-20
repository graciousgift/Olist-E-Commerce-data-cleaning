# Olist E-Commerce: Data Cleaning, List Matching & Verification

**Tools:** Excel, Python (pandas, openpyxl)  |  **Data:** [Brazilian E-Commerce Public Dataset by Olist (Kaggle)](https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce): 8 tables, about 1.6M rows

## The brief
A typical data-support job: update spreadsheets, match items between two lists, find missing product details, remove duplicates, format consistently and verify accuracy. I treated the Olist dataset as that job.

## What I did
| Task | What I did | Result |
|---|---|---|
| **Matching two lists** | Matched 73 product category names (list A) against the English translation list (list B) | 71 matched; 2 had no match (`pc_gamer`, `portateis_cozinha_e_preparadores_de_alimentos`, 13 products) and were translated manually and flagged |
| **Duplicate removal** | Checked every table on its key | Geolocation: 261,831 exact duplicate rows (26.2%) collapsed to 19,010 zip-prefix rows. Reviews: 814 repeated `review_id`s checked and kept, because they are one review covering several orders, not true duplicates. 1,025 products flagged as possible duplicate listings (not deleted, needs a human check) |
| **Formatting / standardising** | Fixed the translation sheet header, misspelled columns (`lenght`), city names | 8,011 city spellings reduced to 5,966 (accents, case, encoding); 42 coordinates outside Brazil excluded |
| **Missing details** | Flagged and ranked | 610 products have no category/name/description/photo data and 6 have missing or zero weight or dimensions. All 615 are in a Research_Queue ranked by revenue (R$182k), with blank columns for researched category, source and verifier |
| **Data verification** | Date logic, payment reconciliation, key matching | 1,228 order exceptions logged: 775 orders with no items, 249 payment totals differing from items + freight by more than R$1, 166 carrier pickups before purchase, 23 deliveries before pickup, 14 status/date conflicts, 1 order with no payment |
| **Documentation** | Issues_Log with rule, action and status for each of the 21 findings | Anyone can see exactly what changed and why |

## Files
- `Olist_Data_Cleaning_Portfolio.xlsx`: Summary (live formulas), Issues_Log, Category_Matching, Products_Clean, Research_Queue, Order_Exceptions, Duplicate_Reviews, Geo_Zip_Lookup
- `olist_cleaning.py`: reproducible script that builds the workbook from the source file

## Limits (read these)
- Product IDs in Olist are anonymised hashes with no product names, so the 610 products with missing details **cannot be web-researched from this data**. The queue shows how I would prioritise and record that research on real product lists.
- The payment mismatches may be legitimate (instalment interest, vouchers). They are flagged for review, not corrected.
- Possible duplicate listings are matched on attributes only. They need manual confirmation before removal.

## Source
Raw data: Kaggle, Olist Brazilian E-Commerce Public Dataset (CC BY-NC-SA 4.0). Original files are not re-hosted here.

## How to run
1. Download the 8 CSV files from the Kaggle link above and combine them into one Excel workbook (one sheet per CSV, in this order: orders, category translation, customers, payments, items, geolocation, products, reviews). Save it as `E_Commerce_Brazil_Data.xlsx`.
2. Open `Olist_Data_Cleaning_Notebook_.ipynb` in Google Colab, upload the workbook, and run all cells. Or run `python olist_cleaning.py E_Commerce_Brazil_Data.xlsx output.xlsx`.

**Licence:** Code is MIT. Data is Olist, CC BY-NC-SA 4.0 (see Kaggle).
