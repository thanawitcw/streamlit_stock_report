This Streamlit application generates a **Daily Stock Report** by combining stock data, sell-out data, purchase order information, and product master data.  
The tool processes multiple Excel files, calculates stock coverage metrics such as **DOH (Days on Hand)**, integrates pending purchase orders, and produces a final multi-sheet Excel report for daily inventory monitoring.

The application is built with **Streamlit**, allowing users to upload required files and generate the report through a simple web interface.

---

# Daily Stock Report Generator

## Data Sources

### 1) Weekly CJ Stock Report
- File: `Sahamit Report ..-..-2025.xlsx`
- Sheet: `Sahamit Report`
- Used fields:
  - CJ_Item
  - Division
  - Product Name
  - Category / Subcategory
  - Brand
  - Store Stock
  - Assortment metrics

### 2) Daily DC Stock Report
- File: `DC_End..-..-2025.xlsx`
- Sheet: `Sheet1`
- Used fields:
  - Material → mapped to `CJ_Item`
  - Plant (DC1, DC2, DC4)
  - Stock Qty
  - Stock Value

### 3) Sell-out Past 30 Days
- File: `sellout_past30D.xlsx`
- Content: daily sales transactions for the last 30 days
- Used fields:
  - Calendar Day
  - CJ_Item
  - Total_SellOut_Qty
  - Total_Sellout_Amt

### 4) Access Database Extract
- File: `data_from_access.xlsx`
- Sheets:
  - `Master_Product`
  - `Pending_All_Div`
- Content:
  - product master data
  - pending purchase orders across divisions
- Used fields:
  - CJ_Item
  - SHM_Item
  - Supplier Name
  - Order Qty
  - Unit
  - Delivery Date
  - DC_Name

### 5) Master Lead Time
- File: `Master_LeadTime.xlsx`
- Sheet: `All_Product`
- Used fields:
  - CJ_Item
  - SHM_Item
  - OwnerSCM
  - Base Lead Time (Days)

---

## What the script does

1. Loads user-uploaded Excel files through a **Streamlit interface**.
2. Processes the **CJ stock report**, renaming columns and removing excluded divisions.
3. Processes the **DC stock file**:
   - Maps plant codes to DC names
   - Replaces negative stock values with zero
   - Aggregates stock by `CJ_Item` and DC.
4. Processes **sell-out data**:
   - Cleans date fields
   - Filters unwanted product groups
   - Aggregates total sales for:
     - last 30 days
     - last 7 days.
5. Loads **Access-extracted datasets** containing:
   - product master
   - pending purchase orders.
6. Processes **pending purchase orders**:
   - Standardizes DC names
   - Converts order quantities to pieces
   - Aggregates PO quantities and minimum delivery dates per DC.
7. Loads **lead time master data** and attaches SCM ownership.
8. Standardizes SKU identifiers (`CJ_Item`, `SHM_Item`) across all datasets.
9. Merges all datasets into a unified table:
   - CJ stock
   - DC stock
   - sales data
   - purchase orders
   - product master
   - lead time information.
10. Calculates key supply chain metrics:
    - total remaining stock across DCs
    - average sales for 90 / 30 / 7 days
    - DC-level sales allocation ratios.
11. Computes **DOH (Days on Hand)** metrics:
    - current DOH by DC
    - DOH across all DCs
    - DOH including pending purchase orders.
12. Calculates **stock cover dates** using:
    - current stock levels
    - incoming PO delivery dates.
    - average sales for 90 days
13. Handles duplicate `CJ_Item` records using custom rules for special products.
14. Executes an SQL query (via `pandasql`) to produce the final report structure.
15. Cleans and formats the dataset:
    - removes rows with zero metrics
    - converts date columns
    - renames fields for reporting.
16. Generates two reporting formats:
    - **Data by Qty** (unit level)
    - **Data by CTN** (carton level).
17. Creates additional PO analysis tables:
    - all pending purchase orders
    - minimum expected delivery date per DC.
18. Exports all results into a **multi-sheet Excel file**.

---

## Output

The script generates a downloadable Excel file:

`Sahamit_Daily_Stock_Report_<date>.xlsx`
