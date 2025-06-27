# Session 1: Step-by-Step Tutorial

## Pre-Workshop Setup

1. **Open Excel** (2016 or later)
2. **Create a new folder** called "MSBA_Stats_Workshop" on your desktop
3. **Download the CSV files** from the session1-workbooks folder
4. **Enable Data Analysis Toolpak** (if not already enabled):
   - Windows: File → Options → Add-ins → Manage: Excel Add-ins → Go → Check "Analysis ToolPak"
   - Mac: Tools → Excel Add-ins → Check "Analysis ToolPak"

---

## Exercise 1: Sales Summary Template (45 minutes)

### Part A: Import and Setup (5 minutes)

1. **Import the data**:
   - File → Open → Navigate to `workbook1_sales_data.csv`
   - Or: Data → From Text/CSV → Select file → Load

2. **Save as Excel workbook**:
   - File → Save As → "Sales_Summary_Template.xlsx"

3. **Format the data**:
   - Select all data (Ctrl/Cmd+A)
   - Home → Format as Table → Choose a style
   - Name your table: "SalesData" (Table Design tab)

### Part B: Calculate Total Sales (10 minutes)

1. **Add Total Sales column**:
   - In cell G1, type "Total Sales"
   - In cell G2, enter: `=SUM(B2:F2)`
   - Copy formula down to G13

2. **Calculate Product Totals**:
   - In cell A14, type "Total"
   - In cell B14, enter: `=SUM(B2:B13)`
   - Copy formula across to F14

3. **Grand Total**:
   - In cell G14, enter: `=SUM(G2:G13)`
   - Verify it matches: `=SUM(B14:F14)`

### Part C: Average Monthly Sales (10 minutes)

1. **Create summary section**:
   - In cell A16, type "Summary Statistics"
   - In cells A17:A19, type: "Average", "Maximum", "Minimum"

2. **Calculate for each product**:
   - In B17: `=AVERAGE(B2:B13)`
   - In B18: `=MAX(B2:B13)`
   - In B19: `=MIN(B2:B13)`
   - Copy formulas across to columns C-F

### Part D: Tax Calculations with Absolute References (10 minutes)

1. **Set up tax rate**:
   - In cell I1, type "Tax Rate:"
   - In cell J1, enter: 0.08 (8%)
   - Name this cell: "TaxRate" (Name Box)

2. **Calculate tax for each month**:
   - In cell H1, type "Tax Amount"
   - In cell H2, enter: `=G2*$J$1`
   - Copy formula down to H13

3. **Test absolute reference**:
   - Change tax rate in J1 to 0.10
   - Verify all tax calculations update automatically

### Part E: Conditional Formatting (10 minutes)

1. **Highlight top performers**:
   - Select range B2:F13
   - Home → Conditional Formatting → Top/Bottom Rules → Top 10%
   - Choose green fill

2. **Create data bars**:
   - Select range G2:G13
   - Home → Conditional Formatting → Data Bars → Choose gradient fill

3. **Color scale for variance**:
   - Add column for "% of Average"
   - Formula: `=G2/AVERAGE($G$2:$G$13)`
   - Apply 3-color scale (red-yellow-green)

---

## Exercise 2: Employee Salaries Analysis (45 minutes)

### Part A: Import and Organize (5 minutes)

1. **Import data**: Open `workbook2_employee_salaries.csv`
2. **Save as**: "Employee_Salaries_Analysis.xlsx"
3. **Create department summary sheet**:
   - Right-click sheet tab → Insert → Worksheet
   - Name it "Department_Analysis"

### Part B: Calculate Central Tendency by Department (15 minutes)

1. **Set up summary table** on Department_Analysis sheet:
   ```
   A1: Department
   B1: Count
   C1: Mean Salary
   D1: Median Salary
   E1: Mode Salary
   ```

2. **List unique departments** (A2:A7):
   - Sales, Marketing, IT, HR, Finance, Operations

3. **Calculate statistics**:
   - Count (B2): `=COUNTIF(Sheet1!B:B,A2)`
   - Mean (C2): `=AVERAGEIF(Sheet1!B:B,A2,Sheet1!D:D)`
   - Median (D2): Use array formula or pivot table
   - Mode (E2): `=MODE.SNGL(IF(Sheet1!B:B=A2,Sheet1!D:D))`
   - Copy formulas down

### Part C: Identify Outliers (15 minutes)

1. **Calculate quartiles for each department**:
   - Add columns: Q1, Q3, IQR, Lower Bound, Upper Bound
   - Q1: `=QUARTILE.INC(IF(Sheet1!B:B=A2,Sheet1!D:D),1)`
   - Q3: `=QUARTILE.INC(IF(Sheet1!B:B=A2,Sheet1!D:D),3)`
   - IQR: `=Q3-Q1`
   - Lower: `=Q1-1.5*IQR`
   - Upper: `=Q3+1.5*IQR`

2. **Flag outliers** in original data:
   - Add column "Outlier?" in original sheet
   - Use VLOOKUP to check bounds

### Part D: Create Comparison Dashboard (10 minutes)

1. **Insert pivot table**:
   - Select all data → Insert → PivotTable
   - Rows: Department, Position
   - Values: Average of Salary

2. **Create box plot**:
   - Select department and salary columns
   - Insert → Statistical Charts → Box and Whisker

3. **Add insights**:
   - Which department has highest average salary?
   - Which has most variation?
   - Where are the outliers?

---

## Exercise 3: Product Sales Variability (45 minutes)

### Part A: Import and Setup (5 minutes)

1. **Import**: `workbook3_product_sales_variability.csv`
2. **Save as**: "Product_Variability_Analysis.xlsx"
3. **Add analysis columns** for each product

### Part B: Calculate Variability Measures (20 minutes)

1. **Create summary table** (starting in E1):
   ```
   Measure         | Product A | Product B | Product C
   Mean            |           |           |
   Median          |           |           |
   Range           |           |           |
   Variance        |           |           |
   Std Deviation   |           |           |
   Coeff of Var %  |           |           |
   ```

2. **Enter formulas** for Product A (column F):
   - Mean: `=AVERAGE(B2:B31)`
   - Median: `=MEDIAN(B2:B31)`
   - Range: `=MAX(B2:B31)-MIN(B2:B31)`
   - Variance: `=VAR.S(B2:B31)`
   - Std Dev: `=STDEV.S(B2:B31)`
   - CV%: `=F6/F2*100`

3. **Copy formulas** to columns G and H

### Part C: Identify Most Consistent Product (10 minutes)

1. **Compare CV%** - lowest CV% = most consistent
2. **Create ranking**:
   - Add "Consistency Rank" row
   - Use RANK function: `=RANK(F7,$F$7:$H$7,1)`

3. **Visualize with conditional formatting**:
   - Apply color scale to CV% row
   - Green = most consistent, Red = least consistent

### Part D: Create Control Chart (10 minutes)

1. **Calculate control limits** for each product:
   - Upper Control Limit: `=Mean + 3*StdDev`
   - Lower Control Limit: `=Mean - 3*StdDev`

2. **Create line chart**:
   - Select days and all three products
   - Insert → Line Chart
   - Add horizontal lines for control limits
   - Format for clarity

---

## Exercise 4: Customer Purchase Analysis (45 minutes)

### Part A: Import and Explore (5 minutes)

1. **Import**: `workbook4_customer_purchases.csv`
2. **Save as**: "Customer_Purchase_Analysis.xlsx"
3. **Quick statistics**:
   - Total customers by segment
   - Average purchase by segment

### Part B: Create Histogram (15 minutes)

1. **Prepare bins**:
   - Create bin ranges: 0-100, 100-200, 200-300, etc.
   - Or let Excel auto-calculate

2. **Insert histogram**:
   - Select Purchase_Amount column
   - Insert → Statistical Charts → Histogram
   - Format axes and labels

3. **Analyze distribution**:
   - Is it normal? Skewed?
   - Where is the peak?
   - Any unusual patterns?

### Part C: Build Box Plots by Segment (15 minutes)

1. **Prepare data**:
   - Sort by Segment
   - Or use pivot table approach

2. **Create box plot**:
   - Select Segment and Purchase_Amount
   - Insert → Statistical Charts → Box and Whisker
   - Format for clarity

3. **Identify insights**:
   - Which segment has highest median purchase?
   - Which has most variation?
   - Any outliers? (mark with special formatting)

### Part D: Statistical Summary Dashboard (10 minutes)

1. **Create segment summary**:
   ```
   Segment  | Count | Mean | Median | StdDev | Min | Max | Q1 | Q3
   Budget   |       |      |        |        |     |     |    |
   Standard |       |      |        |        |     |     |    |
   Premium  |       |      |        |        |     |     |    |
   ```

2. **Add visual indicators**:
   - Conditional formatting for high/low values
   - Mini charts (sparklines) for distribution

---

## Exercise 5: Data Cleaning (20 minutes)

### Part A: Import Messy Data (2 minutes)

1. **Import**: `workbook5_messy_customer_data.csv`
2. **Save as**: "Customer_Data_Cleaning.xlsx"
3. **Create backup**: Copy to Sheet2

### Part B: Clean Names (5 minutes)

1. **Standardize format**:
   - Add helper column: `=PROPER(TRIM(A2))`
   - Copy and paste values back

2. **Check results**:
   - Filter to see changes
   - Verify no data lost

### Part C: Fix Phone Numbers (5 minutes)

1. **Standardize format**:
   - Remove all non-numeric: `=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B2,"-",""),"(",""),")","")` 
   - Format as: `=TEXT(cleaned_number,"(000) 000-0000")`

2. **Handle missing area codes**:
   - Flag incomplete numbers
   - Add default area code if needed

### Part D: Handle Missing Values (5 minutes)

1. **Identify missing data**:
   - Use Conditional Formatting → Highlight Cells → Blanks
   - Count blanks: `=COUNTBLANK(range)`

2. **Fill missing amounts**:
   - Calculate average by similar customers
   - Or use median of segment
   - Formula: `=IF(ISBLANK(F2),AVERAGE(similar_range),F2)`

### Part E: Create Data Quality Report (3 minutes)

1. **Summary statistics**:
   ```
   Field          | Total Records | Complete | Missing | % Complete
   Customer_Name  |              |          |         |
   Phone_Number   |              |          |         |
   Email          |              |          |         |
   City           |              |          |         |
   Purchase_Date  |              |          |         |
   Amount         |              |          |         |
   ```

2. **Add visual summary**:
   - Data bars for % Complete
   - Color coding for quality levels

---

## Common Troubleshooting Tips

### Formula Errors
- `#DIV/0!`: Check for division by zero
- `#VALUE!`: Check data types match formula requirements
- `#REF!`: Check cell references haven't been deleted
- `#N/A`: Check VLOOKUP ranges and exact match settings

### Performance Issues
- Turn off automatic calculation for large datasets
- Use tables for better performance
- Avoid volatile functions in large datasets

### Platform Differences (Windows vs Mac)
- Mac: Use Cmd instead of Ctrl
- Mac: Use Option instead of Alt
- Mac: Some shortcuts require Fn key
- Mac: Data Analysis Toolpak location may differ

---

## Answer Keys

### Exercise 1 Key Insights
- Total annual sales: ~$2,074,600
- Best performing product: Product C
- Most growth: Product E (37% increase)
- Seasonal pattern: Q4 strongest

### Exercise 2 Key Insights
- Highest avg salary: Finance ($75,857)
- Most variation: IT (due to director outlier)
- Most employees: Operations (8)
- Salary gaps: Significant between roles

### Exercise 3 Key Insights
- Most consistent: Product C (CV = 1.88%)
- Most variable: Product B (CV = 2.59%)
- All products within control limits
- No concerning trends identified

### Exercise 4 Key Insights
- Premium segment: Highest average ($425)
- Notable outliers: 2 in Standard segment
- Distribution: Right-skewed for all segments
- Clear segmentation in purchasing behavior

### Exercise 5 Key Insights
- 30 total records
- 13% missing email addresses
- 20% missing purchase amounts
- All names had formatting issues
- Phone formats were inconsistent