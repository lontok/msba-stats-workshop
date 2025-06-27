# MSBA Statistics & Excel Workshop

## Day 1: Foundations of Descriptive Statistics & Excel Essentials

### Session 1 (2 hours)

#### Introduction & Workshop Overview (15 minutes)
**Learning Objectives:**
- Master fundamental descriptive statistics concepts
- Develop proficiency in Excel for business analytics
- Apply statistical methods to real business problems
- Build confidence in data analysis techniques

**Workshop Expectations:**
- Active participation in exercises
- Questions are encouraged
- Focus on practical applications
- All materials will be provided

**Technical Requirements:**
- Excel 2016 or later (Office 365 preferred)
- Windows or Mac OS
- Basic mouse/trackpad skills

**Quick Excel Skills Assessment:**
- Basic formula knowledge check
- Navigation familiarity
- Current comfort level with data manipulation

---

#### Excel Fundamentals for Analytics (45 minutes)

**Platform Note:** Mac users should use Cmd (⌘) instead of Ctrl, and Option instead of Alt. Function keys may require holding the Fn key.

**1. Navigating Excel Efficiently**

Essential keyboard shortcuts:

| Action | Windows | Mac |
|--------|---------|-----|
| Navigate to data boundaries | Ctrl+Arrow | Cmd+Arrow |
| Select data ranges | Ctrl+Shift+Arrow | Cmd+Shift+Arrow |
| Edit cell contents | F2 | F2 (or Fn+F2) |
| Toggle absolute references | F4 | F4 (or Fn+F4) |
| New line within cell | Alt+Enter | Option+Enter |
| Copy | Ctrl+C | Cmd+C |
| Paste | Ctrl+V | Cmd+V |
| Undo | Ctrl+Z | Cmd+Z |
| Find | Ctrl+F | Cmd+F |
| Save | Ctrl+S | Cmd+S |

**2. Cell References**
- Relative references (A1): Change when copied
- Absolute references ($A$1): Stay fixed when copied
- Mixed references ($A1 or A$1): Partially fixed

**3. Essential Formulas**
```excel
=SUM(A1:A10)          # Add values in range
=AVERAGE(B1:B20)      # Calculate mean
=COUNT(C1:C30)        # Count numeric values
=COUNTA(D1:D40)       # Count non-empty cells
=IF(E1>100,"High","Low")  # Conditional logic
```

**Excel Exercise: Creating a Sales Summary Template**
- Task: Build a dynamic sales report template
- Data: Monthly sales figures for 5 products
- Requirements:
  - Calculate total sales per product
  - Find average monthly sales
  - Identify best/worst performing months
  - Use absolute references for tax calculations

---

#### Introduction to Descriptive Statistics (45 minutes)

**1. Types of Data**
- **Nominal**: Categories without order (e.g., product types, departments)
- **Ordinal**: Ordered categories (e.g., satisfaction ratings, education levels)
- **Interval**: Numeric with equal intervals, no true zero (e.g., temperature)
- **Ratio**: Numeric with true zero (e.g., sales revenue, customer count)

**2. Measures of Central Tendency**
- **Mean**: Average value
  - When to use: Symmetric data without outliers
  - Excel: `=AVERAGE(range)`
  
- **Median**: Middle value
  - When to use: Skewed data or outliers present
  - Excel: `=MEDIAN(range)`
  
- **Mode**: Most frequent value
  - When to use: Categorical data or finding peaks
  - Excel: `=MODE.SNGL(range)`

**Excel Exercise: Calculate Measures of Central Tendency**
- Dataset: Employee salaries by department
- Tasks:
  1. Calculate mean salary per department
  2. Find median salary to handle outliers
  3. Identify modal salary range
  4. Compare results and interpret differences

---

#### Break & Q&A (15 minutes)

---

### Session 2 (2 hours)

#### Measures of Variability (45 minutes)

**1. Range**
- Definition: Maximum - Minimum
- Excel: `=MAX(range)-MIN(range)`
- Limitation: Sensitive to outliers

**2. Variance**
- Population variance: σ²
- Sample variance: s²
- Excel: `=VAR.P(range)` or `=VAR.S(range)`

**3. Standard Deviation**
- Square root of variance
- Same units as original data
- Excel: `=STDEV.P(range)` or `=STDEV.S(range)`

**4. Coefficient of Variation**
- CV = (Standard Deviation / Mean) × 100%
- Compares variability across different scales
- Excel: `=STDEV.S(range)/AVERAGE(range)*100`

**Excel Exercise: Analyzing Product Sales Variability**
- Dataset: Daily sales for 3 products over 30 days
- Tasks:
  1. Calculate range for each product
  2. Compute standard deviation
  3. Find coefficient of variation
  4. Determine which product has most consistent sales
  5. Create summary table with conditional formatting

---

#### Data Visualization Basics (45 minutes)

**1. Creating Effective Histograms**
- Understanding frequency distributions
- Choosing appropriate bin sizes
- Excel steps:
  1. Select data range
  2. Insert → Charts → Histogram (Insert tab → Charts group)
  3. Adjust bin width and boundaries
  4. Format for clarity (right-click chart for options)

**2. Building Box Plots**
- Five-number summary visualization
- Identifying outliers
- Excel steps:
  1. Calculate quartiles: `=QUARTILE.EXC(range, n)`
  2. Insert → Charts → Box and Whisker (Statistical Charts)
  3. Interpret results

**3. Distribution Patterns**
- Normal (bell-shaped)
- Skewed (left or right)
- Bimodal
- Uniform

**Excel Exercise: Building Charts with Business Data**
- Dataset: Customer purchase amounts
- Tasks:
  1. Create histogram of purchase amounts
  2. Build box plot by customer segment
  3. Identify distribution shape
  4. Find and highlight outliers
  5. Add trend lines and data labels

---

#### Data Cleaning & Preparation (20 minutes)

**1. Text Functions**
```excel
=TRIM(A1)              # Remove extra spaces
=CLEAN(B1)             # Remove non-printable characters
=SUBSTITUTE(C1,"old","new")  # Replace text
=UPPER(D1)             # Convert to uppercase
=PROPER(E1)            # Capitalize first letters
```

**2. Handling Missing Data**
- Identify blanks: `=ISBLANK(cell)`
- Count blanks: `=COUNTBLANK(range)`
- Fill blanks: `=IF(ISBLANK(A1),AVERAGE(range),A1)`
- Quick select blanks: Ctrl/Cmd+G → Special → Blanks

**Excel Exercise: Cleaning a Messy Customer Dataset**
- Dataset: Customer information with errors
- Tasks:
  1. Remove extra spaces from names
  2. Standardize phone number format
  3. Fill missing values appropriately
  4. Create data quality report

---

#### Wrap-up & Practice Assignment (10 minutes)

**Key Takeaways:**
- Excel is powerful for statistical analysis
- Descriptive statistics summarize data effectively
- Clean data is essential for accurate analysis
- Visualization reveals patterns and outliers

**Practice Assignment:**
- Analyze a retail dataset:
  - Calculate all descriptive statistics
  - Create appropriate visualizations
  - Clean and prepare data
  - Write brief interpretation of findings

**Resources:**
- Excel function reference guide
- Sample datasets for practice
- Video tutorials for review

**Next Session Preview:**
- Probability concepts
- Normal distribution
- Advanced Excel functions