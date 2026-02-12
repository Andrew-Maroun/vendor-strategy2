# Vendor Spend Strategy Assessment

## Overview

This project contains a comprehensive vendor spend analysis for a portfolio company post-acquisition. The objective is to analyze ~386 vendor relationships totaling $7.89M in annual spend, identify cost-saving opportunities, and present actionable recommendations to senior leadership.

## Project Structure

```
vendor-strategy2/
├── README.md                                              # This file
├── Instructions-VendorAssessment.txt                      # Original assessment instructions
├── A - TEMPLATE - RWA - Vendor Spend Strategy (NAME) (1).xlsx  # Original template (input)
├── Vendor_Analysis_Assessment_Completed.xlsx               # Completed analysis (output)
└── vendor_analysis.py                                      # Analysis script (Claude Code CLI)
```

## How This Was Done

### Tool Used
**Claude Code CLI** (Model: claude-opus-4-6) — all analysis was performed exclusively using the Claude Code command-line interface.

### Step-by-Step Process

1. **Data Extraction**: Used Claude Code to read the Excel template via Python's `openpyxl` library. Extracted all 386 vendor names and their 12-month spend data.

2. **Spend Distribution Analysis**: Analyzed the distribution of vendor spend to identify concentration risks. Found that Salesforce alone represents 39.5% of total spend ($3.12M of $7.89M).

3. **Vendor Classification**: For each of the 386 vendors, Claude Code was used to:
   - **Identify the vendor** based on company name, regional context (e.g., D.O.O. = Croatian LLC), and known industry databases
   - **Assign a department** from the 12 valid categories in the Config tab
   - **Write a specific description** of what the vendor provides (avoiding generic descriptions)
   - **Recommend an action**: Terminate, Consolidate, or Optimize

4. **Strategic Opportunity Identification**: Grouped vendors by function to identify the three highest-impact savings opportunities:
   - CRM & Salesforce License Optimization ($850K/year)
   - Office Space & Facilities Rationalization ($550K/year)
   - Professional Services & Accounting Consolidation ($430K/year)

5. **Quality Checks**: Ran automated validation scripts to verify:
   - All 386 vendors have department, description, and recommendation (no blanks)
   - All departments match the 12 valid categories from the Config tab
   - All recommendations are one of: Terminate, Consolidate, Optimize
   - No descriptions are generic (e.g., "business services provider")
   - Financial estimates are based on realistic industry benchmarks

6. **Output Generation**: Populated all tabs of the Excel workbook:
   - Vendor Analysis Assessment (386 rows)
   - Top 3 Opportunities (with savings estimates)
   - Methodology (approach, tools, prompts, quality checks)
   - CEO/CFO Recommendations (executive memo)

### Prompts Used in Claude Code CLI
- "Analyze vendor spend data from Excel file and categorize each vendor by department, description, and strategic recommendation"
- "Identify the top 3 highest-impact cost reduction opportunities with financial justification"
- "Generate a Python script to populate the Excel template with all analysis results"
- "Run quality checks to verify completeness and accuracy of all vendor classifications"

## Key Findings

| Metric | Value |
|--------|-------|
| Total vendors analyzed | 386 |
| Total annual spend | $7,887,360 |
| Recommended for Terminate | 74 vendors |
| Recommended for Consolidate | 102 vendors |
| Recommended for Optimize | 210 vendors |
| **Total estimated annual savings** | **$1,830,000 (23.2%)** |

## Output File

The completed analysis is in `Vendor_Analysis_Assessment_Completed.xlsx` with the following tabs:
- **Vendor Analysis Assessment**: All 386 vendors with Department, Description, and Recommendation
- **Top 3 Opportunities**: Three highest-impact savings initiatives with explanations and estimated savings
- **Methodology**: Detailed explanation of approach, tools, prompts, and quality checks
- **CEOCFO Recommendations**: Executive memo summarizing findings for CEO and CFO

## Running the Analysis Script

```bash
pip install openpyxl pandas
python3 vendor_analysis.py
```

This will read the template file and produce the completed output file.
