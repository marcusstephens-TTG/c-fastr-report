# c-fastr-report
Harness to create test-environment C FASTR reports
This repository contains all the stuff necessary to create C FASTR reports using consultant input, test and prod

# c-fastr-report

This is a Streamlit app for generating culture health reports from survey data.

## Files

- `CFasterReportGen.py`: Main app script
- `CFASTR_Category_Mapping_V1.csv`: Category mapping file
- `CFastR_Survey_Data.csv`: Survey data (replace with your own)
- `category_copy_thresholds.csv`: Thresholds file
- `client_report_template.docx`: Word template

## How to Run

```bash
pip install streamlit pandas matplotlib python-docx
streamlit run CFasterReportGen.py
```

## Data

Make sure the CSV files and template are present in the project folder.
