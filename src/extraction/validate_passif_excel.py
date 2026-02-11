import os
import pandas as pd
from src.extraction.validator_passifs import ValidatorPassifs
from src.extraction.excel_exporter import beautify_excel_layout
def validate_capitaux_propres_passif(excel_path: str, company_name: str):
    # Path to the Excel file (update as needed)
    EXCEL_PATH = excel_path
    
    COMPANY_NAME = company_name  # Update with the actual company name for assurance column

    # Load Excel data
    df = pd.read_excel(EXCEL_PATH)

    # Dynamically find the year columns (should match '31/12/YYYY' format)
    year_cols = [col for col in df.columns if col.startswith('31/12/')]

    # Filter out rows with no code, no designation, and no value in any year columns
    def is_row_empty(row):
        code_empty = pd.isna(row.get('Code')) or str(row.get('Code')).strip() == ''
        desc_empty = pd.isna(row.get('Sous-catégorie')) or str(row.get('Sous-catégorie')).strip() == ''
        years_empty = all(pd.isna(row.get(col)) or row.get(col) == 0 for col in year_cols)
        return code_empty and desc_empty and years_empty

    filtered_df = df[~df.apply(is_row_empty, axis=1)].copy()

    # Prepare validator (assumes ValidatorPassifs is implemented with the business rules)
    validator = ValidatorPassifs()

    def validate_row(row):
        result = validator.validate(row.to_dict())
        return 'PASS' if result else 'FAIL'

    filtered_df['ValidationResult'] = filtered_df.apply(validate_row, axis=1)
    filtered_df['Assurance'] = COMPANY_NAME

    # Save to a new Excel file
    output_path = EXCEL_PATH.replace('.xlsx', '_validated.xlsx')
    filtered_df.to_excel(output_path, index=False)

    # Beautify the output file and add company name
    beautify_excel_layout(output_path)

    print(f'Validation complete. Output saved to: {output_path}')
