import pandas as pd
from typing import List


AION_COLUMNS: List[str] = ["Part", "Value", "Description", "Notes"]


def read_aion_fx_xlsx_file(file_path: str) -> pd.DataFrame:
    """
    Read an Aion FX Excel file and extract all relevant sheets into a combined DataFrame.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Combined DataFrame with all relevant sheets
    """
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names

    # Process all sheets except the first one (Instructions) and any combined sheets
    # Look for sheets that contain parts data
    relevant_sheets = []
    for sheet in sheet_names:
        if "instruction" not in sheet.lower() and "combined" not in sheet.lower():
            relevant_sheets.append(sheet)
    
    combined_df = pd.DataFrame()

    for sheet in relevant_sheets:
        df = xl.parse(sheet)
        df = df[AION_COLUMNS]
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    return combined_df


# def read_aion_fx_url(url: str) -> pd.DataFrame: