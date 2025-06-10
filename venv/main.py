from fastapi import FastAPI, Query, HTTPException
import pandas as pd
import re
from typing import List, Dict, Any
import uvicorn
from contextlib import asynccontextmanager
import os

# Global variable to store the parsed data
excel_data = {}

def parse_excel_data():
    """Parse the Excel/CSV file and extract table structures"""
    try:
        # Try multiple file paths and formats
        possible_files = [
            'venv/Data/capbudg.xls',
            'Data/capbudg.xls',
            'capbudg.xls',
            'venv/Data/capbudg.xlsx',
            'Data/capbudg.xlsx',
            'capbudg.xlsx'
        ]
        
        df = None
        used_file = None
        
        print(f"Trying to load from {len(possible_files)} possible file locations...")
        
        for file_path in possible_files:
            try:
                print(f"Trying: {file_path}")
                if os.path.exists(file_path):
                    if file_path.endswith('.xls'):
                        df = pd.read_excel(file_path, header=None, engine='xlrd')
                    else:
                        df = pd.read_excel(file_path, header=None)
                    used_file = file_path
                    print(f"Successfully loaded: {file_path}")
                    break
                else:
                    print(f"File not found: {file_path}")
            except Exception as e:
                print(f"Failed to load {file_path}: {e}")
                continue
        
        if df is None:
            current_dir = os.getcwd()
            print(f"Current working directory: {current_dir}")
            if os.path.exists('.'):
                print(f"Files in current directory: {os.listdir('.')}")
            raise FileNotFoundError(f"Could not find any of these files: {possible_files}")
        
        print(f"File loaded successfully from: {used_file}")
        print(f"Data shape: {df.shape}")
        
        # Initialize data structure
        tables = {}
        table_headers = {}
        current_table = None
        
        # Print first few rows for debugging
        print("\nFirst 10 rows of the Excel file:")
        for idx in range(min(10, len(df))):
            row_values = [str(val).strip() if not pd.isna(val) else "" for val in df.iloc[idx]]
            print(f"Row {idx}: {row_values}")
        
        for idx, row in df.iterrows():
            row_values = [str(val).strip() if not pd.isna(val) else "" for val in row]
            first_col = row_values[0] if row_values else ""
            
            # More flexible table detection
            first_col_upper = first_col.upper()
            
            if "INITIAL INVESTMENT" in first_col_upper:
                current_table = "Initial Investment"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "CASHFLOW" in first_col_upper and "DETAIL" in first_col_upper:
                current_table = "Cashflow Details"
                table_headers[current_table] = {"start": idx, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "REVENUE" in first_col_upper and "PROJECTION" in first_col_upper:
                current_table = "Revenue Projections"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "OPERATING" in first_col_upper and "EXPENSE" in first_col_upper:
                current_table = "Operating Expenses"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "DISCOUNT RATE" in first_col_upper:
                current_table = "Discount Rate"
                table_headers[current_table] = {"start": idx, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "WORKING CAPITAL" in first_col_upper:
                current_table = "Working Capital"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "GROWTH RATE" in first_col_upper:
                current_table = "Growth Rates"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "OPERATING CASHFLOW" in first_col_upper:
                current_table = "Operating Cashflows"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "SALVAGE VALUE" in first_col_upper:
                current_table = "Salvage Value"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "INVESTMENT MEASURE" in first_col_upper or "NPV" in first_col_upper:
                current_table = "Investment Measures"
                table_headers[current_table] = {"start": idx, "rows": []}
                print(f"Found table: {current_table} at row {idx}")
            elif "BOOK VALUE" in first_col_upper and "DEPRECIATION" in first_col_upper:
                current_table = "Book Value & Depreciation"
                table_headers[current_table] = {"start": idx + 1, "rows": []}
                print(f"Found table: {current_table} at row {idx}")

            # Collect rows for current table
            if current_table and current_table in table_headers:
                # Skip header rows and empty rows
                if (first_col and 
                    first_col not in ["", "INITIAL INVESTMENT", "CASHFLOW DETAILS", "REVENUE PROJECTIONS", 
                                    "OPERATING EXPENSES", "DISCOUNT RATE", "WORKING CAPITAL", "GROWTH RATES", 
                                    "OPERATING CASHFLOWS", "SALVAGE VALUE", "BOOK VALUE & DEPRECIATION", "INVESTMENT MEASURES"] and
                    ("=" in first_col or any(char.isdigit() for char in first_col) or 
                     any(keyword in first_col.lower() for keyword in ["cost", "rate", "value", "investment", "revenue", "expense"]))):
                    table_headers[current_table]["rows"].append(idx)
                    print(f"Added row {idx} to {current_table}: {first_col}")

        # Build final table structure
        for table_name, config in table_headers.items():
            if not config["rows"]:
                print(f"No rows found for table: {table_name}")
                continue

            table_data = {}
            print(f"\nProcessing table: {table_name}")

            for row_idx in config["rows"]:
                if row_idx < len(df):
                    row_label = df.iloc[row_idx, 0]
                    if pd.isna(row_label) or str(row_label).strip() == "":
                        continue

                    row_values = []
                    for col_idx in range(1, len(df.columns)):
                        cell_value = df.iloc[row_idx, col_idx]
                        if not pd.isna(cell_value) and str(cell_value).strip() != "":
                            row_values.append(cell_value)

                    if row_values:  # Only add if there are values
                        table_data[str(row_label)] = row_values
                        print(f"  Row: {row_label} -> {row_values}")

            if table_data:
                tables[table_name] = table_data
                print(f"Added table {table_name} with {len(table_data)} rows")

        print(f"\nFinal result: {len(tables)} tables loaded")
        for table_name, table_data in tables.items():
            print(f"  {table_name}: {len(table_data)} rows")
        
        return tables

    except Exception as e:
        print(f"Error parsing Excel data: {e}")
        import traceback
        traceback.print_exc()
        return {}

def extract_numeric_value(value_str):
    """Extract numeric value from string, handling currency, percentages, etc."""
    if pd.isna(value_str):
        return 0

    value_str = str(value_str).strip()
    
    # Handle negative values in parentheses
    if value_str.startswith('(') and value_str.endswith(')'):
        value_str = '-' + value_str[1:-1]
    
    # Remove currency symbols, commas, and other formatting
    value_str = re.sub(r'[\$,£€¥]', '', value_str)
    
    # Check if it's a percentage
    is_percentage = '%' in value_str
    value_str = value_str.replace('%', '')

    try:
        numeric_value = float(value_str)
        # Convert percentage to decimal if needed (uncomment if you want percentages as decimals)
        # if is_percentage:
        #     numeric_value = numeric_value / 100
        return numeric_value
    except ValueError:
        print(f"Could not convert to numeric: '{value_str}'")
        return 0

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    global excel_data
    try:
        print("Starting to parse Excel data...")
        excel_data = parse_excel_data()
        print(f"Successfully loaded {len(excel_data)} tables from Excel file")
        if excel_data:
            print(f"Table names: {list(excel_data.keys())}")
        else:
            print("WARNING: No tables were loaded!")
    except Exception as e:
        print(f"ERROR during startup: {e}")
        import traceback
        traceback.print_exc()
        excel_data = {}
    yield
    # Shutdown (if needed)

app = FastAPI(
    title="Excel Processing API", 
    version="1.0.0",
    lifespan=lifespan
)

@app.get("/")
async def root():
    return {
        "message": "Excel Processing API",
        "endpoints": [
            "/list_tables - Get list of all tables",
            "/get_table_details?table_name=<name> - Get row names for a table",
            "/row_sum?table_name=<name>&row_name=<name> - Get sum of values in a row"
        ],
        "available_tables": list(excel_data.keys()) if excel_data else []
    }

@app.get("/debug")
async def debug_info():
    def explore_directory(path, max_depth=2, current_depth=0):
        items = {}
        if current_depth >= max_depth:
            return items
        
        try:
            for item in os.listdir(path):
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    items[f"{item}/"] = explore_directory(item_path, max_depth, current_depth + 1)
                else:
                    items[item] = "file"
        except PermissionError:
            items["<Permission Denied>"] = "error"
        
        return items
    
    return {
        "current_directory": os.getcwd(),
        "directory_structure": explore_directory('.', max_depth=3),
        "excel_data_loaded": len(excel_data) > 0,
        "excel_data_keys": list(excel_data.keys()) if excel_data else [],
        "excel_data_details": {table: list(data.keys()) for table, data in excel_data.items()} if excel_data else {}
    }

@app.get("/list_tables")
async def list_tables():
    """List all available tables from the Excel file"""
    if not excel_data:
        raise HTTPException(status_code=500, detail="Excel data not loaded")
    return {"tables": list(excel_data.keys())}

@app.get("/reload_data")
async def reload_data():
    """Reload data from Excel file (useful for debugging)"""
    global excel_data
    excel_data = parse_excel_data()
    return {
        "message": "Data reloaded",
        "tables_loaded": len(excel_data),
        "table_names": list(excel_data.keys())
    }

@app.get("/get_table_details")
async def get_table_details(table_name: str = Query(..., description="Name of the table")):
    """Get row names for a specific table"""
    if not excel_data:
        raise HTTPException(status_code=500, detail="Excel data not loaded")
    
    if table_name not in excel_data:
        available_tables = list(excel_data.keys())
        raise HTTPException(
            status_code=404, 
            detail=f"Table '{table_name}' not found. Available tables: {available_tables}"
        )
    
    return {
        "table_name": table_name,
        "row_names": list(excel_data[table_name].keys())
    }

@app.get("/row_sum")
async def row_sum(
    table_name: str = Query(..., description="Name of the table"),
    row_name: str = Query(..., description="Name of the row")
):
    """Calculate sum of all numerical values in a specific row of a table"""
    if not excel_data:
        raise HTTPException(status_code=500, detail="Excel data not loaded")
    
    if table_name not in excel_data:
        available_tables = list(excel_data.keys())
        raise HTTPException(
            status_code=404, 
            detail=f"Table '{table_name}' not found. Available tables: {available_tables}"
        )
    
    if row_name not in excel_data[table_name]:
        available_rows = list(excel_data[table_name].keys())
        raise HTTPException(
            status_code=404, 
            detail=f"Row '{row_name}' not found in table '{table_name}'. Available rows: {available_rows}"
        )

    row_values = excel_data[table_name][row_name]
    numeric_values = [extract_numeric_value(val) for val in row_values]
    total_sum = sum(numeric_values)

    return {
        "table_name": table_name,
        "row_name": row_name,
        "sum": total_sum,
        "values_processed": len(row_values),
        "numeric_values": numeric_values  # Include for debugging, remove if not needed
    }

if __name__ == "__main__":
    print("Starting FastAPI server on http://localhost:9090")
    print("Available endpoints:")
    print("  GET http://localhost:9090/list_tables")
    print("  GET http://localhost:9090/get_table_details?table_name=<name>")
    print("  GET http://localhost:9090/row_sum?table_name=<name>&row_name=<name>")
    print("  GET http://localhost:9090/debug (for debugging)")
    print("  GET http://localhost:9090/reload_data (to reload Excel data)")
    uvicorn.run(app, host="localhost", port=9090)