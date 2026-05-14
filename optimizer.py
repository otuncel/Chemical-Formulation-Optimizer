import pandas as pd
from pulp import *
import win32com.client
import os
import sys

def run_optimization():
    # 1. Dynamic Path Handling
    # Gets the absolute path of the directory where the script is located
    base_path = os.path.dirname(os.path.abspath(__file__))
    # Joins the directory path with the Excel filename
    file_path = os.path.join(base_path, "optimization_input.xlsm")
    
    print(f"Searching for file at: {file_path}")

    # Check if the file exists before proceeding
    if not os.path.exists(file_path):
        print(f"Critical Error: The file 'optimization_input.xlsm' was not found in: {base_path}")
        return

    # 2. Data Loading
    try:
        df_materials = pd.read_excel(file_path, sheet_name='Materials')
        df_constraints = pd.read_excel(file_path, sheet_name='Constraints')
    except Exception as e:
        print(f"Error reading Excel sheets: {e}")
        return

    # 3. Problem Definition
    # Goal: Minimize total cost
    prob = LpProblem("Cost_Minimization", LpMinimize)
    
    # Decision Variables: Amount (kg) to be used from each raw material
    material_vars = [LpVariable(f"x_{i}", lowBound=0, upBound=df_materials.loc[i, 'Stock_Available_kg']) 
                     for i in range(len(df_materials))]

    # Objective Function: Total Cost = Sum of (Quantity * Unit Price)
    prob += lpSum([material_vars[i] * df_materials.loc[i, 'Unit_Price_USD_per_kg'] for i in range(len(df_materials))])

    # 4. Constraints Extraction
    total_batch = df_constraints[df_constraints['Parameter'] == 'Total_Batch_Size']['Value'].values[0]
    min_active = df_constraints[df_constraints['Parameter'] == 'Min_Total_Active_Matter']['Value'].values[0]

    # Constraint: Total weight must equal the target batch size
    prob += lpSum(material_vars) == total_batch
    
    # NEW Constraint: Individual Material Max Usage Limits (%)
    for i in range(len(df_materials)):
        max_limit_percent = df_materials.loc[i, 'Max_Usage_Limit_%']
        # Usage (kg) <= (Total Batch Size * Max Limit %) / 100
        prob += material_vars[i] <= (total_batch * max_limit_percent / 100)

    # Constraint: Final active matter percentage
    prob += lpSum([material_vars[i] * (df_materials.loc[i, 'Active_Matter_Content_%'] / 100) for i in range(len(df_materials))]) >= (total_batch * min_active / 100)

    # 5. Solver Execution
    status = prob.solve(PULP_CBC_CMD(msg=0))

# 6. Result Processing and Excel Export
    if LpStatus[status] == 'Optimal':
        try:
            # Connect to the ALREADY OPEN Excel application
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                is_already_open = True
            except:
                excel = win32com.client.Dispatch("Excel.Application")
                is_already_open = False
            
            excel.Visible = True # Ensure it stays visible
            
            # Find the workbook if it's already open, otherwise open it
            try:
                wb = excel.Workbooks(os.path.basename(file_path))
            except:
                wb = excel.Workbooks.Open(file_path)
            
            ws = wb.Sheets("Results")
            
            # Clear existing data
            ws.Cells.ClearContents()
            
            # Write Table Headers
            ws.Cells(1, 1).Value = "Material Name"
            ws.Cells(1, 2).Value = "Optimal Quantity (kg)"
            
            # Loop through results
            current_row = 2
            for i in range(len(df_materials)):
                qty = value(material_vars[i])
                if qty > 0:
                    ws.Cells(current_row, 1).Value = df_materials.loc[i, 'Material_Name']
                    ws.Cells(current_row, 2).Value = round(qty, 2)
                    current_row += 1
            
            # Write Summary
            ws.Cells(current_row + 1, 1).Value = "Total Cost:"
            ws.Cells(current_row + 1, 2).Value = f"{value(prob.objective):.2f} USD"
            
            # --- MODIFIED LOGIC: Do not use excel.Quit() if it was already open ---
            wb.Save()
            
            if not is_already_open:
                wb.Close()
                excel.Quit()
            
            # Release COM objects
            del ws
            del wb
            del excel
            
            print("Success: Results updated in the active Excel session.")
        except Exception as e:
            print(f"Error during Excel data export: {e}")
if __name__ == "__main__":
    run_optimization()
