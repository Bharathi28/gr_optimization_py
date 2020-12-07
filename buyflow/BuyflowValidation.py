from pathlib import Path
from testbase.CommonUtilities import getExcelData

project_path = str(Path(__file__).parents[1])
print(project_path)
print (project_path + "\\Input_Output\\BuyflowValidation\\new_run_input.xlsx")


data = getExcelData(input_path = project_path + "\\Input_Output\\BuyflowValidation\\new_run_input.xlsx", sheet_name = "rundata", start_row=1)
