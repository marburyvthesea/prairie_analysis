import sys
sys.path.append('/Users/johnmarshall/Documents/Analysis/PythonAnalysisScripts/calciumanalysis/prairie_analysis/')
import read_pv_jjm_matlab
import assemble_excel_files
import pandas as pd
import xlrd
import xlsxwriter




folder_path = sys.argv[1]

read_pv_jjm_matlab.output_folder_to_excel(folder_path)

assemble_excel_files.create_linescan_sheet_with_averages(str(folder_path)+'linescan' + '.xlsx')

