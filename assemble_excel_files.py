import sys
sys.path.append('/Users/johnmarshall/Documents/Analysis/PythonAnalysisScripts/calciumanalysis/prairie_analysis/')
import read_pv_jjm_matlab
import pandas as pd
import xlrd
import xlsxwriter


def create_linescan_sheet_with_averages(excel_sheet_path):
	wbRD = xlrd.open_workbook(excel_sheet_path)
	wbRD_pd = pd.ExcelFile(excel_sheet_path)
	wbWR = xlsxwriter.Workbook(excel_sheet_path.rstrip('.xlsx')+str('sweeps_averaged.xlsx')) 

	##this section just writes old data

	data_rows = wbRD.sheets()[0].nrows
	data_columns = wbRD.sheets()[0].ncols

	newSheet = wbWR.add_worksheet(wbRD.sheets()[0].name)    
	for row in range(data_rows):
		for col in range(data_columns):
			newSheet.write(row, col, wbRD.sheets()[0].cell(row, col).value);
                    
	##create columns with channel 1 and 2 averages across sweeps
	#add first average 3 spaces out from final existing data column

	channel_1_mean_column = wbRD.sheets()[0].ncols+3
	channel_2_mean_column = wbRD.sheets()[0].ncols+5
	headers = {channel_1_mean_column: 'Channel 1 mean',
				channel_2_mean_column: 'Channel 2 mean'
          		}

	#columns with channel 1 and 2 values from each sweep
	data_columns_to_average = {channel_1_mean_column:range(2, data_columns, 4),
                            	channel_2_mean_column:range(4, data_columns, 4)}

	for column_to_write in [channel_1_mean_column, channel_2_mean_column]:
		for row_to_write in range(3, data_rows):
			fields_to_average = ','.join([str(chr(65+col_int)+str(row_to_write+1)) for col_int in data_columns_to_average[column_to_write]])
			newSheet.write_formula(row_to_write, column_to_write, "=AVERAGE("+fields_to_average+")")
			newSheet.write_formula(row_to_write, column_to_write+1, "=STDEV("+fields_to_average+")")
        
	#column headers
	for column in headers.keys():
		newSheet.write(1, column, headers[column])

	wbWR.close()
	return()
	
def combine_linescan_and_voltage_recording(ls_path, vr_path):
	"""combine linescan file with vlotage recording file into one sheet"""
	## create wb to write files 
	wbWR = xlsxwriter.Workbook(ls_path.rstrip('.xlsx')+str('combined.xlsx')) 
	
	## load ls averages wb
	ls_wb = xlrd.open_workbook(ls_path)
	ls_wb_sheet = ls_wb.sheets()[0]
	ls_rows = ls_wb_sheet.nrows
	ls_columns = ls_wb_sheet.ncols
	
	## write ls sheet
	newSheet = wbWR.add_worksheet('line scan')  
	for row in range(ls_rows):
		for col in range(ls_columns):
			newSheet.write(row, col, ls_wb_sheet.cell(row, col).value)
	
	## get voltage recording sheet
	vrecording_wb = xlrd.open_workbook(vr_path)
	vrecording_sheet = vrecording_wb.sheets()[0]
	vr_rows = vrecording_sheet.nrows
	vr_columns = vrecording_sheet.cols
	
	## write voltage recording sheet
	newSheet_2 = wbWR.add_worksheet('voltage recording')
	for row in range(vr_rows):
		for col in range(vr_columns):
			newSheet_2.write(row, col, vrecording_sheet.cell(row, col).value)
	
	wbWr.close()
	return()
	










