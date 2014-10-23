#Libraries to parse xls docs
from xlrd import open_workbook,cellname

#Libraries to create xlx files for further use and data preparation
import xlsxwriter


#Open data sheet
book = open_workbook('bat/complete_batsmen_stats.xls')


#Index data sheet
sheet = book.sheet_by_index(0)


# Create Data Sheets
def create_new_sheet(sheet_name):
	workbook = xlsxwriter.Workbook(sheet_name)
	new_sheet = workbook.add_worksheet()
	return (workbook,new_sheet)


# Get the Min and MAx for GRA Calculation
def min_max(a):
	min_a = 10000000000
	max_a = -1
	for x in a:
		if max_a < float(a[x]):
			max_a = float(a[x])
		if min_a > float(a[x]):
			min_a = float(a[x])	
	return (max_a,min_a)		


#Grey Relation Analysis
def normalize_complete_batsmen_stats():
	runs = {}
	innings = {}
	avg = {}
	sr = {}
	century = {}
	hcentury = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):			
			if cellname(row_index,col_index)[0] == 'A':
				curr_batsmen = sheet.cell(row_index,col_index).value
				runs[curr_batsmen] = sheet.cell(row_index,col_index+3).value
				innings[curr_batsmen] = sheet.cell(row_index,col_index+1).value
				avg[curr_batsmen] = sheet.cell(row_index,col_index+6).value
				sr[curr_batsmen] = sheet.cell(row_index,col_index+7).value
				hcentury[curr_batsmen] = sheet.cell(row_index,col_index+8).value
				century[curr_batsmen] = sheet.cell(row_index,col_index+9).value
	rpm = {}
	for x in runs:
		rpm[x] = runs[x]/innings[x]
	max_rpm , min_rpm = min_max(rpm)
	max_avg , min_avg = min_max(avg)
	max_sr , min_sr = min_max(sr)
	max_c , min_c = min_max(century)
	max_hc , min_hc = min_max(hcentury)

	#Create new sheet
	workbook,normalized_complete_batsmen_stats_sheet = create_new_sheet(
														"normalized_complete_batsmen_stats.xls"
														)

	#Initialize rows,columns
	row_count = 0
	normalized_complete_batsmen_stats_sheet.write(0,0,"Player Name")
	normalized_complete_batsmen_stats_sheet.write(0,1,"Runs per innings")
	normalized_complete_batsmen_stats_sheet.write(0,2,"AVG")
	normalized_complete_batsmen_stats_sheet.write(0,3,"SR")
	normalized_complete_batsmen_stats_sheet.write(0,4,"50's")
	normalized_complete_batsmen_stats_sheet.write(0,5,"100's")
	row_count+=1

	for x in runs:
		normalized_complete_batsmen_stats_sheet.write(row_count,0,x)
		normalized_complete_batsmen_stats_sheet.write(row_count,1,
																"%.2f"%((rpm[x] - min_rpm) / (max_rpm - min_rpm))
																)
		normalized_complete_batsmen_stats_sheet.write(row_count,2,
																"%.2f"%((float(avg[x]) - min_avg) / (max_avg - min_avg))
																)
		normalized_complete_batsmen_stats_sheet.write(row_count,3,
																"%.2f"%((float(sr[x]) - min_sr) / (max_sr - min_sr))
																)
		normalized_complete_batsmen_stats_sheet.write(row_count,4,
																"%.2f"%((hcentury[x] - min_hc) / (max_hc - min_hc))
																)
		normalized_complete_batsmen_stats_sheet.write(row_count,5,
																"%.2f"%((century[x] - min_c) / (max_c - min_c))
																)
		row_count+=1
	workbook.close()	





				
				


normalize_complete_batsmen_stats()
