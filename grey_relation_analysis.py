#Libraries to parse xls docs
from xlrd import open_workbook,cellname

#Libraries to create xlx files for further use and data preparation
import xlsxwriter


#Open data sheet
book = open_workbook('grey_relation_coefficients.xls')


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
			if cellname(row_index,col_index)[0] == 'A' and sheet.cell(row_index,col_index+1).value >= 20:
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


def grc_calculation():
	grc_rpm = {}
	delmax_rpm = -1
	delmin_rpm = 1
	grc_avg = {}
	delmax_avg = -1
	delmin_avg = 1
	grc_sr = {}
	delmax_sr = -1
	delmin_sr = 1
	grc_c= {}
	delmax_c = -1
	delmin_c = 1
	grc_hc = {}
	delmax_hc = -1
	delmin_hc = 1
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_rpm = float(sheet.cell(row_index,col_index+1).value)
				if curr_rpm != 0.00:
					if 1 - curr_rpm > delmax_rpm:
						delmax_rpm = 1.0- curr_rpm
				if curr_rpm != 1.00:		
					if 1 - curr_rpm < delmin_rpm:
						delmin_rpm 	= 1.0 - curr_rpm
				curr_avg = float(sheet.cell(row_index,col_index+2).value)
				if curr_avg != 0.0:
					if 1 - curr_avg > delmax_avg:
						delmax_avg = 1.0 - curr_avg
				if curr_avg != 1.00:		
					if 1 - curr_avg < delmin_avg:
						delmin_avg 	= 1.0 - curr_avg		
				curr_sr = float(sheet.cell(row_index,col_index+3).value)
				if curr_sr != 0.0:
					if 1 - curr_sr > delmax_sr:
						delmax_sr = 1.0 - curr_sr
				if curr_sr != 1.00:		
					if 1 - curr_sr < delmin_sr:
						delmin_sr 	= 1.0 - curr_sr	
				curr_hc = float(sheet.cell(row_index,col_index+4).value)
				if curr_hc != 0.0:
					if 1 - curr_hc > delmax_hc:
						delmax_hc = 1.0 - curr_hc
				if curr_hc != 1.00:		
					if 1 - curr_hc < delmin_hc:
						delmin_hc 	= 1.0 - curr_hc						
				curr_c = float(sheet.cell(row_index,col_index+5).value)
				if curr_c != 0.0:
					if 1 - curr_c > delmax_c:
						delmax_c = 1.0 - curr_c
				if curr_c != 1.00:		
					if 1 - curr_c < delmin_c:
						delmin_c 	= 1.0 - curr_c	
	#print(delmax_rpm,delmin_rpm,delmax_avg,delmin_avg,delmax_sr,delmin_sr,delmax_hc,delmin_hc,delmax_c,delmin_c)					
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):
		 if cellname(row_index,col_index)[0] == 'A':
		 	curr_batsmen = sheet.cell(row_index,col_index).value
		 	curr_rpm = float(sheet.cell(row_index,col_index+1).value)
		 	curr_avg = float(sheet.cell(row_index,col_index+2).value)
		 	curr_sr = float(sheet.cell(row_index,col_index+3).value)
		 	curr_hc = float(sheet.cell(row_index,col_index+4).value)
		 	curr_c = float(sheet.cell(row_index,col_index+5).value)
		 	grc_rpm[curr_batsmen] = "%.2f"%((delmin_rpm + (.5 * delmax_rpm)) / ((1-curr_rpm) + .5*delmax_rpm))
		 	grc_avg[curr_batsmen] = "%.2f"%((delmin_avg + (.5 * delmax_avg)) / ((1-curr_avg) + .5*delmax_avg))
		 	grc_sr[curr_batsmen] = "%.2f"%((delmin_sr + (.5 * delmax_sr)) / ((1-curr_sr) + .5*delmax_sr))
		 	grc_hc[curr_batsmen] = "%.2f"%((delmin_hc + (.5 * delmax_hc)) / ((1-curr_hc) + .5*delmax_hc))
		 	grc_c[curr_batsmen] = "%.2f"%((delmin_c + (.5 * delmax_c)) / ((1-curr_c) + .5*delmax_c))	
	#Create new sheet
	workbook,grey_relation_coefficients_sheet = create_new_sheet(
														"grey_relation_coefficients.xls"
														)

	#Initialize rows,columns
	row_count = 0
	grey_relation_coefficients_sheet.write(0,0,"Player Name")
	grey_relation_coefficients_sheet.write(0,1,"Runs per innings")
	grey_relation_coefficients_sheet.write(0,2,"AVG")
	grey_relation_coefficients_sheet.write(0,3,"SR")
	grey_relation_coefficients_sheet.write(0,4,"50's")
	grey_relation_coefficients_sheet.write(0,5,"100's")
	row_count+=1

	for x in grc_rpm:
		grey_relation_coefficients_sheet.write(row_count,0,x)
		grey_relation_coefficients_sheet.write(row_count,1,grc_rpm[x])
		grey_relation_coefficients_sheet.write(row_count,2,grc_avg[x])
		grey_relation_coefficients_sheet.write(row_count,3,grc_sr[x])
		grey_relation_coefficients_sheet.write(row_count,4,grc_hc[x])
		grey_relation_coefficients_sheet.write(row_count,5,grc_c[x])
		row_count+=1	

	workbook.close()	



def grd_calculation():
	w_rpm = 35
	w_avg = 30
	w_sr = 10
	w_hc = 10
	w_c = 15
	grd_batsmen = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_batsmen = sheet.cell(row_index,col_index).value
				curr_rpm = float(sheet.cell(row_index,col_index+1).value)
			 	curr_avg = float(sheet.cell(row_index,col_index+2).value)
			 	curr_sr = float(sheet.cell(row_index,col_index+3).value)
		 		curr_hc = float(sheet.cell(row_index,col_index+4).value)
		 		curr_c = float(sheet.cell(row_index,col_index+5).value)
				grd_batsmen[curr_batsmen] = (w_rpm * curr_rpm +
											w_avg * curr_avg +
											w_sr * curr_sr +
											w_hc * curr_hc +
											w_c * curr_c)
	#Create new sheet
	workbook,grey_relation_grades_sheet = create_new_sheet(
														"grey_relation_grades.xls"
														)

	#Initialize rows,columns
	row_count = 0		
	grey_relation_grades_sheet.write(0,0,"Player Name")
	grey_relation_grades_sheet.write(0,1,"Grade")
	row_count+=1
	for x in grd_batsmen:
		grey_relation_grades_sheet.write(row_count,0,x)
		grey_relation_grades_sheet.write(row_count,1,grd_batsmen[x])
		row_count+=1
	workbook.close()	


											
												


	
				

#normalize_complete_batsmen_stats()
#grc_calculation()
grd_calculation()