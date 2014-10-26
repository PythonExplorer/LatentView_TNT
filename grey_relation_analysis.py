#Libraries to parse xls docs
from xlrd import open_workbook,cellname

#Libraries to create xlx files for further use and data preparation
import xlsxwriter

def open(x):
	#Open data sheet
	book = open_workbook(x)
	#Index data sheet
	sheet = book.sheet_by_index(0)
	return sheet
def open_bat(x):
	#Open data sheet
	book = open_workbook("bat/"+x)
	#Index data sheet
	sheet = book.sheet_by_index(0)
	return sheet
def open_bowl(x):
	#Open data sheet
	book = open_workbook("bowl/"+x)
	#Index data sheet
	sheet = book.sheet_by_index(0)
	return sheet	

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
		try:
			if max_a < float(a[x]):
				max_a = float(a[x])
			if min_a > float(a[x]):
				min_a = float(a[x])
		except ValueError:
			pass			
	return (max_a,min_a)		


#Grey Relation Analysis


def normalize_complete_bowler_stats(yr):
	innings = {}
	wickets = {}
	economy = {}
	sr = {}
	avg = {}
	ftw = {}
	wptm = {}
	sheet = open_bowl("complete_bowler_stats"+str(yr)+".xls")
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):			
			if cellname(row_index,col_index)[0] == 'A' and (sheet.cell(row_index,col_index+4).value >= 1
														and sheet.cell(row_index,col_index+2).value >= 100
														and sheet.cell(row_index,col_index+1).value >= 3):
				curr_bowler = sheet.cell(row_index,col_index).value
				innings[curr_bowler] = sheet.cell(row_index,col_index+1).value
				wickets[curr_bowler] = sheet.cell(row_index,col_index+4).value
				avg[curr_bowler] = 100.0/float(sheet.cell(row_index,col_index+5).value)
				economy[curr_bowler] = 100.0/float(sheet.cell(row_index,col_index+6).value)
				sr[curr_bowler] = 100.0/float(sheet.cell(row_index,col_index+7).value)
				ftw[curr_bowler] = sheet.cell(row_index,col_index+9).value
	for x in innings:
		wptm[x] = (wickets[x]*10.0)/ innings[x]
		ftw[x] = (ftw[x]*10.0)/innings[x]
	max_wptm , min_wptm = min_max(wptm)
	max_avg , min_avg = min_max(avg)
	max_sr , min_sr = min_max(sr)
	max_ftw , min_ftw = min_max(ftw)
	max_eco , min_eco = min_max(economy)

	#Create new sheet
	workbook,normalized_complete_bowler_stats_sheet = create_new_sheet(
														"normalized_complete_bowler_stats"+str(yr)+".xls"
														)

	#Initialize rows,columns
	row_count = 0
	normalized_complete_bowler_stats_sheet.write(0,0,"Player Name")
	normalized_complete_bowler_stats_sheet.write(0,1,"Wickets per innings")
	normalized_complete_bowler_stats_sheet.write(0,2,"AVG")
	normalized_complete_bowler_stats_sheet.write(0,3,"SR")
	normalized_complete_bowler_stats_sheet.write(0,4,"Economy")
	normalized_complete_bowler_stats_sheet.write(0,5,"5W")
	row_count+=1

	for x in innings:
		normalized_complete_bowler_stats_sheet.write(row_count,0,x)
		if max_wptm == min_wptm:
			normalized_complete_bowler_stats_sheet.write(row_count,1,0)
		else:		
			normalized_complete_bowler_stats_sheet.write(row_count,1,
																"%.2f"%((wptm[x] - min_wptm) / (max_wptm - min_wptm))
																)
		if max_avg == min_avg:
			normalized_complete_bowler_stats_sheet.write(row_count,2,0)
		else:			
			normalized_complete_bowler_stats_sheet.write(row_count,2,
																"%.2f"%((float(avg[x]) - min_avg) / (max_avg - min_avg))															
																)
		if max_sr == min_sr:
			normalized_complete_bowler_stats_sheet.write(row_count,3,0)
		else:			
			normalized_complete_bowler_stats_sheet.write(row_count,3,
																"%.2f"%((float(sr[x]) - min_sr) / (max_sr - min_sr))
																)
		if max_eco == min_eco:
			normalized_complete_bowler_stats_sheet.write(row_count,4,0)
		else:			
			normalized_complete_bowler_stats_sheet.write(row_count,4,
																"%.2f"%((float(economy[x]) - min_eco) / (max_eco - min_eco))
																)
		if max_ftw == min_ftw:
			normalized_complete_bowler_stats_sheet.write(row_count,5,0)
		else:			
			normalized_complete_bowler_stats_sheet.write(row_count,5,
																"%.2f"%((float(ftw[x]) - min_ftw) / (max_ftw - min_ftw))
																)
		row_count+=1	
	workbook.close()

def normalize_complete_batsmen_stats(yr):
	sheet = open_bat("complete_batsmen_stats"+str(yr)+".xls")
	runs = {}
	innings = {}
	avg = {}
	sr = {}
	century = {}
	hcentury = {}
	cpm={}
	hcpm={}
	if yr == 2011:
		lt = 0
	elif yr == 2014:
		lt = 2
	else:
		lt = 5		
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):			
			if cellname(row_index,col_index)[0] == 'A' and sheet.cell(row_index,col_index+1).value >= lt:
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
		cpm[x] = (century[x]*1.0)/innings[x]
		hcpm[x] = (hcentury[x]*1.0)/innings[x]
	max_rpm , min_rpm = min_max(rpm)
	max_avg , min_avg = min_max(avg)
	max_sr , min_sr = min_max(sr)
	max_c , min_c = min_max(cpm)
	max_hc , min_hc = min_max(hcpm)

	#Create new sheet
	workbook,normalized_complete_batsmen_stats_sheet = create_new_sheet(
														"normalized_complete_batsmen_stats"+str(yr)+".xls"
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
		if max_rpm == min_rpm:
			normalized_complete_batsmen_stats_sheet.write(row_count,1,0)
		else:		
			normalized_complete_batsmen_stats_sheet.write(row_count,1,
																"%.2f"%((rpm[x] - min_rpm) / (max_rpm - min_rpm))
																)
		if max_avg == min_avg:
			normalized_complete_batsmen_stats_sheet.write(row_count,2,0)
		else:			
			normalized_complete_batsmen_stats_sheet.write(row_count,2,
																"%.2f"%((float(avg[x]) - min_avg) / (max_avg - min_avg))																
																)
		if max_sr == min_sr:
			normalized_complete_batsmen_stats_sheet.write(row_count,3,0)
		else:			
			normalized_complete_batsmen_stats_sheet.write(row_count,3,
																"%.2f"%((float(sr[x]) - min_sr) / (max_sr - min_sr))
																)
		if max_hc == min_hc:
			normalized_complete_batsmen_stats_sheet.write(row_count,4,0)
		else:			
			normalized_complete_batsmen_stats_sheet.write(row_count,4,
																"%.2f"%((hcpm[x] - min_hc) / (max_hc - min_hc))
																)
		if max_c == min_c:
			normalized_complete_batsmen_stats_sheet.write(row_count,5,0)
		else:			
			normalized_complete_batsmen_stats_sheet.write(row_count,5,
																"%.2f"%((cpm[x] - min_c) / (max_c - min_c))
																)
		row_count+=1	
	workbook.close()

def bowler_grc_calculation(yr):
	sheet = open("normalized_complete_bowler_stats"+str(yr)+".xls")
	grc_wptm = {}
	delmax_wptm = -1
	delmin_wptm = 1
	grc_avg = {}
	delmax_avg = -1
	delmin_avg = 1
	grc_sr = {}
	delmax_sr = -1
	delmin_sr = 1
	grc_eco= {}
	delmax_eco = -1
	delmin_eco = 1
	grc_ftw = {}
	delmax_ftw = -1
	delmin_ftw = 1
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_wptm = float(sheet.cell(row_index,col_index+1).value)
				if curr_wptm != 0.00:
					if 1 - curr_wptm > delmax_wptm:
						delmax_wptm = 1.0- curr_wptm
				if curr_wptm != 1.00:		
					if 1 - curr_wptm < delmin_wptm:
						delmin_wptm 	= 1.0 - curr_wptm
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
				curr_eco = float(sheet.cell(row_index,col_index+4).value)
				if curr_eco != 0.0:
					if 1 - curr_eco > delmax_eco:
						delmax_eco = 1.0 - curr_eco
				if curr_eco != 1.00:		
					if 1 - curr_eco < delmin_eco:
						delmin_eco 	= 1.0 - curr_eco						
				curr_ftw = float(sheet.cell(row_index,col_index+5).value)
				if curr_ftw != 0.0:
					if 1 - curr_ftw > delmax_ftw:
						delmax_ftw = 1.0 - curr_ftw
				if curr_ftw != 1.00:		
					if 1 - curr_ftw < delmin_ftw:
						delmin_ftw 	= 1.0 - curr_ftw	
	#print(delmax_rpm,delmin_rpm,delmax_avg,delmin_avg,delmax_sr,delmin_sr,delmax_hc,delmin_hc,delmax_c,delmin_c)					
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):
		 if cellname(row_index,col_index)[0] == 'A':
		 	curr_bowler = sheet.cell(row_index,col_index).value
		 	curr_wptm = float(sheet.cell(row_index,col_index+1).value)
		 	curr_avg = float(sheet.cell(row_index,col_index+2).value)
		 	curr_sr = float(sheet.cell(row_index,col_index+3).value)
		 	curr_eco = float(sheet.cell(row_index,col_index+4).value)
		 	curr_ftw = float(sheet.cell(row_index,col_index+5).value)
		 	grc_wptm[curr_bowler] = "%.2f"%((delmin_wptm + (.5 * delmax_wptm)) / ((1-curr_wptm) + .5*delmax_wptm))
		 	grc_avg[curr_bowler] = "%.2f"%((delmin_avg + (.5 * delmax_avg)) / ((1-curr_avg) + .5*delmax_avg))
		 	grc_sr[curr_bowler] = "%.2f"%((delmin_sr + (.5 * delmax_sr)) / ((1-curr_sr) + .5*delmax_sr))
		 	grc_eco[curr_bowler] = "%.2f"%((delmin_eco + (.5 * delmax_eco)) / ((1-curr_eco) + .5*delmax_eco))
		 	grc_ftw[curr_bowler] = "%.2f"%((delmin_ftw + (.5 * delmax_ftw)) / ((1-curr_ftw) + .5*delmax_ftw))	
	#Create new sheet
	workbook,grey_relation_coefficients_sheet = create_new_sheet(
														"bowler_grey_relation_coefficients"+str(yr)+".xls"
														)

	#Initialize rows,columns
	row_count = 0
	grey_relation_coefficients_sheet.write(0,0,"Player Name")
	grey_relation_coefficients_sheet.write(0,1,"Wickets per innings")
	grey_relation_coefficients_sheet.write(0,2,"AVG")
	grey_relation_coefficients_sheet.write(0,3,"SR")
	grey_relation_coefficients_sheet.write(0,4,"Economy")
	grey_relation_coefficients_sheet.write(0,5,"5W")
	row_count+=1

	for x in grc_wptm:
		grey_relation_coefficients_sheet.write(row_count,0,x)
		grey_relation_coefficients_sheet.write(row_count,1,grc_wptm[x])
		grey_relation_coefficients_sheet.write(row_count,2,grc_avg[x])
		grey_relation_coefficients_sheet.write(row_count,3,grc_sr[x])
		grey_relation_coefficients_sheet.write(row_count,4,grc_eco[x])
		grey_relation_coefficients_sheet.write(row_count,5,grc_ftw[x])
		row_count+=1	

	workbook.close()	
def batsmen_grc_calculation(yr):
	sheet = open("normalized_complete_batsmen_stats"+str(yr)+".xls")
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
														"grey_relation_coefficients"+str(yr)+".xls"
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


def bowler_grd_calculation(yr):
	sheet = open("bowler_grey_relation_coefficients"+str(yr)+".xls")
	w_wptm = 30
	w_avg = 15
	w_sr = 35
	w_eco = 10
	w_ftw = 10
	grd_bowler = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_bowler = sheet.cell(row_index,col_index).value
				curr_wptm = float(sheet.cell(row_index,col_index+1).value)
			 	curr_avg = float(sheet.cell(row_index,col_index+2).value)
			 	curr_sr = float(sheet.cell(row_index,col_index+3).value)
		 		curr_eco = float(sheet.cell(row_index,col_index+4).value)
		 		curr_ftw = float(sheet.cell(row_index,col_index+5).value)
				grd_bowler[curr_bowler] = (w_wptm * curr_wptm +
											w_avg * curr_avg +
											w_sr * curr_sr +
											w_eco * curr_eco +
											w_ftw * curr_ftw
											)
	#Create new sheet
	workbook,grey_relation_grades_sheet = create_new_sheet(
														"bowler_grey_relation_grades"+str(yr)+".xls"
														)

	#Initialize rows,columns
	row_count = 0		
	grey_relation_grades_sheet.write(0,0,"Player Name")
	grey_relation_grades_sheet.write(0,1,"Grade")
	row_count+=1
	for x in grd_bowler:
		grey_relation_grades_sheet.write(row_count,0,x)
		grey_relation_grades_sheet.write(row_count,1,grd_bowler[x])
		row_count+=1
	workbook.close()	
def batsmen_grd_calculation(yr):
	sheet = open("grey_relation_coefficients"+str(yr)+".xls")
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
														"grey_relation_grades"+str(yr)+".xls"
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

def complete_bowler_performance():
	sheet1 = open("bowler_grey_relation_grades2011.xls")
	sheet2 = open("bowler_grey_relation_grades2012.xls")
	sheet3 = open("bowler_grey_relation_grades2013.xls")
	sheet4 = open("bowler_grey_relation_grades2014.xls")
	total_bowler_grd = {}
	for row_index in range(1,sheet1.nrows):
		for col_index in range(sheet1.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_bowler = sheet1.cell(row_index,col_index).value
				curr_grade = sheet1.cell(row_index,col_index+1).value
				if curr_bowler not in total_bowler_grd:
					total_bowler_grd[curr_bowler] = 15*curr_grade
	for row_index in range(1,sheet2.nrows):
		for col_index in range(sheet2.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_bowler = sheet2.cell(row_index,col_index).value
				curr_grade = sheet2.cell(row_index,col_index+1).value
				if curr_bowler not in total_bowler_grd:
					total_bowler_grd[curr_bowler] = curr_grade				
				else:
					total_bowler_grd[curr_bowler] += 25*curr_grade
	for row_index in range(1,sheet3.nrows):
		for col_index in range(sheet3.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_bowler = sheet3.cell(row_index,col_index).value
				curr_grade = sheet3.cell(row_index,col_index+1).value
				if curr_bowler not in total_bowler_grd:
					total_bowler_grd[curr_bowler] = curr_grade				
				else:
					total_bowler_grd[curr_bowler] += 40*curr_grade					
	for row_index in range(1,sheet4.nrows):
		for col_index in range(sheet4.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_bowler = sheet4.cell(row_index,col_index).value
				curr_grade = sheet4.cell(row_index,col_index+1).value
				if curr_bowler not in total_bowler_grd:
					total_bowler_grd[curr_bowler] = curr_grade				
				else:
					total_bowler_grd[curr_bowler] += 20*curr_grade				

	#Create new sheet
	workbook,grey_relation_grades_sheet = create_new_sheet(
														"bowler_final_grey_relation_grades.xls"
														)

	#Initialize rows,columns
	row_count = 0		
	grey_relation_grades_sheet.write(0,0,"Player Name")
	grey_relation_grades_sheet.write(0,1,"Grade")
	row_count+=1
	for x in total_bowler_grd:
		grey_relation_grades_sheet.write(row_count,0,x)
		grey_relation_grades_sheet.write(row_count,1,total_bowler_grd[x]/100.0)
		row_count+=1
	workbook.close()										

def complete_batsmen_performance():
	sheet1 = open("grey_relation_grades2011.xls")
	sheet2 = open("grey_relation_grades2012.xls")
	sheet3 = open("grey_relation_grades2013.xls")
	sheet4 = open("grey_relation_grades2014.xls")
	total_batsmen_grd = {}
	for row_index in range(1,sheet1.nrows):
		for col_index in range(sheet1.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_batsmen = sheet1.cell(row_index,col_index).value
				curr_grade = sheet1.cell(row_index,col_index+1).value
				if curr_batsmen not in total_batsmen_grd:
					total_batsmen_grd[curr_batsmen] = 5*curr_grade
	for row_index in range(1,sheet2.nrows):
		for col_index in range(sheet2.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_batsmen = sheet2.cell(row_index,col_index).value
				curr_grade = sheet2.cell(row_index,col_index+1).value
				if curr_batsmen not in total_batsmen_grd:
					total_batsmen_grd[curr_batsmen] = curr_grade				
				else:
					total_batsmen_grd[curr_batsmen] += 35*curr_grade
	for row_index in range(1,sheet3.nrows):
		for col_index in range(sheet3.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_batsmen = sheet3.cell(row_index,col_index).value
				curr_grade = sheet3.cell(row_index,col_index+1).value
				if curr_batsmen not in total_batsmen_grd:
					total_batsmen_grd[curr_batsmen] = curr_grade				
				else:
					total_batsmen_grd[curr_batsmen] += 45*curr_grade					
	for row_index in range(1,sheet4.nrows):
		for col_index in range(sheet4.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_batsmen = sheet4.cell(row_index,col_index).value
				curr_grade = sheet4.cell(row_index,col_index+1).value
				if curr_batsmen not in total_batsmen_grd:
					total_batsmen_grd[curr_batsmen] = curr_grade				
				else:
					total_batsmen_grd[curr_batsmen] += 15*curr_grade				

	#Create new sheet
	workbook,grey_relation_grades_sheet = create_new_sheet(
														"final_grey_relation_grades.xls"
														)

	#Initialize rows,columns
	row_count = 0		
	grey_relation_grades_sheet.write(0,0,"Player Name")
	grey_relation_grades_sheet.write(0,1,"Grade")
	row_count+=1
	for x in total_batsmen_grd:
		grey_relation_grades_sheet.write(row_count,0,x)
		grey_relation_grades_sheet.write(row_count,1,total_batsmen_grd[x]/100.0)
		row_count+=1
	workbook.close()										
												




				
yrs = [2011,2012,2013,2014]
#for x in yrs:
#	normalize_complete_batsmen_stats(x)
#	batsmen_grc_calculation(x)
#	batsmen_grd_calculation(x)
#complete_batsmen_performance()	

for x in yrs:
	normalize_complete_bowler_stats(x)
	bowler_grc_calculation(x)
	bowler_grd_calculation(x)
complete_bowler_performance()	