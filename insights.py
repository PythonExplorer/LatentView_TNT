#Libraries to parse xls docs
from xlrd import open_workbook,cellname

#Libraries to create xlx files for further use and data preparation
import xlsxwriter


#Open data sheet
book = open_workbook('match/complete_match_stats.xls')

#Index data sheet
sheet = book.sheet_by_index(0)

# Create Data Sheets
def create_new_sheet(sheet_name):
	workbook = xlsxwriter.Workbook(sheet_name)
	new_sheet = workbook.add_worksheet()
	return (workbook,new_sheet)

def worldcup_stadium_stats():
	stadiums = {}
	for row_index in range(2,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'P':
				curr_stadium = sheet.cell(row_index,col_index).value
				curr_country = sheet.cell(row_index,col_index+1).value
				if curr_country != "":
					if curr_stadium not in stadiums:
						stadiums[curr_stadium]=[0,0,0,0,0,0]
					stadiums[curr_stadium][0]+=1
					curr_runs_1 = sheet.cell(row_index,col_index-8).value	
					curr_runs_2 = sheet.cell(row_index,col_index-5).value	
					batting_first = sheet.cell(row_index,col_index-10).value
					winner = sheet.cell(row_index,col_index-2).value
					if winner == "Tie" or winner == "No Result":
						stadiums[curr_stadium][3]+=1
					elif winner == batting_first:
						stadiums[curr_stadium][1]+=1	
					else:
						stadiums[curr_stadium][2]+=1	
					stadiums[curr_stadium][4]+=curr_runs_1
					stadiums[curr_stadium][5]+=curr_runs_2
					stadiums[curr_stadium].append(curr_country)
	for x in stadiums:
		stadiums[x][4] = int(stadiums[x][4]/stadiums[x][0])
		stadiums[x][5] = int(stadiums[x][5]/stadiums[x][0])
	#Create new sheet
	workbook,worldcup_stadium_stats_sheet = create_new_sheet("worldcup_stadium_stats.xls")
	#Initialize rows,columns
	row_count = 0
	worldcup_stadium_stats_sheet.write(0,0,"Stadium Name")
	worldcup_stadium_stats_sheet.write(0,1,"Number of Matches")
	worldcup_stadium_stats_sheet.write(0,2,"Batting First Wins")
	worldcup_stadium_stats_sheet.write(0,3,"Batting Second Wins")
	worldcup_stadium_stats_sheet.write(0,4,"No Winners")
	worldcup_stadium_stats_sheet.write(0,5,"AVG SCR 1")
	worldcup_stadium_stats_sheet.write(0,6,"AVG SCR 2")
	worldcup_stadium_stats_sheet.write(0,7,"Country")
	row_count+=1

	for x in stadiums:
		worldcup_stadium_stats_sheet.write(row_count,0,x)
		for y in range(0,7):
			worldcup_stadium_stats_sheet.write(row_count,y+1,stadiums[x][y])
		row_count+=1
	workbook.close()		

worldcup_stadium_stats()
