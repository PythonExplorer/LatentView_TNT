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


# Create Data Sheets
def create_new_sheet(sheet_name):
	workbook = xlsxwriter.Workbook(sheet_name)
	new_sheet = workbook.add_worksheet()
	return (workbook,new_sheet)


def winning_probabilities(sheet_name):
	sheet = open(sheet_name)
	teams = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(0,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				team1 = sheet.cell(row_index,col_index+1).value
				team2 = sheet.cell(row_index,col_index+2).value
				if team1 not in teams:
					teams[team1] = {}
				if team2 not in teams:
					teams[team2] = {}
				if team1 not in teams[team2]:
					teams[team2][team1] = {"Win":0,"Loss":0,"Tie":0,"No Result":0}
				if team2 not in teams[team1]:
					teams[team1][team2] = {"Win":0,"Loss":0,"Tie":0,"No Result":0}
				winning_team = 	sheet.cell(row_index,col_index+13).value
				if team1 == winning_team:
					teams[team1][team2]["Win"]+=1
					teams[team2][team1]["Loss"]+=1
				elif team2 == winning_team:
					teams[team1][team2]["Loss"]+=1
					teams[team2][team1]["Win"]+=1
				elif winning_team == "Tie":
					teams[team1][team2]["Tie"]+=1
					teams[team2][team1]["Tie"]+=1
				else:
					teams[team1][team2]["No Result"]+=1
					teams[team2][team1]["No Result"]+=1	
	#Create new sheet
	workbook,match_winner_sheet = create_new_sheet("match_winners.xls")
	#Initialize rows,columns
	row_count = 0
	match_winner_sheet.write(0,0,"Team Name")
	match_winner_sheet.write(0,1,"Opponent Name")
	match_winner_sheet.write(0,2,"Wins")
	match_winner_sheet.write(0,3,"Loss")
	match_winner_sheet.write(0,4,"Ties")
	match_winner_sheet.write(0,5,"No Results")
	match_winner_sheet.write(0,6,"P_Win")
	match_winner_sheet.write(0,7,"P_Loss")
	row_count+=1

	for x in teams:
		for y in teams[x]:
			try:
				match_winner_sheet.write(row_count,0,x)
				match_winner_sheet.write(row_count,1,y)
				match_winner_sheet.write(row_count,2,teams[x][y]["Win"])
				match_winner_sheet.write(row_count,3,teams[x][y]["Loss"])
				match_winner_sheet.write(row_count,4,teams[x][y]["Tie"])
				match_winner_sheet.write(row_count,5,teams[x][y]["No Result"])
				match_winner_sheet.write(row_count,6,"%.2f"%(teams[x][y]["Win"]*1.0/(teams[x][y]["Win"] + teams[x][y]["Loss"]
																						+teams[x][y]["Tie"]+teams[x][y]["No Result"])))
				match_winner_sheet.write(row_count,7,"%.2f"%(teams[x][y]["Loss"]*1.0/(teams[x][y]["Win"] + teams[x][y]["Loss"]
																						+teams[x][y]["Tie"]+teams[x][y]["No Result"])))
			except:
				print(x,y)
				exit()	
			row_count+=1
	workbook.close()
def team_avg_scores(sheet_name):
	teams={}
	sheet = open(sheet_name)
	for row_index in range(1,sheet.nrows):
		for col_index in range(0,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'D':
				bat1 = sheet.cell(row_index,col_index+2).value
				bat2 = sheet.cell(row_index,col_index+3).value
				if bat1 not in teams:
					teams[bat1] = [0,0,0,0,0,0]
				if bat2 not in teams:
					teams[bat2] = [0,0,0,0,0,0]
				score1 = float(sheet.cell(row_index,col_index+4).value)
				wik1 = float(sheet.cell(row_index,col_index+5).value)
				score2 = float(sheet.cell(row_index,col_index+7).value)
				wik2 = float(sheet.cell(row_index,col_index+8).value)
				teams[bat1][0]+=1
				teams[bat2][5]+=1
				teams[bat1][1]+=score1
				teams[bat2][3]+=score2
				teams[bat1][2]+=wik1
				teams[bat2][4]+=wik2
	#Create new sheet
	workbook,team_scores_sheet = create_new_sheet("team_scores.xls")
	#Initialize rows,columns
	row_count = 0
	team_scores_sheet.write(0,0,"Team Name")
	team_scores_sheet.write(0,1,"Total No of Matches")
	team_scores_sheet.write(0,2,"Avg score-1")
	team_scores_sheet.write(0,3,"Avg wkt-1")
	team_scores_sheet.write(0,4,"Avg score-2")
	team_scores_sheet.write(0,5,"Avg wkt-2")
	row_count+=1
	for x in teams:
		team_scores_sheet.write(row_count,0,x)
		team_scores_sheet.write(row_count,1,teams[x][0]+teams[x][5])
		if teams[x][0] != 0:
			team_scores_sheet.write(row_count,2,teams[x][1]//teams[x][0])
			team_scores_sheet.write(row_count,3,teams[x][2]//teams[x][0])
		else:
			team_scores_sheet.write(row_count,2,0)
			team_scores_sheet.write(row_count,3,0)
		if teams[x][5] != 0:	
			team_scores_sheet.write(row_count,4,teams[x][3]//teams[x][5])
			team_scores_sheet.write(row_count,5,teams[x][4]//teams[x][5])
		else:
			team_scores_sheet.write(row_count,4,0)
			team_scores_sheet.write(row_count,5,0)
		row_count+=1
	workbook.close()	

team_avg_scores("match/complete_match_stats.xls")
#winning_probabilities("match/complete_match_stats.xls")			





							
