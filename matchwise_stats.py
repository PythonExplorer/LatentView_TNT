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

def toss_stats(sheet_name):
	sheet = open(sheet_name)
	team_toss_stats = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(0,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'B':
				curr_team =  sheet.cell(row_index,col_index).value
				opp_team = sheet.cell(row_index,col_index+1).value
				toss_winner = sheet.cell(row_index,col_index+2).value
				winner_decision = sheet.cell(row_index,col_index+3).value
				match_winner = sheet.cell(row_index,col_index+12).value
				if curr_team not in team_toss_stats:
					team_toss_stats[curr_team] = [0,0,0,0,0,0]
				if opp_team not in team_toss_stats:
					team_toss_stats[opp_team] = [0,0,0,0,0,0]
				team_toss_stats[curr_team][0]+=1
				team_toss_stats[opp_team][0]+=1
				if toss_winner == curr_team:
					team_toss_stats[curr_team][1]+=1
					if winner_decision == "bat":
						team_toss_stats[curr_team][2]+=1
						if match_winner == curr_team:
							team_toss_stats[curr_team][4]+=1
					if winner_decision == "field":
						team_toss_stats[curr_team][3]+=1
						if match_winner == curr_team:
							team_toss_stats[curr_team][5]+=1
				if toss_winner == opp_team:
					team_toss_stats[opp_team][1]+=1	
					if winner_decision == "bat":
						team_toss_stats[opp_team][2]+=1
						if match_winner == opp_team:
							team_toss_stats[opp_team][4]+=1
					if winner_decision == "feild":
						team_toss_stats[opp_team][3]+=1
						if match_winner == opp_team:
							team_toss_stats[opp_team][5]+=1
	#Create new sheet
	workbook,toss_stats_sheet = create_new_sheet("toss_stats.xls")
	#Initialize rows,columns
	row_count = 0
	toss_stats_sheet.write(0,0,"Team Name")
	toss_stats_sheet.write(0,1,"Total No of Matches")	
	toss_stats_sheet.write(0,2,"Toss Wins")
	toss_stats_sheet.write(0,3,"Toss wins bat")					
	toss_stats_sheet.write(0,4,"Toss wins bowl")
	toss_stats_sheet.write(0,5,"Toss wins bat win")
	toss_stats_sheet.write(0,6,"Toss wins bowl win")
	row_count+=1
	for x in team_toss_stats:
		toss_stats_sheet.write(row_count,0,x)
		for y in range(0,6):
			toss_stats_sheet.write(row_count,y+1,team_toss_stats[x][y])
		row_count+=1
	workbook.close()	

def ducks_stats(sheet_name):
	sheet = open(sheet_name)
	batsmen_duck_count = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(0,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'A':
				curr_player = sheet.cell(row_index,col_index+1).value
				player_runs = sheet.cell(row_index,col_index+2).value
				is_notout = sheet.cell(row_index,col_index+4).value
				if player_runs == 0 and is_notout == "NO":
					if curr_player not in batsmen_duck_count:
						batsmen_duck_count[curr_player]=0
					batsmen_duck_count[curr_player]+=1
	#Create new sheet
	workbook,ducks_stats_sheet = create_new_sheet("ducks_stats.xls")
	#Initialize rows,columns
	row_count = 0
	ducks_stats_sheet.write(0,0,"Player Name")
	ducks_stats_sheet.write(0,1,"Ducks Count")	
	row_count+=1
	for x in batsmen_duck_count:
		ducks_stats_sheet.write(row_count,0,x)	
		ducks_stats_sheet.write(row_count,1,batsmen_duck_count[x])
		row_count+=1
	workbook.close()							
							
def largest_margin(sheet_name):
	matchid = 0
	max_margin = -1
	sheet = open(sheet_name)
	for row_index in range(1,sheet.nrows):
		for col_index in range(0,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'H':
				curr_matchid = sheet.cell(row_index,col_index-7).value  
				score1 = sheet.cell(row_index,col_index).value
				score2 = sheet.cell(row_index,col_index+3).value
				result = sheet.cell(row_index,col_index+6).value
				if result != "No Result":
					if max_margin < abs(score1-score2):
						max_margin = abs(score1-score2)
						matchid = curr_matchid
	print(matchid)					

def extreme_totals(sheet_name):
	sheet = open(sheet_name)
	max_matchid = 0
	min_matchid = 0
	max_total = -1
	min_total = 500
	sheet = open(sheet_name)
	for row_index in range(1,sheet.nrows):
		for col_index in range(0,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'H':
				curr_matchid = sheet.cell(row_index,col_index-7).value 
				score1 = sheet.cell(row_index,col_index).value
				score2 = sheet.cell(row_index,col_index+3).value
				result = sheet.cell(row_index,col_index+6).value
				if result != "No Result":
					if max_total < max(score2,score1):
						max_total = max(score2,score1)
						max_matchid = curr_matchid
					if min_total > min(score2,score1):
						min_total = min(score2,score1)	
						min_matchid = curr_matchid
	print(max_matchid,min_matchid)					

def mom_count(sheet_name):
	sheet = open(sheet_name)
	player_mom_count = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(0,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'O':
				curr_player = sheet.cell(row_index,col_index).value
				if curr_player not in player_mom_count and curr_player != "":
					player_mom_count[curr_player] = 0
				if curr_player != "":	
					player_mom_count[curr_player]+=1

	#Create new sheet
	workbook,mom_count_sheet = create_new_sheet("mom_count.xls")
	#Initialize rows,columns
	row_count = 0
	mom_count_sheet.write(0,0,"Player Name")
	mom_count_sheet.write(0,1,"MOM Count")	
	row_count+=1
	for x in player_mom_count:
		mom_count_sheet.write(row_count,0,x)	
		mom_count_sheet.write(row_count,1,player_mom_count[x])
		row_count+=1
	workbook.close()								

mom_count("match/complete_match_stats.xls")
#extreme_totals("match/complete_match_stats.xls")
#largest_margin("match/complete_match_stats.xls")
#toss_stats("match/complete_match_stats.xls")
#team_avg_scores("match/complete_match_stats.xls")
#winning_probabilities("match/complete_match_stats.xls")			
#ducks_stats("bat/batsmen_match_stats.xls")




							
