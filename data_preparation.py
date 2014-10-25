#Libraries to parse xls docs
from xlrd import open_workbook,cellname
import csv
#Libraries to create xlx files for further use and data preparation
import xlsxwriter


#Open data sheet
book = open_workbook('bowl/match_bowler_stats.xls')

#Index data sheet
sheet = book.sheet_by_index(0)


#Global Variables
stadium_match_list = {}
team_match_list = {}
match_ids = {}
total_matches = 0
row_count = 0
col_count = 0


#Sheet Details 
def sheet_details():
	print sheet.name
	print sheet.nrows
	print sheet.ncols

# Create Data Sheets
def create_new_sheet(sheet_name):
	workbook = xlsxwriter.Workbook(sheet_name)
	new_sheet = workbook.add_worksheet()
	return (workbook,new_sheet)


# Data Preparation functions

# Contains Match ID's and corresponding teams
def map_matchid_team():
	for row_index in range(2,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'C':
				curr_match_id = sheet.cell(row_index,col_index).value
				if not curr_match_id in match_ids:
					match_ids[curr_match_id] = []
			if cellname(row_index,col_index)[0] == 'L' or cellname(row_index,col_index)[0] == 'M':
				curr_team = sheet.cell(row_index,col_index).value
				if len(match_ids[curr_match_id]) < 2:
					match_ids[curr_match_id].append(curr_team)
	workbook,match_id_team_sheet = create_new_sheet("matchid_team.xls")
	#Initialize rows,columns
	row_count = 0
	match_id_team_sheet.write(0,0,"Match ID")
	match_id_team_sheet.write(0,1,"Team-1")
	match_id_team_sheet.write(0,2,"Team-2")
	row_count+=1
	for x in match_ids:
		match_id_team_sheet.write(row_count,0,x)
		match_id_team_sheet.write(row_count,1,match_ids[x][0])
		match_id_team_sheet.write(row_count,2,match_ids[x][1])
		row_count+=1
	workbook.close()


def map_team_match_count():
	for x in match_ids:
		for team in match_ids[x]:
			if team in team_match_list:
				team_match_list[team]+=1
			else:	
				team_match_list[team]=1

	workbook,team_match_sheet = create_new_sheet("team_match.xls")
	#Initialize rows,columns
	row_count = 0
	team_match_sheet.write(0,0,"Team")
	team_match_sheet.write(0,1,"Number of Matches")
	row_count+=1
	for x in team_match_list:
		team_match_sheet.write(row_count,0,x)
		team_match_sheet.write(row_count,1,team_match_list[x])
		row_count+=1
	workbook.close()


def score_batsmen_match():
	matchid_players_run = {}
	for row_index in range(2,sheet.nrows):
		for col_index in range(1,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'C':
				curr_match_id = sheet.cell(row_index,col_index).value	
				dismissed_batsmen = sheet.cell(row_index,col_index+11).value
				if curr_match_id not in matchid_players_run:
					matchid_players_run[curr_match_id]={}
				curr_batsmen = sheet.cell(row_index,col_index+2).value
				nonstricker_batsmen = sheet.cell(row_index,col_index+4).value
				curr_runs = sheet.cell(row_index,col_index+5).value
				curr_ball = sheet.cell(row_index,col_index+6).value
				if curr_ball > 0:
					wide = 1
				else:
					wide = 0	
				if curr_batsmen in matchid_players_run[curr_match_id]:
					matchid_players_run[curr_match_id][curr_batsmen][0]+=curr_runs
					matchid_players_run[curr_match_id][curr_batsmen][1]+=(1-wide)
				else:
					matchid_players_run[curr_match_id][curr_batsmen]=[curr_runs,1-wide,0]
				if nonstricker_batsmen not in matchid_players_run[curr_match_id]:
					matchid_players_run[curr_match_id][nonstricker_batsmen] = [0,0,0]
				if dismissed_batsmen!="" and dismissed_batsmen in matchid_players_run[curr_match_id]:
					matchid_players_run[curr_match_id][dismissed_batsmen][2] = 1
	#Create new sheet
	workbook,batsmen_match_stat_sheet = create_new_sheet("batsmen_match_stats.xls")
	#Initialize rows,columns
	row_count = 0
	batsmen_match_stat_sheet.write(0,0,"Match ID")
	batsmen_match_stat_sheet.write(0,1,"Player")
	batsmen_match_stat_sheet.write(0,2,"Runs")
	batsmen_match_stat_sheet.write(0,3,"Balls Faced")	
	batsmen_match_stat_sheet.write(0,4,"Strike Rate")
	batsmen_match_stat_sheet.write(0,4,"NO")
	row_count+=1
	for x in matchid_players_run:
		for y in matchid_players_run[x]:
			batsmen_match_stat_sheet.write(row_count,0,x)
			batsmen_match_stat_sheet.write(row_count,1,y)
			batsmen_match_stat_sheet.write(row_count,2,matchid_players_run[x][y][0])
			batsmen_match_stat_sheet.write(row_count,3,matchid_players_run[x][y][1])
			if (matchid_players_run[x][y][1]):
				batsmen_match_stat_sheet.write(row_count,4,
					"%.2f"%(float((matchid_players_run[x][y][0])/
						(matchid_players_run[x][y][1]))*100)
					)
			else:
				batsmen_match_stat_sheet.write(row_count,4,0)	
			if matchid_players_run[x][y][2] == 1:
				batsmen_match_stat_sheet.write(row_count,4,"NO")		
			else:
				batsmen_match_stat_sheet.write(row_count,4,"YES")		

			row_count+=1
	workbook.close()				



def map_batsmen_stats():
	player_stats = {}
	player_matchid = {}
	for row_index in range(2,sheet.nrows):
		for col_index in range(1,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'E':
				curr_batsmen = sheet.cell(row_index,col_index).value
				if curr_batsmen not in player_matchid:
					player_matchid[curr_batsmen] = []
				if curr_batsmen in player_stats:
					player_stats[curr_batsmen][0]+=sheet.cell(row_index,col_index+3).value
				else:
					player_stats[curr_batsmen] = [sheet.cell(row_index,col_index+3).value,0,0]

				curr_match_id = sheet.cell(row_index,col_index-2).value

				if curr_match_id not in player_matchid[curr_batsmen]: 
					player_stats[curr_batsmen][2]+=1
					player_matchid[curr_batsmen].append(curr_match_id)
			if cellname(row_index,col_index)[0] == 'N':
				dismissed_batsmen = sheet.cell(row_index,col_index).value
				if dismissed_batsmen in player_stats:
					player_stats[dismissed_batsmen][1]+=1
				else:
					player_stats[dismissed_batsmen] = [0,1,1]	
					curr_match_id = sheet.cell(row_index,col_index-11).value
					if dismissed_batsmen in player_matchid:
						player_matchid[dismissed_batsmen].append(curr_match_id)
					else:
						player_matchid[dismissed_batsmen]=[]
						player_matchid[dismissed_batsmen].append(curr_match_id)
						

	#Create new sheet
	workbook,batsmen_stat_sheet = create_new_sheet("batsmen_stats.xls")
	#Initialize rows,columns
	row_count = 0
	batsmen_stat_sheet.write(0,0,"Player Name")
	batsmen_stat_sheet.write(0,1,"Number of Innings")
	batsmen_stat_sheet.write(0,2,"Total Runs")
	batsmen_stat_sheet.write(0,3,"Not Outs")
	batsmen_stat_sheet.write(0,4,"Batting Average")

	row_count+=1
	for x in player_stats:
		batsmen_stat_sheet.write(row_count,0,x)
		batsmen_stat_sheet.write(row_count,1,player_stats[x][2])
		batsmen_stat_sheet.write(row_count,2,player_stats[x][0])
		batsmen_stat_sheet.write(row_count,3,player_stats[x][2]-player_stats[x][1])
		if (player_stats[x][1]):
			batsmen_stat_sheet.write(row_count,4,
				"%.2f"%float((player_stats[x][0])/(player_stats[x][1])))
		else:
			batsmen_stat_sheet.write(row_count,4,0)		
		row_count+=1
	workbook.close()						



def calc_runs_per_match():
	runs_per_match = {}
	for row_index in range(2,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'X':
				curr_match_id = sheet.cell(row_index,col_index-21).value
				curr_runs = sheet.cell(row_index,col_index).value
				is_wicket = sheet.cell(row_index,col_index+1).value
				curr_inn = sheet.cell(row_index,col_index-20).value
				is_wide = sheet.cell(row_index,col_index-3).value
				is_nb = sheet.cell(row_index,col_index-4).value
				if curr_match_id in runs_per_match:
					if curr_inn == 1:
						runs_per_match[curr_match_id][0]+=curr_runs
						if not (is_wide or is_nb):
							runs_per_match[curr_match_id][2]+=1

					else:
						runs_per_match[curr_match_id][3]+=curr_runs
						if not (is_wide or is_nb):
							runs_per_match[curr_match_id][5]+=1
					if is_wicket != "":
						if curr_inn == 1:
							runs_per_match[curr_match_id][1]+=1
						else:
							runs_per_match[curr_match_id][4]+=1
				else:
					runs_per_match[curr_match_id]=[]
					runs_per_match[curr_match_id].append(curr_runs)
					if is_wicket != "":
						runs_per_match[curr_match_id].append(1)
					else:
						runs_per_match[curr_match_id].append(0)
					if not (is_wide or is_nb):
						runs_per_match[curr_match_id].append(1)
					else:
						runs_per_match[curr_match_id].append(0)
					runs_per_match[curr_match_id].append(0)
					runs_per_match[curr_match_id].append(0)
					runs_per_match[curr_match_id].append(0)	
	return runs_per_match							

def map_matchid_year():
	id_date = {}
	with open('sample.csv', 'rU') as data:	
		reader = csv.reader(data)
		row = list(reader)
		for x in row[2:]:
			if x[1] not in id_date:
				id_date[int(x[1])] = x[5].split("/")[2]
	return id_date				

def complete_batsmen_stats():
	matchid_yr = map_matchid_year()
	yearly_batsmen_stats = {}
	for x in matchid_yr:
		curr_yr = matchid_yr[x]
		if curr_yr not in yearly_batsmen_stats:
			yearly_batsmen_stats[curr_yr] = {}
	for row_index in range(1,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'B':
				curr_batsmen = sheet.cell(row_index,col_index).value
				curr_match_id = int(sheet.cell(row_index,col_index-1).value	)
				curr_yr = matchid_yr[curr_match_id]
				if curr_batsmen not in yearly_batsmen_stats[curr_yr]:
					yearly_batsmen_stats[curr_yr][curr_batsmen] = [0,0,0,0,0,0,0]
				curr_runs = sheet.cell(row_index,col_index+1).value	
				curr_balls = sheet.cell(row_index,col_index+2).value
				is_notout = sheet.cell(row_index,col_index+3).value

				yearly_batsmen_stats[curr_yr][curr_batsmen][0]+=1
				yearly_batsmen_stats[curr_yr][curr_batsmen][1]+=curr_balls
				yearly_batsmen_stats[curr_yr][curr_batsmen][2]+=curr_runs
				if is_notout == "YES":
					yearly_batsmen_stats[curr_yr][curr_batsmen][3]+=1
				if curr_runs > yearly_batsmen_stats[curr_yr][curr_batsmen][4]:
					yearly_batsmen_stats[curr_yr][curr_batsmen][4]=curr_runs
				if curr_runs in range(50,100):
					yearly_batsmen_stats[curr_yr][curr_batsmen][5]+=1	
				if curr_runs >= 100:
					yearly_batsmen_stats[curr_yr][curr_batsmen][6]+=1					
	for xx in yearly_batsmen_stats:
		batsmen_stats = yearly_batsmen_stats[xx]
		#Create new sheet
		sheet_name = "complete_batsmen_stats"+str(xx)+".xls"
		workbook,complete_batsmen_stats_sheet = create_new_sheet(sheet_name)
		#Initialize rows,columns
		row_count = 0
		complete_batsmen_stats_sheet.write(0,0,"Player Name")
		complete_batsmen_stats_sheet.write(0,1,"Number of Innings")
		complete_batsmen_stats_sheet.write(0,2,"Balls Faced")
		complete_batsmen_stats_sheet.write(0,3,"Total Runs")
		complete_batsmen_stats_sheet.write(0,4,"NO")
		complete_batsmen_stats_sheet.write(0,5,"HS")
		complete_batsmen_stats_sheet.write(0,6,"AVG")
		complete_batsmen_stats_sheet.write(0,7,"SR")
		complete_batsmen_stats_sheet.write(0,8,"50's")
		complete_batsmen_stats_sheet.write(0,9,"100's")
		row_count+=1

		for x in batsmen_stats:
			for y in batsmen_stats[x]:
				complete_batsmen_stats_sheet.write(row_count,0,x)
				complete_batsmen_stats_sheet.write(row_count,1,batsmen_stats[x][0])
				complete_batsmen_stats_sheet.write(row_count,2,batsmen_stats[x][1])
				complete_batsmen_stats_sheet.write(row_count,3,batsmen_stats[x][2])
				complete_batsmen_stats_sheet.write(row_count,4,batsmen_stats[x][3])
				complete_batsmen_stats_sheet.write(row_count,5,batsmen_stats[x][4])
				dismissals = batsmen_stats[x][0]-batsmen_stats[x][3]
				if dismissals != 0:
					complete_batsmen_stats_sheet.write(row_count,6,"%.2f"%(float)(
																	batsmen_stats[x][2]/dismissals)
																	)
				else:
					complete_batsmen_stats_sheet.write(row_count,6,0)
				if 	batsmen_stats[x][1] != 0:
					complete_batsmen_stats_sheet.write(row_count,7,"%.2f"%(float)(
																	batsmen_stats[x][2]* 100/batsmen_stats[x][1]) 
																	)
				else:
					complete_batsmen_stats_sheet.write(row_count,7,0)
				complete_batsmen_stats_sheet.write(row_count,8,batsmen_stats[x][5])
				complete_batsmen_stats_sheet.write(row_count,9,batsmen_stats[x][6])
			row_count+=1		
		workbook.close()				



# Going to create a dictionary based on the unique match id. 
# Match ID : [Team1 , Team2 , Toss , Toss Decision , Team Batting 1st , Team Batting 2nd ,
#				Score-1 , RunRate - 1 , Score - 2 , RunRate - 2 , Winner , MOM , Venue ,
#				Country]

def match_stats():
	complete_match_stats={}
	runs_stats = {}
	venues_aus_wc15 = ["Sydney Cricket Ground","Melbourne Cricket Ground","Adelaide Oval",
						"Brisbane Cricket Ground, Woolloongabba","Western Australia Cricket Association Ground",
						"Bellerive Oval","Manuka Oval"
					]
	venues_nz_wc15	= ["Eden Park","Hagley Oval","Seddon Park","McLean Park","Westpac Stadium",
						"Saxton Oval","University Oval"
					]
	runs_per_match = calc_runs_per_match()
	for row_index in range(2,sheet.nrows):
		for col_index in range(sheet.ncols):
			if cellname(row_index,col_index)[0] == 'C':
				curr_match_id = sheet.cell(row_index,col_index).value
				team1 = sheet.cell(row_index,col_index+9).value
				team2 = sheet.cell(row_index,col_index+10).value
				toss_dec = sheet.cell(row_index,col_index+6).value
				toss_winner = sheet.cell(row_index,col_index+7).value
				batting_first = sheet.cell(row_index,col_index+2).value
				curr_venue = sheet.cell(row_index,col_index+8).value
				curr_mom = sheet.cell(row_index,col_index+11).value
				if team1 == batting_first:
					batting_second = team2
				else:
					batting_second = team1
				if curr_match_id not in complete_match_stats:
					complete_match_stats[curr_match_id] = [] 
					complete_match_stats[curr_match_id].append(team1)
					complete_match_stats[curr_match_id].append(team2)
					complete_match_stats[curr_match_id].append(toss_winner)
					complete_match_stats[curr_match_id].append(toss_dec)
					complete_match_stats[curr_match_id].append(batting_first)
					complete_match_stats[curr_match_id].append(batting_second)
					complete_match_stats[curr_match_id].append(runs_per_match[curr_match_id][0])
					complete_match_stats[curr_match_id].append(runs_per_match[curr_match_id][1])
					if (runs_per_match[curr_match_id][2]/6) != 0:
						complete_match_stats[curr_match_id].append("%.2f"%((runs_per_match[curr_match_id][0])/
																	(runs_per_match[curr_match_id][2]/6.0)
																))
					else:
						complete_match_stats[curr_match_id].append(0)
						
					complete_match_stats[curr_match_id].append(runs_per_match[curr_match_id][3])
					complete_match_stats[curr_match_id].append(runs_per_match[curr_match_id][4])
					if (runs_per_match[curr_match_id][5]/6) != 0:	
						complete_match_stats[curr_match_id].append("%.2f"%((runs_per_match[curr_match_id][3])/
																	(runs_per_match[curr_match_id][5]/6.0)
																))
					else:
						complete_match_stats[curr_match_id].append(0)
					if curr_mom == "":
						complete_match_stats[curr_match_id].append("No Result")	
					else:	
						if runs_per_match[curr_match_id][0]>runs_per_match[curr_match_id][3]:
							complete_match_stats[curr_match_id].append(batting_first)
						elif runs_per_match[curr_match_id][0]<runs_per_match[curr_match_id][3]:	
							complete_match_stats[curr_match_id].append(batting_second)
						else:	
							complete_match_stats[curr_match_id].append("Tie")
					complete_match_stats[curr_match_id].append(curr_mom)
					complete_match_stats[curr_match_id].append(curr_venue)
					if curr_venue in venues_aus_wc15:
						complete_match_stats[curr_match_id].append("AUS")
					elif curr_venue in venues_nz_wc15:
						complete_match_stats[curr_match_id].append("NZ")
					else:
						complete_match_stats[curr_match_id].append("")
						
	#Create new sheet
	workbook,complete_match_stats_sheet = create_new_sheet("complete_match_stats.xls")
	#Initialize rows,columns
	row_count = 0
	complete_match_stats_sheet.write(0,0,"Match ID")
	complete_match_stats_sheet.write(0,1,"Team1")
	complete_match_stats_sheet.write(0,2,"Team2")
	complete_match_stats_sheet.write(0,3,"Toss Winner")
	complete_match_stats_sheet.write(0,4,"Toss Decision")
	complete_match_stats_sheet.write(0,5,"Bat - 1")
	complete_match_stats_sheet.write(0,6,"Bat - 2")
	complete_match_stats_sheet.write(0,7,"Score - 1")
	complete_match_stats_sheet.write(0,8,"Wkt - 1")
	complete_match_stats_sheet.write(0,9,"RR - 1")
	complete_match_stats_sheet.write(0,10,"Score - 2")
	complete_match_stats_sheet.write(0,11,"Wkt - 2")
	complete_match_stats_sheet.write(0,12,"RR - 2")
	complete_match_stats_sheet.write(0,13,"Winner")
	complete_match_stats_sheet.write(0,14,"MOM")
	complete_match_stats_sheet.write(0,15,"Venue")
	complete_match_stats_sheet.write(0,16,"Country")
	row_count+=1
	for x in complete_match_stats:
		complete_match_stats_sheet.write(row_count,0,x)
		for y in range(1,17):
			complete_match_stats_sheet.write(row_count,y,complete_match_stats[x][y-1])
		row_count+=1	
	workbook.close()



def check_total_runs():
	total_runs=0
	for row_index in range(2,sheet.nrows):
		for col_index in range(1,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'H':
				total_runs+=sheet.cell(row_index,col_index).value
	return total_runs			

	

def map_match_bowler_stats():
	match_bowler_stats = {}
	for row_index in range(2,sheet.nrows):
		for col_index in range(1,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'C':
				curr_match_id = sheet.cell(row_index,col_index).value
				curr_bowler = sheet.cell(row_index,col_index+14).value
				curr_wides = sheet.cell(row_index,col_index+17).value
				curr_nb = sheet.cell(row_index,col_index+18).value
				curr_byes = sheet.cell(row_index,col_index+19).value
				curr_lbyes = sheet.cell(row_index,col_index+20).value
				total_runs = sheet.cell(row_index,col_index+21).value
				curr_runs = total_runs - curr_byes - curr_lbyes;
				if curr_match_id in match_bowler_stats:
					if curr_bowler in match_bowler_stats[curr_match_id]:
						if curr_wides == 0 and curr_nb == 0:
							match_bowler_stats[curr_match_id][curr_bowler][0]+=1
						match_bowler_stats[curr_match_id][curr_bowler][1]+=curr_runs
						is_wicket = sheet.cell(row_index,col_index+22).value
						if is_wicket != "run out" and is_wicket != "":
							match_bowler_stats[curr_match_id][curr_bowler][2]+=1
					else:
						match_bowler_stats[curr_match_id][curr_bowler]=	[0,0,0]
						if curr_wides == 0 and curr_nb == 0:
							match_bowler_stats[curr_match_id][curr_bowler][0]+=1
						match_bowler_stats[curr_match_id][curr_bowler][1]+=curr_runs
						is_wicket = sheet.cell(row_index,col_index+22).value
						if is_wicket != "run out" and is_wicket != "":
							match_bowler_stats[curr_match_id][curr_bowler][2]+=1
				else:
					match_bowler_stats[curr_match_id] = {}
					match_bowler_stats[curr_match_id][curr_bowler]=	[0,0,0]
					if curr_wides == 0 and curr_nb == 0:
						match_bowler_stats[curr_match_id][curr_bowler][0]+=1
					match_bowler_stats[curr_match_id][curr_bowler][1]+=curr_runs
					is_wicket = sheet.cell(row_index,col_index+22).value
					if is_wicket != "run out" and is_wicket != "":
						match_bowler_stats[curr_match_id][curr_bowler][2]+=1
	#Create new sheet
	workbook,match_bowler_stats_sheet = create_new_sheet("match_bowler_stats.xls")
	#Initialize rows,columns
	row_count = 0
	match_bowler_stats_sheet.write(0,0,"Match ID")
	match_bowler_stats_sheet.write(0,1,"Bowler")
	match_bowler_stats_sheet.write(0,2,"Balls")
	match_bowler_stats_sheet.write(0,3,"Runs")
	match_bowler_stats_sheet.write(0,4,"Wickets")
	row_count+=1

	#Write into workbook
	for x in match_bowler_stats:
		for y in match_bowler_stats[x]:
			match_bowler_stats_sheet.write(row_count,0,x)
			match_bowler_stats_sheet.write(row_count,1,y)
			match_bowler_stats_sheet.write(row_count,2,match_bowler_stats[x][y][0])
			match_bowler_stats_sheet.write(row_count,3,match_bowler_stats[x][y][1])
			match_bowler_stats_sheet.write(row_count,4,match_bowler_stats[x][y][2])
			row_count+=1
	workbook.close()		


def complete_bowler_stats():
	matchid_yr = map_matchid_year()
	yearly_bowler_stats = {}
	for x in matchid_yr:
		curr_yr = matchid_yr[x]
		if curr_yr not in yearly_bowler_stats:
			yearly_bowler_stats[curr_yr] = {}
	for row_index in range(2,sheet.nrows):
		for col_index in range(1,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'B':
				curr_bowler = sheet.cell(row_index,col_index).value
				curr_balls = sheet.cell(row_index,col_index+1).value
				curr_runs = sheet.cell(row_index,col_index+2).value
				curr_wickets = sheet.cell(row_index,col_index+3).value
				curr_match_id = int(sheet.cell(row_index,col_index-1).value	)
				curr_yr = matchid_yr[curr_match_id]
				if curr_bowler not in yearly_bowler_stats[curr_yr]:
					yearly_bowler_stats[curr_yr][curr_bowler] = [0,0,0,0,[0,0],0]
				yearly_bowler_stats[curr_yr][curr_bowler][0]+=1
				yearly_bowler_stats[curr_yr][curr_bowler][1]+=curr_balls
				yearly_bowler_stats[curr_yr][curr_bowler][2]+=curr_runs
				yearly_bowler_stats[curr_yr][curr_bowler][3]+=curr_wickets
				if curr_wickets > yearly_bowler_stats[curr_yr][curr_bowler][4][0]:
					yearly_bowler_stats[curr_yr][curr_bowler][4][0] = curr_wickets
					yearly_bowler_stats[curr_yr][curr_bowler][4][1] = curr_runs
				elif curr_wickets == yearly_bowler_stats[curr_yr][curr_bowler][4][0]:
					if curr_runs < yearly_bowler_stats[curr_yr][curr_bowler][4][1]:
						yearly_bowler_stats[curr_yr][curr_bowler][4][1] = curr_runs
				if curr_wickets >= 5:
					yearly_bowler_stats[curr_yr][curr_bowler][5]+=1	
	
	for xx in yearly_bowler_stats:
		bowler_stats = yearly_bowler_stats[xx]
		#Create new sheet
		sheet_name = "complete_bowler_stats"+str(xx)+".xls"
		#Create new sheet
		workbook,complete_bowler_stats_sheet = create_new_sheet(sheet_name)
		#Initialize rows,columns
		row_count = 0
		complete_bowler_stats_sheet.write(0,0,"Bowler")
		complete_bowler_stats_sheet.write(0,1,"Innings")
		complete_bowler_stats_sheet.write(0,2,"Balls")
		complete_bowler_stats_sheet.write(0,3,"Runs")
		complete_bowler_stats_sheet.write(0,4,"Wickets")
		complete_bowler_stats_sheet.write(0,5,"Avg")
		complete_bowler_stats_sheet.write(0,6,"Econ")
		complete_bowler_stats_sheet.write(0,7,"SR")
		complete_bowler_stats_sheet.write(0,8,"BBF")
		complete_bowler_stats_sheet.write(0,9,"5W")
		row_count+=1

		for x in bowler_stats:
			complete_bowler_stats_sheet.write(row_count,0,x)
			complete_bowler_stats_sheet.write(row_count,1,bowler_stats[x][0])
			complete_bowler_stats_sheet.write(row_count,2,bowler_stats[x][1])
			complete_bowler_stats_sheet.write(row_count,3,bowler_stats[x][2])
			complete_bowler_stats_sheet.write(row_count,4,bowler_stats[x][3])
			if bowler_stats[x][3] != 0:
				complete_bowler_stats_sheet.write(row_count,5,(
												"%.2f"%(float)(bowler_stats[x][2]/
												bowler_stats[x][3])
												)
											)
			else:
				complete_bowler_stats_sheet.write(row_count,5,"NA")
			if bowler_stats[x][1] != 0:	
				complete_bowler_stats_sheet.write(row_count,6,
												"%.2f"%(float)(bowler_stats[x][2]*6/
													bowler_stats[x][1]
													)
												)	
			else:
				complete_bowler_stats_sheet.write(row_count,6,0)
			if bowler_stats[x][3] != 0:
				complete_bowler_stats_sheet.write(row_count,7,(
												"%.2f"%(float)(bowler_stats[x][1]/
												bowler_stats[x][3])
												)
											)
			else:
				complete_bowler_stats_sheet.write(row_count,7,"NA")
			complete_bowler_stats_sheet.write(row_count,8,str(bowler_stats[x][4][0])+"/"+str(bowler_stats[x][4][1]))			
			complete_bowler_stats_sheet.write(row_count,9,bowler_stats[x][5])
			row_count+=1
		workbook.close()




#print(check_total_runs())
complete_bowler_stats()
#map_batsmen_stats()
#map_matchid_team()
#map_team_match_count()
#map_match_bowler_stats()
#score_batsmen_match()
#match_stats()
#complete_batsmen_stats()


