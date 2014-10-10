#Libraries to parse xls docs
from xlrd import open_workbook,cellname

#Libraries to create xlx files for further use and data preparation
import xlsxwriter


#Open data sheet
book = open_workbook('sample2.xls')

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
				curr_team = sheet.cell(row_index,col_index).value;
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
					matchid_players_run[curr_match_id][curr_batsmen]=[curr_runs,1-wide]
				if nonstricker_batsmen not in matchid_players_run[curr_match_id]:
					matchid_players_run[curr_match_id][nonstricker_batsmen] = [0,0]
	#Create new sheet
	workbook,batsmen_match_stat_sheet = create_new_sheet("batsmen_match_stats.xls")
	#Initialize rows,columns
	row_count = 0
	batsmen_match_stat_sheet.write(0,0,"Match ID")
	batsmen_match_stat_sheet.write(0,1,"Player")
	batsmen_match_stat_sheet.write(0,2,"Runs")
	batsmen_match_stat_sheet.write(0,3,"Balls Faced")	
	batsmen_match_stat_sheet.write(0,4,"Strike Rate")

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
				float((player_stats[x][0])/(player_stats[x][1])))
		else:
			batsmen_stat_sheet.write(row_count,4,0)		
		row_count+=1
	workbook.close()						

def match_stats():

def check_total_runs():
	total_runs=0
	for row_index in range(2,sheet.nrows):
		for col_index in range(1,sheet.ncols):
			if cellname(row_index,col_index)[0] == 'H':
				total_runs+=sheet.cell(row_index,col_index).value
	return total_runs			

	

#print(check_total_runs())

#map_batsmen_stats()
#map_matchid_team()
#map_team_match_count()

score_batsmen_match()




