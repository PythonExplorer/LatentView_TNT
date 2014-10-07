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

def map_player_runs():
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
	batsmen_stat_sheet.write(0,1,"Total Runs")
	batsmen_stat_sheet.write(0,2,"Number of Dismissals")
	batsmen_stat_sheet.write(0,3,"Number of Innings")
	row_count+=1
	for x in player_stats:
		batsmen_stat_sheet.write(row_count,0,x)
		batsmen_stat_sheet.write(row_count,1,player_stats[x][0])
		batsmen_stat_sheet.write(row_count,2,player_stats[x][1])
		batsmen_stat_sheet.write(row_count,3,player_stats[x][2])
		row_count+=1
	workbook.close()						



map_player_runs()
#map_matchid_team()
#map_team_match_count()




