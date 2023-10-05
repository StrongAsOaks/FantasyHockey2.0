'''
Created on Jul 15, 2020

@author: James
'''

#add a compare feature with expected results vs actual 

import xlsxwriter
import xlrd
import operator
import os
from datetime import datetime
from django.template.defaultfilters import center

locPlayerData = r"C:\Users\James\git\player-efficiency-calculator\PlayerEfficiencyCalculator\source\players.xlsx"
locTeamsData = r"C:\Users\James\git\player-efficiency-calculator\PlayerEfficiencyCalculator\source\teams.xlsx"
players = []
teams = []
sortedPlayers = []
computersTeam = []
excludedPlayers = []

salaryCap = 55000
salaryCapTracker = 0
wingCount = 0
centerCount = 0
defenseCount = 0
goalieCount = 0

wingCountMax = 4
centerCountMax = 2
DefenseCountMax = 2
GoalieCountMax = 1

class Player:
    def __init__(self, id, position, name, fantasyPointsPerGame, games, salary, team, opponent, injury, costEffciency, adjustedCostEfficiency):
        self.id = id
        self.position = position
        self.name = name
        self.fantasyPointsPerGame = round(fantasyPointsPerGame, 2)
        self.games = games
        self.salary = salary
        self.team = team
        self.opponent = opponent
        self.injury = injury
        self.fantasyPointsPerGame = fantasyPointsPerGame
        self.costEffciency = round(costEffciency, 2)
        self.adjustedCostEfficiency = adjustedCostEfficiency
        
class Team:
    def __init__(self, name, games, wins, losses, winLossPercent, GFPG, GAPG, PP, PK, SFPG, SAPG):
        self.name = name
        self.games = games
        self.wins = wins
        self.losses = losses
        self.winLossPercent = wins / games
        self.GFPG = GFPG
        self.GAPG = GAPG
        self.PP = PP
        self.PK = PK
        self.SFPG = SFPG
        self.SAPG = SAPG
    
def populateHeaders(sheet, workbook):
    # Add a header format.
    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})
    sheet.write(0,0, 'Id', header_format)
    sheet.write(0,1, 'Position', header_format)
    sheet.write(0,2, 'Name', header_format)
    sheet.write(0,3, 'FPPG', header_format)
    sheet.write(0,4, 'Games', header_format)
    sheet.write(0,5, 'Salary', header_format)
    sheet.write(0,6, 'Team', header_format)
    sheet.write(0,7, 'Opponent', header_format)
    sheet.write(0,8, 'Injury', header_format)
    sheet.write(0,9, 'CostEfficiency', header_format)
    sheet.write(0,10, 'AdjustedCostEfficiency', header_format)
    
    sheet.write(0,12, 'Id', header_format)
    sheet.write(0,13, 'Position', header_format)
    sheet.write(0,14, 'Name', header_format)
    sheet.write(0,15, 'FPPG', header_format)
    sheet.write(0,16, 'Games', header_format)
    sheet.write(0,17, 'Salary', header_format)
    sheet.write(0,18, 'Team', header_format)
    sheet.write(0,19, 'Opponent', header_format)
    sheet.write(0,20, 'Injury', header_format)
    sheet.write(0,21, 'CostEfficiency', header_format)
    sheet.write(0,22, 'AdjustedCostEfficiency', header_format)
    
    sheet.set_column(0, 0, 13)
    sheet.set_column(1, 1, 10)
    sheet.set_column(2, 2, 24)
    sheet.set_column(3, 3, 10)
    sheet.set_column(4, 4, 10)
    sheet.set_column(5, 5, 10)
    sheet.set_column(6, 6, 10)
    sheet.set_column(7, 7, 10)
    sheet.set_column(8, 8, 10)
    sheet.set_column(9, 9, 22)
    sheet.set_column(10, 10, 30)
    sheet.set_column(12, 12, 13)
    sheet.set_column(13, 13, 10)
    sheet.set_column(14, 14, 24)
    sheet.set_column(15, 15, 10)
    sheet.set_column(16, 16, 10)
    sheet.set_column(17, 17, 10)
    sheet.set_column(18, 18, 10)
    sheet.set_column(19, 19, 10)
    sheet.set_column(20, 20, 10)
    sheet.set_column(21, 21, 22)
    sheet.set_column(22, 22, 30)
    
def populateData(players, sheet):
    i = 1
    for Player in players:
        sheet.write(i,0, Player.id)
        sheet.write(i,1, Player.position)
        sheet.write(i,2, Player.name)
        sheet.write(i,3, Player.fantasyPointsPerGame)
        sheet.write(i,4, Player.games)
        sheet.write(i,5, Player.salary)
        sheet.write(i,6, Player.team)
        sheet.write(i,7, Player.opponent)
        sheet.write(i,8, Player.injury)
        sheet.write(i,9, Player.costEffciency)
        sheet.write(i,10, Player.adjustedCostEfficiency)
        
        i += 1
    
def abbreviateTeamNames(teams):
    for Team in teams:
        if Team.name == 'Anaheim Ducks':
            Team.name = 'ANA'
        elif Team.name == 'Arizona Coyotes':
            Team.name = 'ARI'
        elif Team.name == 'Boston Bruins':
            Team.name = 'BOS'
        elif Team.name == 'Buffalo Sabres':
            Team.name = 'BUF'
        elif Team.name == 'Carolina Hurricanes':
            Team.name = 'CAR'
        elif Team.name == 'Calgary Flames':
            Team.name = 'CGY'
        elif Team.name == 'Chicago Blackhawks':
            Team.name = 'CHI'
        elif Team.name == 'Columbus Blue Jackets':
            Team.name = 'CBJ'
        elif Team.name == 'Colorado Avalanche':
            Team.name = 'COL'
        elif Team.name == 'Dallas Stars':
            Team.name = 'DAL'
        elif Team.name == 'Detroit Red Wings':
            Team.name = 'DET'
        elif Team.name == 'Edmonton Oilers':
            Team.name = 'EDM'
        elif Team.name == 'Florida Panthers':
            Team.name = 'FLA'
        elif Team.name == 'Los Angeles Kings':
            Team.name = 'LAK'
        elif Team.name == 'Minnesota Wild':
            Team.name = 'MIN'
        elif "Canadiens" in Team.name:
            Team.name = 'MON'
        elif Team.name == 'Nashville Predators':
            Team.name = 'NSH'
        elif Team.name == 'New Jersey Devils':
            Team.name = 'NJD'
        elif Team.name == 'New York Islanders':
            Team.name = 'NYI'
        elif Team.name == 'New York Rangers':
            Team.name = 'NYR'
        elif Team.name == 'Ottawa Senators':
            Team.name = 'OTT'
        elif Team.name == 'Philadelphia Flyers':
            Team.name = 'PHI'
        elif Team.name == 'Pittsburgh Penguins':
            Team.name = 'PIT'
        elif Team.name == 'San Jose Sharks':
            Team.name = 'SJS'
        elif Team.name == 'St. Louis Blues':
            Team.name = 'STL'
        elif Team.name == 'Tampa Bay Lightning':
            Team.name = 'TBL'
        elif Team.name == 'Toronto Maple Leafs':
            Team.name = 'TOR'
        elif Team.name == 'Vancouver Canucks':
            Team.name = 'VAN'
        elif Team.name == 'Vegas Golden Knights':
            Team.name = 'VGK'
        elif Team.name == 'Winnipeg Jets':
            Team.name = 'WPG'
        elif Team.name == 'Washington Capitals':
            Team.name = 'WSH'
            
    return teams
               
def getTeamInfo():
    wb = xlrd.open_workbook(locTeamsData)
    sheet = wb.sheet_by_index(0)
    
    rows = sheet.nrows
    i = 1
    
    while (i < rows):
        teams.append(Team(sheet.cell_value(i,0),sheet.cell_value(i,2),sheet.cell_value(i,3),sheet.cell_value(i,4),'',sheet.cell_value(i,14),sheet.cell_value(i,15),sheet.cell_value(i,16),sheet.cell_value(i,17),sheet.cell_value(i,20), sheet.cell_value(i,21)))
        
        i += 1
    
    sortedTeams = sorted(teams, key=operator.attrgetter('wins'), reverse=True)
    
    finalTeams = abbreviateTeamNames(sortedTeams)
    
    return(finalTeams)

def getPlayerInfo():
    wb = xlrd.open_workbook(locPlayerData)
    sheet = wb.sheet_by_index(0)
    
    rows = sheet.nrows
    i = 1
    
    while (i < rows):
        if sheet.cell_value(i,5) != '' and sheet.cell_value(i,6) != '' and sheet.cell_value(i,5) != 0 and sheet.cell_value(i,6) > 5 and sheet.cell_value(i,11) == '':
            players.append(Player(sheet.cell_value(i,0),sheet.cell_value(i,1),sheet.cell_value(i,2) + ' ' + sheet.cell_value(i,4),sheet.cell_value(i,5),sheet.cell_value(i,6),sheet.cell_value(i,7),sheet.cell_value(i,9),sheet.cell_value(i,10),sheet.cell_value(i,11),(sheet.cell_value(i,5) * 10000) / sheet.cell_value(i,7), ''))
        
        i += 1

    return(players) 
    
def calculateAdjustedCostEffciency(players, teams): 
    oppShootingPercent = 1
    oppPP = 1
    oppSAPG = 1
    oppSFPG = 1
    oppGAPG = 1
    oppPK = 1
    
    for Player in players:
        if Player.position == 'G':
            for Team in teams:
                if Team.name == Player.opponent:
                    oppShootingPercent = Team.GFPG / Team.SFPG
                    oppPP = Team.PP / 10
            Player.adjustedCostEfficiency = Player.costEffciency / oppShootingPercent / oppPP / 4
            Player.adjustedCostEfficiency = round(Player.adjustedCostEfficiency, 2)
        else:
            for Team in teams:
                if Team.name == Player.opponent:
                    oppSAPG = Team.SAPG
                    oppSFPG = Team.SFPG
                    oppGAPG = Team.GAPG
                    oppPK = Team.PK
            Player.adjustedCostEfficiency = Player.costEffciency * oppSAPG * oppSFPG * oppGAPG / oppPK / 30
            Player.adjustedCostEfficiency = round(Player.adjustedCostEfficiency, 2)  
              
    
    #Sort by Position than adjustedCostEffciency
    players = sorted(players, key=operator.attrgetter('adjustedCostEfficiency'), reverse=True)          
    return players

def createPlayerSheet(players):
    #Get Date and Time
    now = datetime.now()
    nowString = now.strftime("%Y-%m-%d %H-%M-%S")
    nowDate = now.strftime("%Y-%m-%d")
    folderPath = fr"C:\Users\James\Documents\NHL Player Ratings\{nowDate}"
    
    if os.path.exists(folderPath):
        print('path exists')
    else:
        os.makedirs(folderPath)
    
    # Create a workbook and add a worksheet.
    wb = xlsxwriter.Workbook(fr'C:\Users\James\Documents\NHL Player Ratings\{nowDate}\PlayerCostEffciency2{nowString}.xlsx')
    sheet = wb.add_worksheet()

    populateHeaders(sheet, wb)
    
    populateData(players, sheet)
    
    computersTeam = generateComputerTeam(players)
    
    populateComputersTeam(computersTeam, sheet)

    wb.close()

def generateComputerTeam(players):
    #Add the next highest person and if the salary is full remove the highest on the team
    global goalieCount
    global centerCount
    global wingCount
    global defenseCount
    global salaryCapTracker
    
    if players[0].position == 'G':
        goalieCount = goalieCount + 1
    elif players[0].position == 'C':
        centerCount = centerCount + 1
    elif players[0].position == 'W':
        wingCount = wingCount + 1
    else:
        defenseCount = defenseCount + 1
    salaryCapTracker = salaryCapTracker + players[0].salary
    computersTeam.append(players[0])
    
    while len(computersTeam) <  9:
        for Player in players:
            if(Player not in excludedPlayers and Player not in computersTeam and logicCheckCanPlayerBeAdded(Player)):
                salaryCapTracker = salaryCapTracker + Player.salary
                computersTeam.append(Player)
                
        
        if salaryCapTracker > salaryCap:
            highestSalary = 0
            for Player in computersTeam:
                if Player.salary > highestSalary:
                    highestSalary = Player.salary
                    playerToBeRemoved = Player
            
            computersTeam.remove(playerToBeRemoved)
            salaryCapTracker = salaryCapTracker - playerToBeRemoved.salary
            if playerToBeRemoved.position == 'G':
                goalieCount = goalieCount - 1
            elif playerToBeRemoved.position == 'C':
                centerCount = centerCount - 1
            elif playerToBeRemoved.position == 'W':
                wingCount = wingCount - 1
            else:
                defenseCount = defenseCount -1
            excludedPlayers.append(playerToBeRemoved)
            
        else:
            break
           
    return computersTeam  
          
def logicCheckCanPlayerBeAdded(Player):
    global goalieCount
    global centerCount
    global wingCount
    global defenseCount
    global goalieCountMax
    global centerCountMax
    global wingCountMax
    global DefenseCountMax
    
    if Player.position == 'G' and goalieCount != GoalieCountMax:
        goalieCount = goalieCount + 1
        return True
    elif Player.position == 'C' and centerCount != centerCountMax:
        centerCount = centerCount + 1
        return True
    elif Player.position == 'W' and wingCount != wingCountMax:
        wingCount = wingCount + 1
        return True
    elif Player.position == 'D' and defenseCount != DefenseCountMax:
        defenseCount = defenseCount + 1
        return True
    else:
        return False

def populateComputersTeam(computersTeam, sheet):
    i = 1
    for Player in computersTeam:
        sheet.write(i,12, Player.id)
        sheet.write(i,13, Player.position)
        sheet.write(i,14, Player.name)
        sheet.write(i,15, Player.fantasyPointsPerGame)
        sheet.write(i,16, Player.games)
        sheet.write(i,17, Player.salary)
        sheet.write(i,18, Player.team)
        sheet.write(i,19, Player.opponent)
        sheet.write(i,20, Player.injury)
        sheet.write(i,21, Player.costEffciency)
        sheet.write(i,22, Player.adjustedCostEfficiency)
        
        i += 1
        
    sheet.write(10,17, salaryCapTracker)

def gatherInfo():
    teams = getTeamInfo()
    
    players = getPlayerInfo()

    players = calculateAdjustedCostEffciency(players, teams)
    
    createPlayerSheet(players)
    
    print("Calculations Completed")
        
    
gatherInfo()
    