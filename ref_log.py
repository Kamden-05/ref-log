import openpyxl
import requests
import json
import pandas as pd
import shutil
import xlsxwriter
import os

#TODO: change POSITION from Ref to Dual if game is a high school dual
#TODO: use Google Maps API to calculate mileage
#TODO: create pivot tables
#TODO: use command line arguments to choose output file
#TODO: modify API request to get games from a specific date range
#TODO: add error handling for API requests, file opening and writing,
#TODO: change file paths to be relative
#TODO: copy most recent log into a backup before modifying file
#TODO: convert xlsxwriter functions to openpyxl functions

# Name of excel file to write to
OUTPUT_FILE = 'test.xlsx'
BACKUP_FILE = 'backup.xlsx'

# API Keys
CLIENT_ID = os.getenv('ASSIGNR_CLIENT_ID')
CLIENT_SECRET = os.getenv('ASSIGNR_CLIENT_SECRET')

# list of column names that we want to keep in our final dataframe
DATE = 'Date'
TIME = 'Time'
VENUE = 'Venue'
LEAGUE = 'League'
AGE_GROUP = 'Age_Group'
HOME_TEAM = 'Home_Team'
AWAY_TEAM = 'Away_Team'
POSITION = 'Position'
ASSIGNOR = 'Assignor'
PAY_STATUS = 'Pay_Status'
FEE = 'Fee'

COLUMNS = [DATE, TIME, VENUE, LEAGUE,
           AGE_GROUP, HOME_TEAM, AWAY_TEAM,
           POSITION, ASSIGNOR, PAY_STATUS, FEE]

# Get auth token from assignr. Returns token for Get requests
def get_token():
    url = 'https://app.assignr.com/oauth/token'
    authData = {'client_id': CLIENT_ID, 
                'client_secret': CLIENT_SECRET,
                'scope': 'read',
                'grant_type': 'client_credentials'}

    response = requests.post(url, data=authData)
    token = response.json()
    return token['token_type'] + ' ' + token['access_token']

# Send get request to assignr for games. Returns game JSON data
def get_assignr_games(token):
    url = 'https://api.assignr.com/api/v2/current_account/games?page=1&limit=50'

    headers = {
        'accept': 'application/json',
        'authorization': token
    }

    response = requests.get(url, headers=headers)

    data = response.json()

    return data['_embedded']['games']

# Parse JSON game data and converts it to a dataframe ready to be written to spreadsheet
def flatten_game_json(gameJson):
    df = pd.json_normalize(gameJson)
    df = df.explode('_embedded.assignments')
    df = pd.json_normalize(json.loads(df.to_json(orient='records')))
    df = df.explode('_embedded.assignments._embedded.fees')
    return(pd.json_normalize(json.loads(df.to_json(orient='records'))))

# performs various pandas functions to clean the specific parts of the data I kepe
def clean_data(df):
    # rename important columns for readability
    df = df.rename(columns={'localized_date' : DATE, 'localized_time' : TIME,'_embedded.venue.name' : VENUE,
                            'league' : LEAGUE, 'age_group' : AGE_GROUP,'home_team' : HOME_TEAM,
                            'away_team' : AWAY_TEAM, '_embedded.assignments.position' : POSITION, 
                            '_embedded.assignments._embedded.fees.value' : FEE,
                            '_embedded.assignor.first_name' : 'assignor_first', '_embedded.assignor.last_name' : 'assignor_last',
                            '_embedded.assignments._embedded.official.first_name' : 'official_first',
                            '_embedded.assignments._embedded.official.last_name' : 'official_last',
                            '_embedded.site.name' : 'site'})

    # drop rows that contain data for other officials
    df = df.loc[(df['official_last'] == 'Wilson') &
                (df['official_first'] == 'Kamden')]
    
    df.insert(0, PAY_STATUS, 'Unpaid')

    # make certain leagues more readable
    df.loc[df[LEAGUE].str.contains('MLS'), LEAGUE] = 'MLS NEXT'
    df.loc[df[LEAGUE].str.contains('MLS'), 'gender'] = 'Boys'
    df.loc[df[LEAGUE].str.contains('TAPPS'), LEAGUE] = 'TAPPS'
    df.loc[df[LEAGUE].str.contains('North Texas Soccer'), LEAGUE] = 'NTX Soccer'
   
    # make certain age groups more readable
    df[AGE_GROUP] = df[AGE_GROUP].str.replace(' Mini', '', case=False)
    df[AGE_GROUP] = df[AGE_GROUP].str.replace(' Full', '', case=False)
    
    # add Chris Luna to assignor name if the game is from LunaCore Assigning
    df.loc[df['site'] == 'LunaCore Group Assigning', 'assignor_last'] = 'Luna'
    df.loc[df['site'] == 'LunaCore Group Assigning', 'assignor_first'] = 'Chris'

    # fill empty assignor name cells with names from previous rows
    df['assignor_last'] = df['assignor_last'].ffill()
    df['assignor_first'] = df['assignor_first'].ffill()

    # merge assignor first and last name columns into one column
    df[ASSIGNOR] = df[['assignor_first', 'assignor_last']].agg(' '.join, axis=1)

    # merge age_group and gender into one column
    df['gender'] = df['gender'].ffill()
    df[AGE_GROUP] = df[[AGE_GROUP, 'gender']].agg(' '.join, axis=1)

    # change Position cell values from 'Asst. Referee' to 'AR'
    df[POSITION] = df[POSITION].replace('Asst. Referee', 'AR')

    # set Date and Time columns to Datetime instead of str
    df[DATE] = pd.to_datetime(df[DATE])

    # filter out unecessary columns
    return(df[COLUMNS])

# Merges new games with already logged games into one dataframe
def merge_dataframe(df, file):
    main_df = pd.read_excel(open(file, 'rb'), sheet_name='Games')
    main_df = main_df.merge(df, how='outer')
    main_df[DATE] = pd.to_datetime(main_df[DATE]).dt.date
    return(main_df)  

def format_sheet(writer, num_rows):
    workbook = writer.book
    worksheet = writer.sheets['Games']
    worksheet.set_zoom(90)

    width = 20;

    header_format = workbook.add_format({
            "valign": "vcenter",
            "align": "center",
            "bg_color": "#951F06",
            "bold": True,
            'font_color': '#FFFFFF',
            'border' : 1, 
            'border_color': ''#D3D3D3'
        })

    # Full border formatting for conditional formatted cells
    full_border = workbook.add_format({ "border" : 1, "border_color": "#D3D3D3"})

    #add title
    title = "2024 Game Log"
    #merge cells
    format = workbook.add_format()
    format.set_font_size(20)
    format.set_font_color("#333333")

    worksheet.merge_range('A1:K1', title, format)
    worksheet.set_row(1, 15) # Set the header row height to 15
    # puting it all together
    # Write the column headers with the defined format.
    for col_num, value in enumerate(game_df.columns.values):
        #print(col_num, value)
        worksheet.write(1, col_num, value, header_format)

    # Adjust the column width.
    worksheet.set_column('A:K', width)

    # Add a number format for cells with money.
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    worksheet.set_column('K:K', 12, currency_format)

    # Highlight rows red under the condition that "Pay_Status" is "Unpaid"
    unpaid = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    formula = '=$J3="Unpaid"'
    worksheet.conditional_format('A3:K' + num_rows, {'type': 'formula', 'criteria': formula, 'value': 'Unpaid', 'format': unpaid })
    worksheet.conditional_format('A3:K' + num_rows, {'type': 'formula', 'criteria': formula, 'value': 'Unpaid', 'format': full_border })

    
token = get_token()
game_json = get_assignr_games(token)
game_df = flatten_game_json(game_json)
game_df = clean_data(game_df)
game_df = merge_dataframe(game_df, BACKUP_FILE)

#shutil.copy(OUTPUT_FILE, BACKUP_FILE)

writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
game_df.to_excel(writer, sheet_name='Games', index=False)

format_sheet(writer, str(len(game_df) + 1))

writer._save()