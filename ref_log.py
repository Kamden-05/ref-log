import openpyxl
import requests
import json
import pandas as pd
import shutil
import xlsxwriter
import os

# TODO: change POSITION from Ref to Dual if game is a high school dual
# TODO: use Google Maps API to calculate mileage
# TODO: create pivot tables
# TODO: use command line arguments to choose output file
# TODO: modify API request to get games from a specific date range
# TODO: change file paths to be relative
# TODO: copy most recent log into a backup before modifying file
# TODO: convert xlsxwriter functions to openpyxl functions
# TODO: remove hard coded cell references

# Name of excel file to write to
OUTPUT_FILE = "test.xlsx"
BACKUP_FILE = "backup.xlsx"

# API Keys
CLIENT_ID = os.getenv("ASSIGNR_CLIENT_ID")
CLIENT_SECRET = os.getenv("ASSIGNR_CLIENT_SECRET")

# dict of final olumn names
COLUMNS = {
    "date": "Date",
    "time": "Time",
    "venue": "Venue",
    "league": "League",
    "age": "Age_Group",
    "home": "Home_Team",
    "away": "Away_Team",
    "position": "Position",
    "assignor": "Assignor",
    "paid": "Pay_Status",
    "fee": "Fee",
}


# Returns assignr auth token for Get requests
def get_assignr_token():
    url = "https://app.assignr.com/oauth/token"
    authData = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "read",
        "grant_type": "client_credentials",
    }

    try:
        response = requests.post(url, data=authData, timeout=30)
        response.raise_for_status()
    except requests.exceptions.HTTPError as err:
        raise SystemExit(err)
    except requests.ConnectionError as err:
        raise SystemExit(err)
    except requests.Timeout as err:
        raise SystemExit(err)
    except requests.RequestException as err:
        raise SystemExit(err)
    except KeyboardInterrupt:
        raise SystemExit(err)

    token = response.json()
    return token["token_type"] + " " + token["access_token"]


# Send get request to assignr for games. Returns game JSON data
def get_assignr_games():
    url = "https://api.assignr.com/api/v2/current_account/games?page=1&limit=50"

    headers = {"accept": "application/json", "authorization": get_assignr_token()}

    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
    except requests.exceptions.HTTPError as err:
        raise SystemExit(err)
    except requests.ConnectionError as err:
        raise SystemExit(err)
    except requests.Timeout as err:
        raise SystemExit(err)
    except requests.RequestException as err:
        raise SystemExit(err)
    except KeyboardInterrupt:
        raise SystemExit(err)

    data = response.json()
    return data["_embedded"]["games"]


# Normalizes JSON game data and converts it to a dataframe
def flatten_json(gameJson):
    df = pd.json_normalize(gameJson)
    df = df.explode("_embedded.assignments")
    df = pd.json_normalize(json.loads(df.to_json(orient="records")))
    df = df.explode("_embedded.assignments._embedded.fees")
    return pd.json_normalize(json.loads(df.to_json(orient="records")))


# performs various pandas functions to clean the specific parts of the data I kepe
def clean_data(df):
    # rename important columns for readability
    df = df.rename(
        columns={
            "localized_date": COLUMNS["date"],
            "localized_time": COLUMNS["time"],
            "_embedded.venue.name": COLUMNS["venue"],
            "league": COLUMNS["league"],
            "age_group": COLUMNS["age"],
            "home_team": COLUMNS["home"],
            "away_team": COLUMNS["away"],
            "_embedded.assignments.position": COLUMNS["position"],
            "_embedded.assignments._embedded.fees.value": COLUMNS["fee"],
            "_embedded.assignor.first_name": "assignor_first",
            "_embedded.assignor.last_name": "assignor_last",
            "_embedded.assignments._embedded.official.first_name": "official_first",
            "_embedded.assignments._embedded.official.last_name": "official_last",
            "_embedded.site.name": "site",
        }
    )

    # drop rows that contain data for other officials
    df = df.loc[(df["official_last"] == "Wilson") & (df["official_first"] == "Kamden")]

    df.insert(0, COLUMNS["paid"], "Unpaid")

    # make certain leagues more readable
    df.loc[df[COLUMNS["league"]].str.contains("MLS"), COLUMNS["league"]] = "MLS NEXT"
    df.loc[df[COLUMNS["league"]].str.contains("MLS"), "gender"] = "Boys"
    df.loc[df[COLUMNS["league"]].str.contains("TAPPS"), COLUMNS["league"]] = "TAPPS"
    df.loc[
        df[COLUMNS["league"]].str.contains("North Texas Soccer"), COLUMNS["league"]
    ] = "NTX Soccer"

    # make certain age groups more readable
    df[COLUMNS["age"]] = df[COLUMNS["age"]].str.replace(" Mini", "", case=False)
    df[COLUMNS["age"]] = df[COLUMNS["age"]].str.replace(" Full", "", case=False)

    # add Chris Luna to assignor name if the game is from LunaCore Assigning
    df.loc[df["site"] == "LunaCore Group Assigning", "assignor_last"] = "Luna"
    df.loc[df["site"] == "LunaCore Group Assigning", "assignor_first"] = "Chris"

    # fill empty assignor name cells with names from previous rows
    df["assignor_last"] = df["assignor_last"].ffill()
    df["assignor_first"] = df["assignor_first"].ffill()

    # merge assignor first and last name columns into one column
    df[COLUMNS["assignor"]] = df[["assignor_first", "assignor_last"]].agg(
        " ".join, axis=1
    )

    # merge age_group and gender into one column
    df["gender"] = df["gender"].ffill()
    df[COLUMNS["age"]] = df[[COLUMNS["age"], "gender"]].agg(" ".join, axis=1)

    # change Position cell values from 'Asst. Referee' to 'AR'
    df[COLUMNS["position"]] = df[COLUMNS["position"]].replace("Asst. Referee", "AR")

    # set Date and Time columns to Datetime instead of str
    df[COLUMNS["date"]] = pd.to_datetime(df[COLUMNS["date"]])

    # filter out unecessary columns
    return df[list(COLUMNS.values())]


# Merges new games with already logged games into one dataframe
def merge_dataframe(df, file):
    main_df = pd.read_excel(
        open(file, "rb"), sheet_name="Games", header=1
    )  # header row is 0-indexed
    main_df = main_df.merge(df, how="outer")
    main_df[COLUMNS["date"]] = pd.to_datetime(main_df[COLUMNS["date"]]).dt.date
    return main_df


def format_sheet(writer, num_rows):
    workbook = writer.book
    worksheet = writer.sheets["Games"]
    worksheet.set_zoom(90)

    width = 20

    header_format = workbook.add_format(
        {
            "valign": "vcenter",
            "align": "center",
            "bg_color": "#951F06",
            "bold": True,
            "font_color": "#FFFFFF",
            "border": 1,
            "border_color": "",  # D3D3D3'
        }
    )

    # Full border formatting for conditional formatted cells
    full_border = workbook.add_format({"border": 1, "border_color": "#D3D3D3"})

    # add title
    title = "2024 Game Log"
    # merge cells
    format = workbook.add_format()
    format.set_font_size(20)
    format.set_font_color("#333333")

    worksheet.merge_range("A1:K1", title, format)
    worksheet.set_row(1, 15)  # Set the header row height to 15

    # Write the column headers with the defined format.
    # TODO: pass dataframe in as function arg
    for col_num, value in enumerate(game_df.columns.values):
        # print(col_num, value)
        worksheet.write(1, col_num, value, header_format)

    # Adjust the column width.
    worksheet.set_column("A:K", width)

    # Add a number format for cells with money.
    currency_format = workbook.add_format({"num_format": "$#,##0.00"})
    worksheet.set_column("K:K", 12, currency_format)

    # Highlight rows red under the condition that "Pay_Status" is "Unpaid"
    unpaid = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    formula = '=$J3="Unpaid"'
    worksheet.conditional_format(
        "A3:K" + num_rows,
        {"type": "formula", "criteria": formula, "value": "Unpaid", "format": unpaid},
    )
    worksheet.conditional_format(
        "A3:K" + num_rows,
        {
            "type": "formula",
            "criteria": formula,
            "value": "Unpaid",
            "format": full_border,
        },
    )


game_json = get_assignr_games()
game_df = flatten_json(game_json)
game_df = clean_data(game_df)

shutil.copy(OUTPUT_FILE, BACKUP_FILE)
game_df = merge_dataframe(game_df, OUTPUT_FILE)

writer = pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter")
game_df.to_excel(writer, sheet_name="Games", index=False)

format_sheet(writer, str(len(game_df) + 1))

writer._save()
