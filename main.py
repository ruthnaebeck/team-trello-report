import secrets
import trello
from sheets import update_sheet

import requests
import datetime as dt
import gspread

# Google Sheets imports / settings
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
from httplib2 import Http

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(secrets.google_json_file, scope)
gc = gspread.authorize(credentials)
service = build('sheets', 'v4', http=credentials.authorize(Http()))

wb = gc.open_by_key(secrets.google_sheet_id)
wks = wb.worksheet('Report')

# Table to organize trello card members by id
members = {}

def get_trello_member(member_id):
  try:
    member_name = members[member_id]
  except KeyError:
    member_json =  requests.request('GET', trello.url_members + member_id + trello.tokens).json()
    member_name = member_json['fullName']
    members[member_id] = member_name
  return {
    'id': member_id,
    'name': member_name,
  }

def main_script():
  # Duplicate the Report worksheet
  today = dt.datetime.today().strftime('%m-%d-%y')
  DATA = {'requests': [
    {
        'duplicateSheet': {
            'sourceSheetId': int(wks.id),
            'insertSheetIndex': 0,
            'newSheetName': 'Report ' + today
        }
    }
  ]}
  service.spreadsheets().batchUpdate(
        spreadsheetId=secrets.google_sheet_id, body=DATA).execute()

  # Collect Trello Card info
  table = []
  print('--------------------')
  for b in trello.trello_boards:
    for l in b['lists']:
      list_json = requests.request('GET', trello.url_lists + l['id'] + '/cards' + trello.tokens).json()
      print(b['name'])
      print('--------------------')
      x = 0
      for c in list_json:
        x += 1
        print(str(x) + ' - ' + c['name'])
        card_members = []
        if len(c['idMembers']) > 0:
          for id in c['idMembers']:
            card_member = get_trello_member(id)
            card_members.append(card_member['name'])
        new_row = [
          b['name'],
          l['name'],
          c['name'],
          c['shortUrl'],
          str(dt.datetime.fromtimestamp(int(c['id'][0:8],16))),
          str(dt.datetime.strptime(c['dateLastActivity'], '%Y-%m-%dT%H:%M:%S.%fZ')),
          ', '.join(card_members)
        ]

        # Append card / ticket to the table
        table.append(new_row)
      print('--------------------')

  # Add card info to the newly created worksheet
  new_wks = wb.worksheet('Report ' + today)
  update_sheet(new_wks, table)

  print('Complete!')

main_script()
