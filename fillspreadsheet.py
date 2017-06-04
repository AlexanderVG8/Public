#!/usr/bin/env python 
import datetime
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from decimal import Decimal

GET_TOKEN_URL = "https://target-sandbox.my.com/api/v2/oauth2/token.json"
GET_DATA_URL = "https://target-sandbox.my.com/api/v1/campaigns.json?fields=id,status,stats_full"
CRED_FILE="API Project-39ab4ed270d4.json"
WRITER_EMAIL='selikhovalexey@gmail.com'
SPREADSHEET_NAME="spreadsheet_name"


def print_header(_sheet, _attempt_num, _sheet_num):
	_sheet.update_cell(1,1,"CampaingId" if _sheet_num==1 else "")
	_sheet.update_cell(1,2,"Date" if _sheet_num==1 else "Month")
	_sheet.update_cell(1,3,"Amount 1" if _sheet_num==1 else "Amount")
	_sheet.update_cell(1,4,"Amount 2" if _sheet_num==1 else "Clicks")
	_sheet.update_cell(1,5,"Clicks 1" if _sheet_num==1 else "Shows")
	_sheet.update_cell(1,6,"Clicks 2" if _sheet_num==1 else "")
	_sheet.update_cell(1,7,"Shows 1" if _sheet_num==1 else "")
	_sheet.update_cell(1,8,"Shows 2" if _sheet_num==1 else "")
	_sheet.update_cell(1,9,_attempt_num if _sheet_num==1 else "")


def print_attempt(_sheet, _json, _attempt_num, _sheet_num):
	_str_num = 2
	_month_data=dict()
	for arr1 in _json:
		for arr2 in arr1["stats_full"]:
			if _sheet_num==1:
				_sheet.update_cell(_str_num, 1, arr1["id"])
				_sheet.update_cell(_str_num, 2, arr2["date"])
				_sheet.update_cell(_str_num, 3 if _attempt_num == 1 else 4, arr2["amount"])
				_sheet.update_cell(_str_num, 5 if _attempt_num == 1 else 6, arr2["clicks"])
				_sheet.update_cell(_str_num, 7 if _attempt_num == 1 else 8, arr2["shows"])
				_str_num += 1
			else:
				_month_num = datetime.datetime.strptime(arr2["date"],'%d.%m.%Y').date().month
				if _month_data.get(_month_num) is None:
					_month_data.update({_month_num:{"amount":0,"clicks":0,"shows":0}})
				_month_data[_month_num]["amount"] += Decimal(arr2["amount"])
				_month_data[_month_num]["clicks"] += Decimal(arr2["clicks"])
				_month_data[_month_num]["shows"] += Decimal(arr2["shows"])
	if _sheet_num==2:
		print(_month_data)
		for month in _month_data:
			print(month)
			_sheet.update_cell(_str_num, 2, month)
			_sheet.update_cell(_str_num, 3, _month_data[month]["amount"])
			_sheet.update_cell(_str_num, 4, _month_data[month]["clicks"])
			_sheet.update_cell(_str_num, 5, _month_data[month]["shows"])			
			_str_num += 1

formdata = {
	"grant_type":"refresh_token",
	"refresh_token":"DrNhXUd1aEnfMJYksISAHNKRXmYAARw2M0ULwY1aNSZrR7xNvlVnQHjkUZuc9R5kiQCT3OONhLnGPYIeL7HXecBo6r1zbSE9wRdZEiDfx9Xyd60tcigkriaAxGt4TkdGbB6BUBjLGFof5yQFQrHQaejMLGNkgwg0PIITkkLl58Lf1AiEu2PYmNhWkYqpUHhaPZ82t5Jlxv",
	"client_secret":"zXHasS74dvjqdZG0rQjvTBsdHN2VkVnlkKpCwDgnMGh7p0y67mARKfsTZeXk8HVzYPWe2lcefb1P8VqzQHHCAPlISBM4S6pZKusOF82iZwgQ67AUvNxgTz2vkJIUzLWjLditYv5Os4uZQOcANBOhCCnAqO4JOVku4QwGhH5AeMb2sdkqPRXJRUddUX7GsQti9vW9nM1P5KjL5x8OFFox6eO6",
	"client_id":"IXaxGDNypkGRYdBm",
	"client_secret":"zXHasS74dvjqdZG0rQjvTBsdHN2VkVnlkKpCwDgnMGh7p0y67mARKfsTZeXk8HVzYPWe2lcefb1P8VqzQHHCAPlISBM4S6pZKusOF82iZwgQ67AUvNxgTz2vkJIUzLWjLditYv5Os4uZQOcANBOhCCnAqO4JOVku4QwGhH5AeMb2sdkqPRXJRUddUX7GsQti9vW9nM1P5KjL5x8OFFox6eO6"
}

try:
	r = requests.post(GET_TOKEN_URL, formdata)
except requests.exceptions.RequestException as e:
	print(e)
	sys.exit(1)

if r.status_code != 200:
	print("Token request is invalid or malformed")
	sys.exit(1)

access_token = r.json()["access_token"];

try:
	r = requests.get(GET_DATA_URL, headers={"Authorization":"Bearer {}".format(access_token)})
except requests.exceptions.RequestException as e:
	print(e)
	sys.exit(1)

if r.status_code != 200:
	print("Data request is invalid")
	sys.exit(1)

scope=['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, scope)

try:
	client = gspread.authorize(creds)
except:
	print(sys.exc_info()[0])
	sys.exit(1)

try:
	spreadsheet = client.open(SPREADSHEET_NAME)
	sheet = spreadsheet.sheet1
except gspread.SpreadsheetNotFound:
	print("Spreadsheet not found. Create...")
	spreadsheet = client.create(SPREADSHEET_NAME)
	spreadsheet.share(WRITER_EMAIL, perm_type='user', role='writer')
	sheet = spreadsheet.sheet1

sheet2 = spreadsheet.get_worksheet(1)
if sheet2 is None:
	sheet2 = spreadsheet.add_worksheet("month",512,10)
sheet2.clear()

if sheet.cell(1,9).value=='2':
	print("Clear sheet...")
	attempt_num = 1
	sheet.clear()
else:
	attempt_num = 2

print_header(sheet, attempt_num, 1)
print_attempt(sheet, r.json(), attempt_num, 1)
print_header(sheet2, attempt_num, 2)
print_attempt(sheet2, r.json(), attempt_num, 2)

print("Done.")