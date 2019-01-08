from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import os
import math
import datetime

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
SPREADSHEET_ID = '1XV7kSXzdo_Y65Sfajxn0DmeblIoBrMcdIpjvdlUB3tg'
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = build('sheets', 'v4', http=creds.authorize(Http()))
sheet = service.spreadsheets()
alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def numToRow(num):
	return alphabet[num%26]*(math.floor(num/26)+1)

def newID(id, name):
	r = 'Sheet1!B:B'
	list = [[id],[name]]
	resource = {
		"majorDimension": "COLUMNS",
		"values": list
	}
	result = sheet.values().append(
		spreadsheetId=SPREADSHEET_ID,
		range=r,
		body=resource,
		valueInputOption="USER_ENTERED"
	).execute()

def newDate(time):
	RANGE = 'Sheet1!A1:'+numToRow(60)+'1'
	result = sheet.values().get(
	spreadsheetId=SPREADSHEET_ID,
	range=RANGE).execute()
	values = result.get('values', [])

	r = 'Sheet1!'+numToRow(len(values[0]))+"1:"+numToRow(60)+'1'
	list = [[str(time.month)+"/"+str(time.day) + " IN"],[str(time.month)+"/"+str(time.day)+" OUT"]]
	resource = {
		"majorDimension": "COLUMNS",
		"values": list
	}
	result = sheet.values().update(
		spreadsheetId=SPREADSHEET_ID,
		range=r,
		body=resource,
		valueInputOption="USER_ENTERED"
	).execute()

def columnContains(check, column):
	RANGE = 'Sheet1!'+column+':'+column
	result = sheet.values().get(
	spreadsheetId=SPREADSHEET_ID,
	range=RANGE).execute()
	values = result.get('values', [])
	for v in values:
		if (v[0]==check):
			return values.index(v)
	return 0

def rowContains(check, row):
	RANGE = 'Sheet1!A'+row+":"+numToRow(60)+row
	result = sheet.values().get(
	spreadsheetId=SPREADSHEET_ID,
	range=RANGE).execute()
	values = result.get('values', [])
	for v in values[0]:
		if (v==check):
			return values[0].index(v)
	return 0

def signIn(index,time):
	column = rowContains(str(time.month)+"/"+str(time.day) + " IN","1")
	if(column<2):
		newDate(time)

	column = rowContains(str(time.month)+"/"+str(time.day) + " IN","1")
	r = 'Sheet1!'+numToRow(column)+str(index+1)
	list = [[str(time.hour)+":"+str(time.minute)]]
	resource = {
		"majorDimension": "COLUMNS",
		"values": list
	}
	result = sheet.values().update(
		spreadsheetId=SPREADSHEET_ID,
		range=r,
		body=resource,
		valueInputOption="USER_ENTERED"
	).execute()

	#RANGE = 'Sheet1!'+row+':'+row
	#result = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=, valueInputOption=value_input_option, body=value_range_body).execute()


running = 1
while running:
	print('Scan ID Now')
	user_id = input()
	if user_id:
		os.system('cls')
		idExists=columnContains(user_id, numToRow(0))
		if (idExists<1):	
			print("Looks like you haven't signed in before. Enter your name: ")
			name = input()
			newID(user_id, name)
		idExists=columnContains(user_id, numToRow(0))

		day = datetime.datetime.now()
		signIn(idExists,day)

		os.system('cls')
		print('Signed in: ' + user_id)
	else: 
		break;
	
print("Done signing in")