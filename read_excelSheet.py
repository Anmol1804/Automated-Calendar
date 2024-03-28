# import pandas as pd
# df=pd.read_excel('sheets/scheduleFormatted.xlsx')
# print(df[0:2].to_markdown()) 

'''
# scraping values of sheet
import xlwings as xw 
wb = xw.Book('sheets/scheduleFormatted.xlsx') 
wks = xw.sheets 
ws = wks[0] 

# getting dates
# dates = [ws.range("A1").value,ws.range("D1").value,ws.range("G1").value,ws.range("J1").value,ws.range("M1").value] 
# print(dates) 
'''

# testing chourei data to push to calendar
import xlwings as xw 
wb = xw.Book('sheets/scheduleFormatted.xlsx') 
wks = xw.sheets 
ws = wks[0] 

# trying data for meet
morning_meet = [[ws.range("A1").value,ws.range("B2").value],[ws.range("D1").value,ws.range("E2").value]] 

event_ls=[]
for i in range(len(morning_meet)):
    event = {
            'summary': morning_meet[i][1],
            'location': 'Nishi shinjuku, japan',
            'description': morning_meet[i][0],
            'colorId': 3,
            'start': {
                # 'dateTime': '2024-03-29T09:00:00-07:00',
                'dateTime': '2024-03-' + morning_meet[i][0][2:4] + 'T09:00:00+03:30',
                'timeZone': 'Japan',
            },
            'end': {
                # 'dateTime': '2024-03-29T09:20:00-07:00',
                'dateTime': '2024-03-' + morning_meet[i][0][2:4] + 'T09:20:00+03:30',
                'timeZone': 'Japan',
            },
            'recurrence': [
                'RRULE:FREQ=DAILY;COUNT=1'
            ],
            'attendees': [
                {'email': 'anmolgera01@gmail.com'},
            ]
        }
    event_ls.append(event)

# print(morning_meet)

