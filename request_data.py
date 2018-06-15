import requests
import json
import pandas as pd
from openpyxl import load_workbook

# choose different sheet and read corresponding series id
sheet = ['A Data', 'B Data']
while True:
    sheet_name = input('Please enter sheet name you want to update: ')
    if sheet_name not in sheet:
        print('Please enter right sheet name!\n')
    else:
        break
seriesID = pd.read_csv('seriesID_' + sheet_name[0] + '.csv')
series_ids = seriesID['BLS SERIES ID'].tolist()

# the BLS API support 50 series query one time, divide series into groups with 50 series
seriesIDs = []
for i in range(0,3700,50):
    seriesIDs.append(series_ids[i:i+50])

seriesIDs.append(series_ids[3700:])

while True:
    try:
        year = int(input('Enter the the YEAR(YYYY) you want to update: '))
        if len(str(year)) != 4:
            print('Invalid value of year. Try again...')
            continue
        while True:
            month = int(input('Enter the the MONTH(MM) you want to update: '))
            if not 0 <= month <= 12:
                print('Invalid value of month. Try again...')
                continue
            else:
                break
        try:
            # use BLS API to request data with specific year and month
            headers = {'Content-type': 'application/json'}
            registrationkey = 'b77029e02d2e4db79365267fde5852fd'
            data = json.dumps({"seriesid":seriesIDs[0],"startyear":str(year), "endyear":str(year),"registrationkey":registrationkey})
            p = requests.post('https://api.bls.gov/publicAPI/v2/timeseries/data/', data=data, headers=headers)
            json_data = json.loads(p.text)
            break
        except:
            print('Error! Please try another year or month.')
    except:
        print("Oops!  That was no valid number.  Try again...")
print('Data are on the way! Please wait...')

df = pd.DataFrame(series for series in json_data['Results']['series'][0]['data'])
df['seriesID'] = json_data['Results']['series'][0]['seriesID']
n = len(df)

for i, series in enumerate(json_data['Results']['series']):
    df_data = pd.DataFrame(item for item in series['data'])
    df_data['seriesID'] = json_data['Results']['series'][i]['seriesID']
    df = df.append(df_data)

df = df.iloc[n:]

for seriesGroup in seriesIDs[1:]:
    headers = {'Content-type': 'application/json'}
    data = json.dumps({'seriesid': seriesGroup, 'startyear': str(year), 'endyear': str(year),
                       'registrationkey': registrationkey})
    p = requests.post('https://api.bls.gov/publicAPI/v2/timeseries/data/', data=data, headers=headers)
    json_data = json.loads(p.text)
    for i, series in enumerate(json_data['Results']['series']):
        df_data = pd.DataFrame(item for item in series['data'])
        df_data['seriesID'] = json_data['Results']['series'][i]['seriesID']
        df = df.append(df_data)
print('Data acquired! Writing into the worksheet...')

# filter the values by year and month
condition_1 = df.year == str(year)
condition_2 = df.period == 'M'+ str(month).zfill(2)
df_val = df[['value', 'seriesID']][condition_1 & condition_2]
seriesID = seriesID.rename(columns={'BLS SERIES ID': 'seriesID'})

# make our values in right order
df_comp = seriesID.merge(df_val, how='left', on='seriesID')

# write our new values into our existing excel in specific column to finish data updating
book = load_workbook('01 M1 BLS Data Drop Utility_2018.xlsx')
writer = pd.ExcelWriter('01 M1 BLS Data Drop Utility_2018.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

df_comp['value'].to_excel(writer, sheet_name, startcol=11+12*(year-2008)+month-1-1,
                          startrow=3, header=False, index=False)

writer.save()
print('Writing complete!')
