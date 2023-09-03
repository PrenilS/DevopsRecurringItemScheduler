# %% import packages
import requests
import base64
import pandas as pd
import json
#%%
# Load the PAT from a JSON configuration file
with open("config.json") as config_file:
    config = json.load(config_file)
    pat = config.get("personal_access_token")
    organization = config.get("organization")
    project = config.get("project")
    team = config.get("team")

# %% Auth header
authorization = str(base64.b64encode(bytes(':'+pat, 'ascii')), 'ascii')
headers = {
        'Accept': 'application/json',
        'Authorization': 'Basic '+ authorization
    }

# %% get iteration paths
from datetime import datetime, timezone
url = f'https://dev.azure.com/{organization}/{project}/{team}/_apis/work/teamsettings/iterations?api-version=7.0'
response = requests.get(url, headers=headers)
data = response.json()
df = pd.json_normalize(data['value'])
df['startDate'] = pd.to_datetime(df['attributes.startDate'])
df['finishDate'] = pd.to_datetime(df['attributes.finishDate'])

# Get today's date
today = pd.Timestamp.now(timezone.utc)

# Find the sprint for today
current_sprint = df[(df['startDate'] <= today) & (df['finishDate'] >= today)].iloc[0]['path']

# Find the next sprint
next_sprint = df[df['startDate'] > today].sort_values('startDate').iloc[0]['path']
next_sprint = df[df['startDate'] > today].sort_values('startDate').iloc[1]['path']

next_sprint_start = df[df['startDate'] > today].sort_values('startDate').iloc[0]['startDate']
next_sprint_end = df[df['startDate'] > today].sort_values('startDate').iloc[1]['finishDate']

# Print the results
print(current_sprint)
print(next_sprint, next_sprint_start, next_sprint_end)

# %% read in the excel sheet
df = pd.read_excel(r"DevopsRecurringItems.xlsx")
df.fillna('', inplace=True)
# %% loop through recurring items and add them to the next sprint
month_mapping = {
            'Jan': 1,
            'Feb': 2,
            'Mar': 3,
            'Apr': 4,
            'May': 5,
            'Jun': 6,
            'Jul': 7,
            'Aug': 8,
            'Sep': 9,
            'Oct': 10,
            'Nov': 11,
            'Dec': 12
        }

headers = {
        'Accept': 'application/json',
        'Authorization': 'Basic '+ authorization,
        'Content-Type': 'application/json-patch+json'
    }
url = f'https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/$Product%20Backlog%20Item?api-version=7.1-preview.3'


for index, row in df.iterrows():
    payload = [
        {
            'op': 'add',
            'path': '/fields/System.Title',
            'value': row['title']
        },
        {
            'op': 'add',
            'path': '/fields/System.AssignedTo',
            'value': row['assignedto']
        },
        {
            'op': 'add',
            'path': '/fields/System.AreaPath',
            'value': row['areapath']
        },
        {
            'op': 'add',
            'path': '/fields/System.IterationPath',
            'value': next_sprint
        },
        {
            'op': 'add',
            'path': '/fields/Custom.Client2',
            'value': row['Client']
        },
        {
            'op': 'add',
            'path': '/fields/Microsoft.VSTS.Scheduling.Effort',
            'value': row['Estimated Effort']
        }    
    ]

    # should add every month on the specified day
    if row['duemonth']=='All':
        if (next_sprint_start.day<=row['dueday']<=next_sprint_end.day) or \
            (next_sprint_end.day < next_sprint_start.day and (row['dueday'] >= next_sprint_start.day or row['dueday'] <= next_sprint_end.day)):
            print(payload)
            response = requests.patch(url, json=payload, headers=headers)

    # should add every week so just add it twice in every sprint
    elif row['duemonth']=='Weekly':
        print(payload)
        print(payload)
        response = requests.patch(url, json=payload, headers=headers)
        response = requests.patch(url, json=payload, headers=headers)

    # should add every sprint
    elif row['duemonth']=='Sprint':
        print(payload)
        response = requests.patch(url, json=payload, headers=headers)

    # only add on specific months
    else:
        months = row['duemonth'].split(',')
        months_as_numbers = [month_mapping[month] for month in months]
        if (next_sprint_start.month in months_as_numbers) or (next_sprint_end.month in months_as_numbers):
            if next_sprint_start.day<=row['dueday']<=next_sprint_end.day or \
            (next_sprint_end.day < next_sprint_start.day and (row['dueday'] >= next_sprint_start.day or row['dueday'] <= next_sprint_end.day)):
                print(payload)
                response = requests.patch(url, json=payload, headers=headers)
