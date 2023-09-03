# %%
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import requests
import base64
import json

# %% Create a Tkinter window
root = tk.Tk()
root.title("Azure DevOps Task Manager")

max_window_width = root.winfo_screenwidth() 
root.maxsize(max_window_width, root.winfo_screenheight())

# Load the PAT from a JSON configuration file
def load_config():
    with open("config.json") as config_file:
        return json.load(config_file)

config = load_config()
pat = config.get("personal_access_token")
organization = config.get("organization")
project = config.get("project")
team = config.get("team")
authorization = str(base64.b64encode(bytes(':'+pat, 'ascii')), 'ascii')
headers = {
            'Accept': 'application/json',
            'Authorization': 'Basic '+ authorization,
            'Content-Type': 'application/json-patch+json'
        }

# Create a notebook with tabs for tasks and settings
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Tab 1: Task Management
task_frame = ttk.Frame(notebook)
notebook.add(task_frame, text="Task Management")

# Create a Treeview widget to display tasks from Excel
task_tree = ttk.Treeview(task_frame, columns=("Area Path", "Title", "Assigned To", "Due Month", "Due Day", "Client", "Estimated Effort"), show="headings")
task_tree.heading("#1", text="Area Path")
task_tree.heading("#2", text="Title")
task_tree.heading("#3", text="Assigned To")
task_tree.heading("#4", text="Due Month")
task_tree.heading("#5", text="Due Day")
task_tree.heading("#6", text="Client")
task_tree.heading("#7", text="Estimated Effort")
task_tree.pack()

# Create a horizontal scrollbar for the Treeview
horizontal_scrollbar = ttk.Scrollbar(task_frame, orient="horizontal", command=task_tree.xview)
horizontal_scrollbar.pack(fill="x")
task_tree.configure(xscrollcommand=horizontal_scrollbar.set)

# Load tasks from Excel and populate the Treeview
def load_tasks():
    try:
        # Clear the Treeview widget before loading new tasks
        task_tree.delete(*task_tree.get_children())
        df = pd.read_excel("DevopsRecurringItems.xlsx")
        for _, row in df.iterrows():
            task_tree.insert("", "end", values=(row["areapath"], row["title"], row["assignedto"], row["duemonth"], row["dueday"], row["Client"], row["Estimated Effort"]))
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Button to load tasks
load_button = ttk.Button(task_frame, text="Load Tasks", command=load_tasks)
load_button.pack()

# Tab 2: Settings
settings_frame = ttk.Frame(notebook)
notebook.add(settings_frame, text="Settings")

# Create and populate Entry widgets for settings
organization_entry = ttk.Entry(settings_frame, width=40)
organization_entry.insert(0, config.get("organization", ""))
project_entry = ttk.Entry(settings_frame, width=40)
project_entry.insert(0, config.get("project", ""))
team_entry = ttk.Entry(settings_frame, width=40)
team_entry.insert(0, config.get("team", ""))
pat_entry = ttk.Entry(settings_frame, width=40)
pat_entry.insert(0, config.get("personal_access_token", ""))

# Label and Entry for Organization
ttk.Label(settings_frame, text="Azure DevOps Organization:").pack()
organization_entry.pack()

# Label and Entry for Project
ttk.Label(settings_frame, text="Azure DevOps Project:").pack()
project_entry.pack()

# Label and Entry for Team
ttk.Label(settings_frame, text="Azure DevOps Team:").pack()
team_entry.pack()

# Label and Entry for Personal Access Token (PAT)
ttk.Label(settings_frame, text="Personal Access Token (PAT):").pack()
pat_entry.pack()

# Function to save settings
def save_settings():
    try:
        config["organization"] = organization_entry.get()
        config["project"] = project_entry.get()
        config["team"] = team_entry.get()
        config["personal_access_token"] = pat_entry.get()
        with open("config.json", "w") as config_file:
            json.dump(config, config_file)
        messagebox.showinfo("Success", "Settings saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Button to save settings
save_button = ttk.Button(settings_frame, text="Save Settings", command=save_settings)
save_button.pack()

def work_item_exists(title, area_path, iteration_path, assigned_to):
    config = load_config()
    pat = config.get("personal_access_token")
    organization = config.get("organization")
    project = config.get("project")
    team = config.get("team")
    authorization = str(base64.b64encode(bytes(':'+pat, 'ascii')), 'ascii')
    headers1 = {
        'Accept': 'application/json',
        'Authorization': 'Basic '+ authorization
    }
    query_url = f'https://dev.azure.com/{organization}/{project}/_apis/wit/wiql?api-version=7.0'
    query = {
        "query": f"SELECT [System.Id] FROM WorkItems WHERE [System.Title] = '{title}' "
                 f"AND [System.AreaPath] = '{area_path}' "
                 f"AND [System.IterationPath] = '{iteration_path}' "
                 f"AND [System.AssignedTo] = '{assigned_to}'"
    }
    response = requests.post(query_url, headers=headers1, json=query)
    if response.status_code == 200:
        data = response.json()
        if data.get('workItems'):
            return True
    return False

# Function to schedule tasks
def schedule_tasks():
    config = load_config()
    pat = config.get("personal_access_token")
    organization = config.get("organization")
    project = config.get("project")
    team = config.get("team")
    authorization = str(base64.b64encode(bytes(':'+pat, 'ascii')), 'ascii')
    headers1 = {
        'Accept': 'application/json',
        'Authorization': 'Basic '+ authorization
    }
    headers = {
        'Accept': 'application/json',
        'Authorization': 'Basic '+ authorization,
        'Content-Type': 'application/json-patch+json'
    }
    from datetime import datetime, timezone
    url = f'https://dev.azure.com/{organization}/{project}/{team}/_apis/work/teamsettings/iterations?api-version=7.0'
    response = requests.get(url, headers=headers1)
    data = response.json()
    df_sprint = pd.json_normalize(data['value'])
    df_sprint['startDate'] = pd.to_datetime(df_sprint['attributes.startDate'])
    df_sprint['finishDate'] = pd.to_datetime(df_sprint['attributes.finishDate'])

    # Get today's date
    today = pd.Timestamp.now(timezone.utc)

    # Find the sprint for today
    current_sprint = df_sprint[(df_sprint['startDate'] <= today) & (df_sprint['finishDate'] >= today)].iloc[0]['path']

    # Find the next sprint
    next_sprint = df_sprint[df_sprint['startDate'] >= today].sort_values('startDate').iloc[0]['path']
    next_sprint = df_sprint[df_sprint['startDate'] >= today].sort_values('startDate').iloc[1]['path']

    next_sprint_start = df_sprint[df_sprint['startDate'] >= today].sort_values('startDate').iloc[0]['startDate']
    next_sprint_end = df_sprint[df_sprint['startDate'] >= today].sort_values('startDate').iloc[1]['finishDate']
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

    url = f'https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/$Product%20Backlog%20Item?api-version=7.1-preview.3'


    for index, row in df.iterrows():
        title = row['title']
        area_path = row['areapath']
        assigned_to = row['assignedto']
        # Check if the work item already exists
        if not work_item_exists(title, area_path, next_sprint, assigned_to):
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

# Button to schedule tasks
schedule_button = ttk.Button(task_frame, text="Schedule Tasks", command=schedule_tasks)
schedule_button.pack()

# Load tasks when the application starts
load_tasks()

# Start the Tkinter main loop
root.mainloop()
