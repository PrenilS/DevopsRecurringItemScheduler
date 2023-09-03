# DevOps Work Item Automation Script

This Python script streamlines the creation of work items in Azure DevOps by leveraging data from an Excel spreadsheet. It also incorporates advanced scheduling capabilities, allowing you to automate work item generation on a monthly, weekly, or specific month basis.

## Getting Started

To effectively utilize this script, you'll need to configure a config.json file. This configuration file now requires additional information: your Azure DevOps organization name, project name, and team ID. Here's an updated example of your config.json file:

```json
{
    "personal_access_token": "YOUR_PAT_HERE",
    "organization_name": "YOUR_ORG_NAME_HERE",
    "project_name": "YOUR_PROJECT_NAME_HERE",
    "team_id": "YOUR_TEAM_ID_HERE"
}
```
Replace "YOUR_PAT_HERE" with your actual PAT.

## Prerequisites

Before running the script, make sure you have the following prerequisites installed:

    Python: You need Python 3.x installed on your system.

    Required Python libraries: Install the required Python libraries using pip by running the following command:

    pip install requests pandas

## Running the Script

    Clone or download the script to your local machine.

    Create a config.json file as described in the "Getting Started" section and place it in the same directory as the script.

    Customize the script based on your needs. You can modify the Azure DevOps organization, project, team, and other parameters as required.

    Prepare your Excel spreadsheet with the work item data. The script assumes that the spreadsheet is located at DevopsRecurringItems.xlsx, but you can change this path in the script.

    Open a terminal or command prompt and navigate to the directory where the script is located.

    Run the script by executing the following command:

    ```shell
    python devops.py
    ```

    Replace devops.py with the actual name of your script.

The script will create work items in Azure DevOps based on the data from your Excel spreadsheet and the specified scheduling criteria.

## Scheduling Criteria

The script supports different scheduling criteria for work items:

    Monthly: Work items will be added to the specified sprint on the specified day of each month within the sprint's duration.

    Weekly: Work items will be added to the specified sprint twice within the sprint's duration, as a way to represent weekly scheduling.

    Sprint: Work items will be added to the specified sprint once, representing a one-time addition.

    Specific Months: Work items will be added to the specified sprint based on specific months provided in the spreadsheet. You can customize the mapping of month abbreviations to numbers in the script.

## Finding Azure DevOps Organization, Project, and Team Information

Before you can configure the config.json file, you need to obtain specific information about your Azure DevOps setup. Here's how to find the required details:
### Azure DevOps Organization Name:

    Log in to your Azure DevOps account.

    In the top right corner, click on your profile picture or initials.

    Select "Switch organizations" from the dropdown menu.

    You will see a list of organizations associated with your account. The name of the organization you want to use will be listed there. Note down the name of the organization you plan to work with.

### Azure DevOps Project Name:

    After selecting your desired organization, click on "Projects" in the left sidebar.

    You will see a list of projects within the organization. Locate the project you intend to work with and make note of its name.

### Azure DevOps Team ID:

    Inside your selected project, click on "Boards" in the left sidebar.

    In the "Teams" section, you will find a list of teams associated with your project.

    Select the team for which you want to automate work item creation.

    In the team's overview, you will find the Team ID in the URL. It typically looks like this:

    bash

    https://dev.azure.com/{ORG_NAME}/{PROJECT_NAME}/_settings/teams/teamsettings/{TEAM_ID}

    Note down the Team ID from the URL.

Now that you have collected the necessary information, you can proceed to fill in the config.json file as described in the "Getting Started" section of the README.

## Configuring the Excel Spreadsheet
To effectively use the script for automating work item creation, you'll need to prepare your Excel spreadsheet with the relevant data. The spreadsheet should include the following columns:

    1. areapath: This column should specify the area path where the work item will be assigned. The area path helps organize work items within your Azure DevOps project.

    2. title: Provide a title or name for the work item. This serves as a brief description of the task or item.

    3. assignedto: Specify the individual or team responsible for completing the work item. You can use usernames, display names, or other identifiers recognized by Azure DevOps.

    4. duemonth: Use this column to indicate the month in which the work item is due. You can use numerical values (e.g., 1 for January, 2 for February) or abbreviations (e.g., Jan, Feb) depending on your preference. Ensure that this value corresponds to your scheduling criteria.

    5. dueday: This column is used to specify the day of the month when the work item is due. For example, if the work item is due on the 15th of the month, you would enter "15" in this column.

    5. Client: Optionally, you can include a column for specifying the client or project associated with the work item.

    6. Estimated Effort: You can use this column to estimate the effort required to complete the work item. It's a numeric value that represents the expected workload.

Here's a step-by-step guide on how to use the duemonth and dueday columns for scheduling:

    Monthly Scheduling: If you want work items to recur monthly on a specific day, use the duemonth column to specify the month (e.g., "Jan" or "1") and the dueday column to specify the day (e.g., "15"). The script will automatically create work items on the specified day of each specified month.

    Weekly Scheduling: For weekly scheduling, you can set up the duemonth and dueday columns to create work items on specific days of the week. However, the script may require additional customization for this type of scheduling based on your specific needs.

    Sprint or One-Time Addition: If you intend for work items to be added just once or as a one-time addition, you can set the duemonth and dueday columns accordingly for the desired due date. You can leave the other rows empty or set them to N/A for one-time additions.

Ensure that your Excel spreadsheet is saved in a format that the script can access, and that it is correctly named as mentioned in the README. By configuring the spreadsheet with the necessary columns and scheduling information, you can seamlessly automate work item creation based on your criteria.