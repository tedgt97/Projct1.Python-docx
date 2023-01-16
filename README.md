# Projct1: Python-docx 
### **Monthly Audit Issue Tracker by using Python-docx package**

<br/>

## Introduction

Audit issue tracking is as important phase of audit cycle as audit fieldwork. The bigger the company is more messy and oblivious people can be if there is absent of well-formed issue tracking system.
Auditor’s roles involve keeping an issues log, due dates, and responsible employees. It is also auditor’s job to send reminder emails to directors and relevant employees until they complete issue and close action plan.\
\
However, numerous numbers of departments and employees make the process harder. Sometimes auditor may forget to include relevant employees in emails, omit some action plans in reminders, or simply make typos in business communications. Not only that, creating each report for each department takes significant amount of time. \
\
This project is to create a script in Python that can automate such processes in order to reduce human errors, save time, and improve work process more efficiently.

<br/>

## Goal
### Generate monthly <ins>Audit Issue Tracking reminder</ins> in <ins>word document</ins> with <ins>tables</ins> of Open Action Plans.

### Objects:
1. Perform data cleaning of raw data from Audit Data Repository
2. Filter data and display open action items for each department
3. Configure word document formats: such as font, style, color and margin

### Features: 
* Program: **Python**
* Packages: **pandas**, **Python-docx**, **datetime**
* Files:
    * [**Automation.ipynb**](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Automation.ipynb) --> Python Script
    * [**tblReport Query.xlsx**](https://github.com/tedgt97/Projct1.Python-docx/blob/main/tblReport%20Query.xlsx) --> raw data file exported from Data Repository

### Result:
![Information Protection Department Report Sample](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Pictures/Report_Result_Sample.png)

<br/>

## Data Validation
#### tblReport Query.xlsx 

! For company privacy policy, every record in the data is made up and does not reflect the real-life information

Fields:
* Project ID (Format YYYY-###): Identification for each project
* Project Name: Audit project name
* Issue Number: Issue numbering (Group)
* DepartmentResponsible: Department that is audited
* Action Status: Status of action plan. In original repository, it can be "Open"/"Closed"/"Pending Verification", but when exported, get only "Open"/"Pending Verification"
* Target Date: Initial target date to complete action plan
* Revised Target Date: 1st & Final revised target date to complete action plan. If an action is extended at least one time, always refer this field as a final date
* 1st Revised Target Date: Date when extended two times
* 2nd Revised Target Date: Date when extended three times
* 3rd Revised Target Date: Date when extended four times
* Management Action: Action Detail in Issue Number (Member)



## 1. CONFIG THIS BEFORE RUN
```
Save_Dir = r"#"
Date1 = '01.02.23'
Date2 = 'January 2, 2023'
```

* "Save_Dir": Folder directory where you want to save Monthly Audit Issue Tracker
* "Date1" & "Date2": Date variables that will be shown in body paragraph

## 2. Data Preparation
```
Data_Repository = Data_Repository.rename(columns = {'Issue Number': 'Issue # Ref', 'Project Name': 'Audit Name', 'Management Action': 'Brief Description'})
```
* Select only certain columns that will be shown in the report and rename into a report convention

<br/>

```
Data_Repository['Target Date'] = pd.to_datetime(Data_Repository['Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['Revised Target Date'] = pd.to_datetime(Data_Repository['Revised Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['1st Revised Target Date'] = pd.to_datetime(Data_Repository['1st Revised Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['2nd Revised Target Date'] = pd.to_datetime(Data_Repository['2nd Revised Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['3rd Revised Target Date'] = pd.to_datetime(Data_Repository['3rd Revised Target Date']).dt.strftime('%m.%d.%y')
```
* datetime.strftime --> Changing [datetime64] into [date object]. Not necessary for ordinary situation, but needed for *Expired Action Highlight* later

<br/>

```
today_o = date.today().strftime('%m.%d.%y')
today = datetime.strptime(today_o, "%m.%d.%y")
```
* Getting today date when generating the report and changing the format into [date object] including HH:MM:SS. Again, not necessary for ordinary situation, but needed for *Expired Action Highlight* later

<br/>

```
#Revised Date chronic order
for i in range(0, len(Data_Repository)):
    if pd.isnull(Data_Repository.iloc[i]['3rd Revised Target Date']) == False:
        list = [Data_Repository.iloc[i]['1st Revised Target Date'], Data_Repository.iloc[i]['2nd Revised Target Date'], Data_Repository.iloc[i]['3rd Revised Target Date']]
        list.sort(key = lambda date: datetime.strptime(date, '%m.%d.%y'))
        Data_Repository.loc[i, '1st Revised Target Date'] = Data_Repository.loc[i, '1st Revised Target Date'] = list[0]
        Data_Repository.loc[i, '2nd Revised Target Date'] = Data_Repository.loc[i, '2nd Revised Target Date'] = list[1]
        Data_Repository.loc[i, '3rd Revised Target Date'] = Data_Repository.loc[i, '3rd Revised Target Date'] = list[2]
    elif pd.isnull(Data_Repository.iloc[i]['2nd Revised Target Date']) == False:
        list = [Data_Repository.iloc[i]['1st Revised Target Date'], Data_Repository.iloc[i]['2nd Revised Target Date']]
        list.sort(key = lambda date: datetime.strptime(date, '%m.%d.%y'))
        Data_Repository.loc[i, '1st Revised Target Date'] = Data_Repository.loc[i, '1st Revised Target Date'] = list[0]
        Data_Repository.loc[i, '2nd Revised Target Date'] = Data_Repository.loc[i, '2nd Revised Target Date'] = list[1]
```
* 