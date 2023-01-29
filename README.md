# Projct1: Python-docx 
#### **Monthly Audit Issue Tracker by using Python-docx package**
\
**Table of Contents**
1. [Introduction](#introduction)
2. [Goal](#goal)
3. [Data Validation](#data-validation)
4. [Script Details](#script-details)\
    4.1 [CONFIG THIS BEFORE RUN](#1-config-this-before-run)\
    4.2 [Data Preparation](#2-data-preparation)\
    4.4 [Body](#44-body)


<br/>

## Introduction

Audit issue tracking is as important phase of audit cycle as audit fieldwork. Auditor’s roles involve constant tracking of issues logs, due dates, and action plans. It is also auditor’s job to send reminder emails to directors and responsible employees until they complete issue and close action plan.\
\
However, numerous numbers of issues and action plans make the process harder. Sometimes auditor may forget to copy relevant employees in emails, omit issues to include in reminder, overlook aged action plans, or simply make typos in business communications. Not only that, creating each report for each department takes significant amount of time. \
\
This project is to create a script in Python that can automate such processes in order to reduce human errors, save time, and improve work process more efficiently.

<br/>

## Goal
### Generate monthly <ins>Audit Issue Tracking Report</ins> in <ins>word document</ins> with <ins>tables</ins> for Open Action Plans.

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
#### [**tblReport Query.xlsx**](https://github.com/tedgt97/Projct1.Python-docx/blob/main/tblReport%20Query.xlsx)

**! Every record in the data is made up and does not reflect the real-life information**

Fields:
* Project ID (Format YYYY-###): Identification for each project
* Project Name: Audit project name
* Issue Number: Issue numbering (Group)
* DepartmentResponsible: Name of department audited
* Action Status: Status of action plan. In original repository, it can be "Open"/"Closed"/"Pending Verification", but current exported data is alreday filtered by "Open" & "Pending Verification"
* Target Date: Initial target date to complete action plan
* Revised Target Date: 1st or Final revised target date to complete action plan. If an action is extended at least one time, always refer this field as a final date
* 1st Revised Target Date: Target date when extended two times
* 2nd Revised Target Date: Target Date when extended three times
* 3rd Revised Target Date: Target Date when extended four times
* Management Action: Action Detail in Issue Number (Member)


<details>
<summary>

## Script Details

### 4.1 CONFIG THIS BEFORE RUN

</summary>

```
Save_Dir = r"#"
Date1 = '01.02.23'
Date2 = 'January 2, 2023'
```

* "Save_Dir": Folder directory where you want to save Monthly Audit Issue Tracker
* "Date1" & "Date2": Date variables that will be shown in body paragraph

</details>

<details>
<summary>

### 4.2 Data Preparation

</summary>

```
Data_Repository = Data_Repository.rename(columns = {'Issue Number': 'Issue # Ref', 'Project Name': 'Audit Name', 'Management Action': 'Brief Description'})
```
* Rename columns that will be shown in the report into certain convention

<br/>

```
Data_Repository['Target Date'] = pd.to_datetime(Data_Repository['Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['Revised Target Date'] = pd.to_datetime(Data_Repository['Revised Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['1st Revised Target Date'] = pd.to_datetime(Data_Repository['1st Revised Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['2nd Revised Target Date'] = pd.to_datetime(Data_Repository['2nd Revised Target Date']).dt.strftime('%m.%d.%y')
Data_Repository['3rd Revised Target Date'] = pd.to_datetime(Data_Repository['3rd Revised Target Date']).dt.strftime('%m.%d.%y')
```
* datetime.strftime --> Changing [datetime64] into [date object]. Needed for *Expired Action Highlight* later

<br/>

```
today_o = date.today().strftime('%m.%d.%y')
today = datetime.strptime(today_o, "%m.%d.%y")
```
* Getting today date when generating the report and changing the format into [date object] including HH:MM:SS. Needed for *Expired Action Highlight* later

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
![Data Error Sample](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Pictures/Data_Error.PNG)

* Notice that "Payment Processing" project has error in Target Date; 1st Revised Target Date comes after 2nd Revised Target Date
    * This is due to human error when entering details in the data repository. 
* Since "Target Date" and "Revised Target Date" are not influenced by this error, simply re-arrange 1st & 2nd & 3rd Revised Target Date in chronic order by using list.sort

</details>


### 4.4 Body


 Every code for **4. Body** is inside of function called ***body***. Thus, be mindful of indentation.

 ```
 def body(To, cc, dept, Date1, Date2, evidence):
    doc = docx.Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.62)
    section.bottom_margin = Inches(0.31)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.81)
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(31, 73, 125)
```
* This is a preset of word document format
* ***doc = docx.Document()*** is Document constructor from **Python-docx package**
    * Every Document objects must follow after the initial constructor
* ***section*** codes configure fortmat of [Layout --> Margins] in Document
    ![Doc Margins Configuration](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Pictures/doc_margins.PNG)

* ***normal_style*** codes configure format of [Home --> Styles --> Normal] in Document
    ![Doc Style Configuration](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Pictures/doc_style.PNG)

> Note that ***body*** function has six different arguments\
    * **To** & **cc** & **evidence** --> defined in <ins>dictionary</ins> from **5. Departments** section\
    * **dept** --> defined in <ins>list</ins> from **5. Departments** section\
    * **Date1** & **Date2** --> already defined in **1. Config This Before Run** Section

<br/>

```
    main1 = '''
To: {}

cc: {}
    '''.format(To, cc)
```
![Email Receivers](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Pictures/main1.PNG)

* ***main1*** prints names of employees who will receive report email
    * "To:" for head of department
    * "cc:" for relevant employees

<br/>

```
    line1 = '''
Subject: {} Open Audit Issues / Action Plans Summary as of {}
    '''.format(dept, Date1)
```

![Subject line](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Pictures/line1.PNG)

* ***line1*** prints subject line of email
    * contains the name of department and date of report

<br/>

```
    line2 = '''

Please find the Outstanding Audit Issues/Action Plans Summary as of {} (attached).

When actions have been completed, please provide the supporting evidence to close the action. Email evidence to {}.

Should you need to revise the target completion date, please send an email to Internal Audit DH, or designee, noting the revised date, reason for the delay and interim action to mitigate risk, as applicable.

Please note: Target Date extensions will need to be approved by the respective Division Head, via email.

    '''.format(Date2, evidence)
```
![line2](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Pictures/line2.PNG)


* ***line2*** prints body paragraph of email
    * "**evidence**" for whom to send evidence of action plan. Mostly auditor who is in charge of (in this case, Ted Jung)

<br/>

```
    parag = doc.add_paragraph(main1, 'Normal')
    parag.add_run(line1).font.color.rgb = RGBColor(0, 32, 96)
    parag.add_run(line2)
```
* In order to add texts in Document, initial paragraph must be created by ***parag***
    * multiple paragraphs can exist as different groups
* After initial texts of paragraph, additional texts can be added by ***add_run***
    * each line added is subordinated to paragraph and follows paragraph's format unless defined seperately like ***line1***

<br/>

```
    excel = Data_Repository[Data_Repository['DepartmentResponsible'] == dept]
```
* By filtering dataframe with department list, readily generate reports for necessary departments only

<br/>

```
    if excel['3rd Revised Target Date'].notnull().any() == True:
        excel = excel[['Issue # Ref', 'Target Date', '1st Revised Target Date', '2nd Revised Target Date', '3rd Revised Target Date', 'Revised Target Date', 'Audit Name', 'Brief Description']]
    elif excel['2nd Revised Target Date'].notnull().any() == True:
        excel = excel[['Issue # Ref', 'Target Date', '1st Revised Target Date', '2nd Revised Target Date', 'Revised Target Date', 'Audit Name', 'Brief Description']]
    elif excel['1st Revised Target Date'].notnull().any() == True:
        excel = excel[['Issue # Ref', 'Target Date', '1st Revised Target Date', 'Revised Target Date', 'Audit Name', 'Brief Description']]
    elif excel['Revised Target Date'].notnull().any() == True:
        excel = excel[['Issue # Ref', 'Target Date', 'Revised Target Date', 'Audit Name', 'Brief Description']]
    else:
        excel = excel[['Issue # Ref', 'Target Date', 'Audit Name', 'Brief Description']]
```
* Default format of the table in the report requires only four columns as shown in ***else:*** code
* However, whenever actions get extended their Target Dates, extra columns should be added in chronic order
    * If extended one time --> **Revised Target Date**
    * If extended two times --> **Revised Target Date** and **1st Revised Target Date**
    * If extended three times --> **Revised Target Date** and **1st Revised Target Date** and **2nd Revised Target Date**
    * If extended four times --> **Revised Target Date** and **1st Revised Target Date** and **2nd Revised Target Date** and **3rd Revised Target Date**
* Number of columns will vary on the maximum number of Target Date extended
    * Each department has different number of extension and therefore it is necessary to set columns differently by each circumstance
> For example:\
[<ins>Commercial Credit_01.02.23.docx</ins>](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Report%20Result/Commercial%20Credit_01.02.23.docx) does not have any action extended its Target Date and therefore has only 4 columns.\
[<ins>Servicing & Loyalty_01.02.23.docx</ins>](https://github.com/tedgt97/Projct1.Python-docx/blob/main/Report%20Result/Servicing%20%26%20Loyalty_01.02.23.docx) has one action extended four times, and therefore has 8 columns