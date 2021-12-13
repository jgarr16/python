#!/usr/bin/env python

# I got this one from a psot by Andrew Jabbitt
# works with Office 365
# https://ajabbitt.medium.com/automating-outlook-calendar-downloads-with-python-5fb3671b56a

import win32com.client, datetime
from datetime import date
from dateutil.parser import *
import calendar
import pandas as pd

# Step 1, Block 1 : access Outlook and get events from the calendar
Outlook = win32com.client.Dispatch("Outlook.Application")
ns = Outlook.GetNamespace("MAPI")
appts = ns.GetDefaultFolder(9).Items

# Step 1, Block 2 : sort events by occurrence and include recurring events
appts.Sort("[Start]")
appts.IncludeRecurrences = "True"


# Step 2, Block 1 : filter to the range : from = (today), to = (today + 1)
begin = date.today().strftime("%m/%d/%Y")
end = date.today() + datetime.timedelta(days=8)
end = end.strftime("%m/%d/%Y")
appts = appts.Restrict("[Start] >= '" +begin+ "' AND [END] <= '" +end+ "'")

# Step 3, Block 1 : create a list of excluded meeting subjects
excludedSubjects=('Exercise',)

# Step 3, Block 2 : populate dictionary of meetings
apptDict = {}
item = 0
for indx, a in enumerate(appts):
    subject = str(a.Subject)
    if subject in (excludedSubjects):
        continue
    else:
        organizer = str(a.Organizer)
        meetingDate = str(a.Start)
        date = parse(meetingDate).date()
        subject = str(a.Subject)
        duration = str(a.Duration)
        apptDict[item] = {"Duration":duration, "Organizer":organizer, "Subject":subject, "Date":date.strftime("%m/%d/%Y")}
        item = item + 1

# Step 4, Block 1 : convert discretionary to datafram and group by Date
aptDf = pd.DataFrame.from_dict(apptDict, orient='index', columns = ['Date','Subject','Duration','Organizer'])
aptDf = aptDf.set_index('Date')
aptDf['Meetings'] = aptDf[['Subject']].agg(' ||| '.join, axis=1)
grouped_aptDf = aptDf.groupby('Date').agg({'Meetings':', '.join})
grouped_aptDf.index = pd.to_datetime(grouped_aptDf.index)
grouped_aptDf.sort_index()

# Step 5, Block 1 : add timestamp to filename and save
filename = date.today().strftime("%Y%m%d") + '_meeting_list.csv'
grouped_aptDf.to_csv(filename, index=True, header=True)


print(len(appts))
