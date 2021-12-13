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


# Step 2, Block 1 : filter to the range : from = (today), to = (today + 24 hours)
begin = datetime.datetime(2021,12,13)
begin = begin - datetime.timedelta(0,180)
end = begin + datetime.timedelta(2,0)
end = end.strftime("%m/%d/%Y")
begin = begin.strftime("%m/%d/%Y")
print("begin: "+str(begin)+"\nend: "+str(end))
appts = appts.Restrict("[Start] >= '" +str(begin)+ "' AND [END] <= '" +str(end)+ "'")

# Step 3, Block 1 : create a list of excluded meeting subjects
excludedSubjects=('Exercise','Project Tag','Weekly OCIO Enterprise Budget Tag-Up',)
swapSubjects={
    "7:45 a.m. Daily Code I Services Status Tag":"IT Ops",
    "Weekly EC Tag-up":"EC",
    "New CIO Staff Meeting Series":"CSM",
    "John & Bryan Tag Up":"Bryan",
    "ACMC":"ACMC",
    "Code I Sr. Mgmt.":"Sr. Ldrs.",
    "OCIO MAP Implementation Weekly Meeting":"MAP",
    "Code IO Tag-up":"Code IO",
    "OCIO Future of Work Subteam":"OCIO FoW",
    "Code I Chiefs - Tag-up":"Chiefs",
    "Code I ITPMB Meeting":"ITPMB",
    "Mission Enabling Bi-weekly Tag":"ME Tag"
    }

# Step 3, Block 2 : populate dictionary of meetings
apptDict = {}
item = 0
for indx, a in enumerate(appts):
    subject = str(a.Subject)
    if subject.startswith(excludedSubjects):
        continue
    elif subject in (swapSubjects):
        organizer = str(a.Organizer)
        meetingDate = str(a.Start)
        date = parse(meetingDate).date()
        subject = swapSubjects[subject]  # change the subject to a preferred, condensed version
        duration = str(a.Duration)
        apptDict[item] = {"Duration":duration, "Organizer":organizer, "Subject":subject, "Date":date.strftime("%m/%d/%Y"), "obsidianDate":date.strftime("%Y-%m-%dT%H:%M")}
        item = item + 1
        # continue
    else:
        organizer = str(a.Organizer)
        meetingDate = str(a.Start)
        date = parse(meetingDate).date()
        subject = str(a.Subject)
        duration = str(a.Duration)
        apptDict[item] = {"Duration":duration, "Organizer":organizer, "Subject":subject, "Date":date.strftime("%m/%d/%Y"), "obsidianDate":date.strftime("%Y-%m-%dT%H:%M")}
        item = item + 1

# Step 4, Block 1 : convert discretionary to datafram and group by Date
aptDf = pd.DataFrame.from_dict(apptDict, orient='index', columns = ['Date','obsidianDate','Subject'])
aptDf = aptDf.set_index('Date')
aptDf['Meetings'] = aptDf[['Subject']].agg(' ||| '.join, axis=1)
grouped_aptDf = aptDf.groupby('Date').agg({'Meetings':', '.join})
grouped_aptDf.index = pd.to_datetime(grouped_aptDf.index)
grouped_aptDf.sort_index()

# Step 5, Block 1 : add timestamp to filename and save
filename = date.today().strftime("%Y%m%d") + '_meeting_list.csv'
grouped_aptDf.to_csv(filename, index=True, header=True)

# print(len(appts)) -- for status checking...
