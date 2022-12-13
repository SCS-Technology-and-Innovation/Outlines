names = {
    "CMS2 500" : "Mathematics for Management",
    "CCS2 505" : "Programming for Data Science",
    "CMIS 530" : "Digital Analytics and Targeting",
    "CMIS 543" : "Digital Customer Experience",
    "CMIS 544" : "Digital Marketing Automation, Planning and Technology",
    "CMIS 545" : "Cloud Computing Architecture",
    "CMIS 549" : "Digital Media and Search Engine Optimization",
    "CMIS 550" : "Fundamentals of Big Data",
    "CMS2 505" : "Quantitative Analysis Tools in Decision Making",
    "CMS2 527" : "Business Intelligence and Analytics",
    "CMS2 529" : "Introduction to Data Analytics"
}

transfers = {
    "CCS2 505" : "CCCS 620", # revision, mandatory C1
    "CMIS 550" : "CCCS 680", # revision
    "CMIS 545" : "CCCS 680", # close enough
    "CMS2 529" : "CCCS 650", # revision, mandatory C2
    "CMS2 527" : "CMS2 627", # ME kept it, code change
    "CMS2 505" : "CCCS 640", # revision, mandatory C2
    "CCS2 505" : "CCCS 610" # close enough, mandatory C1
}

schedule = {
    "CCCS 610" : "F23",
    "CCCS 620" : "F23",
    "CCCS 630" : "F23",
    "CCCS 640" : "W24",
    "CCCS 650" : "W24",
    "CCCS 660" : "W24",
    "CCCS 670" : "S24",
    "CCCS 680" : "S24",
    "CCCS 690" : "S24",
    "CMIS 530" : "S23",
    "CMIS 543" : "W23",
    "CMIS 544" : "W23",
    "CMIS 549" : "F23"
}

from sys import argv
debug = 'debug' in argv

import pandas as pd
data = pd.ExcelFile('Students_pending_DiplomaDigAnalytics.xlsx')
students  = data.parse(data.sheet_names[0])
header = [h.strip() for h in students.columns.values.tolist()]
courses = header[5:16]

for index, student in students.iterrows():
    studentID = str(student[1])
    if len(studentID) == 9:
        lastName = student[2]
        firstName = student[3]
        reg = student[5:16]
        done = [ True if s > 0 else False for s in reg ]
        print(studentID, sum(done))
    else:
        if debug:
            print('ignoring', studentID)
