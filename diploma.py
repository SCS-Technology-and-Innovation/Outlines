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
    "CMS2 529" : "Introduction to Data Analytics",
    'DACS' : 'Data Analysis for Complex Systems',
    'DDDM' : 'Data-Driven Decision Making',
    'CCCS 610' : 'Digital Thinking for Data Analysis',
    'CCCS 620' : 'Data Analysis and Modeling',
    'CCCS 630' : 'Complex Systems',
    'CCCS 640' : 'Applied Decision Science',
    'CCCS 650' : 'Applied Data Science',
    'CCCS 660' : 'Computational Intelligence',
    'CCCS 670' : 'Information Visualization',
    'CCCS 680' : 'Scalable Data Analysis',
    'CCCS 690' : 'Applied Computational Research',
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

season = {
    'F' : 'Fall',
    'W' : 'Winter',
    'S' : 'Summer'
}

def full(st):
    year = f'20{st[-2:]}'
    term = season[st[0]]
    return f'{term} {year}'
        
# load the template
template = None
with open('studyplan.tex') as source:
    template = source.read()

completions = {
    'DACS' : ({ 'CCCS 610', 'CCCS 620', 'CCCS 630' }, {'CCCS', 'CMIS', 'CMS2', 'CCS2' }),
    'DDDM' : ({ 'CCCS 640', 'CCCS 650', 'CCCS 660' }, {'CCCS', 'CMIS', 'CMS2', 'CCS2' }),
}

def match(course, patterns):
    if course is None:
        return False
    for p in patterns:
        if p in course:
            return True
    return False
                
def printout(label, name, status):
    content = template.replace('!!NAME!!', name)
    # when to take the missing courses
    completed = sum(status.values())
    missing = len(status) - completed
    listing = f'You have not yet registered to {missing} courses from the diploma.\n\\begin{{itemize}}\n'
    available = set()
    for (course, reg) in status.items():
        when = ''
        note = ''
        if reg:
            available.add(course)
        else:
            when = schedule.get(course, None)
            if when is None:
                sub = transfers.get(course, None)
                if sub is not None:
                    note = 'can be substituted by course {sub} {names[sub]} which'
                    when = schedule.get(sub, None)
            if when is not None:
                term = full(when)
                listing += f'\\item {course} {{\\em {names[course]}}} will be available for registration in {term}\n'
            else:
                listing += f'\\item \\textcolor{{red}}{{ERROR: nothing defined for {course} yet}}\n'
            
    listing += '\\end{itemize}'
    content = content.replace('!!LIST!!', listing)
    if completed > 0:
        listing = f'You have already registered for {completed} courses.\n\\begin{{itemize}}'
        spent = set()
        mentioned = set()
        # if they were to transfer
        for cert in completions:
            sublist = ''
            (mandatory, complementary) = completions[cert]
            for course in available:
                if course in spent:
                    continue
                alt = transfers.get(course, None)
                if alt in mandatory:
                    sublist += f'\\item {course} {{\\em {names[course]}}} can substitute the mandatory course {alt} {names[alt]}'
                    spent.add(course)
                elif match(alt, complementary) or match(course, complementary):
                    sublist += f'\\item {course} {{\\em {names[course]}}} could be used as a complementary course'
                    mentioned.add(course)
            if len(sublist) > 0:
                listing += f'\\item  {{ \\bf Graduate Certificate in {{\\em {names[cert]}}}}} \\begin{{itemize}}{sublist} \\end{{itemize}}'
        for course in available - (spent | mentioned):
            listing += f'\\item \\textcolor{{red}}{{ERROR: nothing defined for {course} yet}}\n'
        listing += '\\end{itemize}'        
        content = content.replace('!!CERT!!', listing)    
    with open(f'studyplan-{label}.tex', 'w') as output:
        print(content, file = output)

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
        name = f'{firstName} {lastName}'
        reg = student[5:16]
        done = { course : count > 0 for (course, count) in zip(courses, reg) }
        printout(studentID, name, done)
    else:
        if debug:
            print('ignoring', studentID)
