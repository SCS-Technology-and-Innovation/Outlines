def passing(grade):
    if 'W' in grade: # withdrawn
        return False
    return 'A' in grade or 'B' in grade

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
    "CMS2 627" : "Business Intelligence and Analytics",
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
    'CMIS 549' : 'Digital Media and Search Engine Optimization',
    'CMR2 650' : 'Digital Marketing Management',
    'CMR2 542' : 'Marketing Principles and Applications', 
    'CMR2 542' : 'Marketing Principles and Applications',
    'CMR2 543' : 'Marketing of Services', 
    'CMR2 643' : 'Marketing of Services', 
    'CMR2 548' : 'Processes of Marketing Research',  
    'CMR2 648' : 'Marketing Research and Reporting', 
    'CPL2 510' : 'Communication and Networking Skills', 
    'CPL2 610' : 'Advanced Communication and Presentation Skills',  
    'CGM2 520' : 'Sales Management and Negotiation Skills', 
    'CGM2 520' : 'Sales Management and Negotiation Skills',
    'CMR2 564' : 'Marketing Communications: A Strategic Approach',
    'CMR2 644' : 'Integrated Marketing Communications', 
    'CMR2 556' : 'Buyer Behaviour', 
    'CMR2 668' : 'Buyer Behaviour',
    'CGM2 520' : 'Sales Management and Negotiation Strategies',
    'CMR2 643' : 'Marketing of Services', 
    'CPL2 610' : 'Advanced Communication and Presentation Skills',
    'CPRL 644' : 'Integrated Digital Communications',
    'CGM2 520' : 'Sales Management and Negotiation Strategies',
    'CMR2 643' : 'Marketing of Services', 
    'CMR2 648' : 'Marketing Research and Reporting', 
    'CPRL 650' : 'Digital Marketing Management', 
    'CPL2 610' : 'Advanced Communication and Presentation Skills',
    'SCOM': 'Diploma in Supply Chain and Operations Management',
    'DSCOM' : 'Diploma in Supply Chain and Operations Management',
    'DSN' :'Graduate Certificate in Dynamic Supply Networks',
    'ISN': 'Graduate Certificate in Integrated Supply Networks',
    'DDABI': 'Diploma Digital Analytics and Business Intelligence',
    'CMS2 500' : 'Mathematics for Management',
    'CGM2 510' : 'Project Management: Tools and Techniques',
    'CGM2 610' : 'Project Management: Tools and Techniques',
    'CSNM 620' : 'Dynamic Supply Networks Data Analytics',
    'CMS2 515' : 'Operations Management',
    'CMS2 524' : 'Management of Service Operations',
    'CSNM 6??' : 'Management of Service Operations', # code pending
    'CMS2 525' : 'Supply Chain Management',
    'CSNM 610' : 'Principles of Dynamic Supply Networks',
    'CMS2 531' : 'Re-Engineering and Integration of Business Systems',
    'CSNM 6??' : 'eBusiness and eLogistics', # code pending
    'CMS2 532' : 'Lean Operations Systems',
    'CSNM 632' : 'Dynamic Supply Networks and Lean Operations Systems',
    'CMS2 540' : 'Six Sigma Quality Management',
    'CSNM 6??' : 'Six-Sigma and Supply Networks', # code pending
    'CMS2 550' : 'Supply Chain Field Project',
    'CSNM 6??' : 'Integrated Supply Networks Field Project',
    'CSNM 608' : 'Dynamic Supply Networks Sustainability',
    'CSNM 6??' : 'Integrated Production and Operations Management',
    'CNSM 612' : 'Dynamic Supply Networks Sourcing and Purchasing',
    'CSNM 6??' : 'Global Supply Management and International Logistics',
    'CSNM 605' : 'Dynamic Supply Networks Transformation',
    'CSNM 6??' : 'ESG Focus on Integrated Supply Networks'
}

APC = { 
    'DDABI': 'Prof.\\ Elisa Schaeffer',
    'DSCOM': 'Mr.\\ John Gradek'
} 

required = { 'DDABI': ['CCS2 505', 'CMIS 530', 'CMIS 543', 'CMIS 544',
                       'CMIS 545', 'CMIS 549', 'CMIS 550', 'CMS2 505',
                       'CMS2 527', 'CMS2 529'] }

suggestions = {
    "CCS2 505" : { "CCCS 610" }, # T&I equiv
    "CMIS 545" : { "CCCS 680" }, # close enough
    "CMIS 550" : { "CCCS 680" }, # revision
    "CMS2 505" : { "CCCS 640", 'CSNM 620' }, # T&I equiv + chart from john and dawne    
    "CMS2 527" : { "CMS2 627" }, # ME kept it, code change
    "CMS2 529" : { "CCCS 650" }, # revision, mandatory C2
    'CGM2 510' : { 'CGM2 610' }, # chart from john and dawne
    'CMIS 549' : { 'CMR2 650' }, # chart equiv marketing (confirmed with nicolette, actual equivalency)
    'CMR2 542' : { 'CMR2 642' }, # chart equiv marketing
    'CMR2 543' : { 'CMR2 643' }, # chart equiv marketing (confirmed with nicolette; customer service, no reqs)
    'CMR2 548' : { 'CMR2 648' }, # chart equiv marketing
    'CMR2 556' : { 'CMR2 668' }, # chart equiv marketing
    'CMIS 544' : { 'CMR2 650', 'CCCS 670' }, # nicolette digital marketing management
    'CMIS 530' : { 'CMR2 648' }, # nicolette suggests this
    'CMIS 549' : { 'CMR2 650', 'CCCS 670' }, # nicolete says it is a good match here plus our option
    'CMR2 564' : { 'CMR2 644' }, # chart equiv marketing
    'CMR2 566' : { 'CGM2 520', 'CMR2 643', 'CPL2 610', 'CPRL 644' }, # chart equiv marketing
    'CMR2 570' : { 'CMR2 691' }, # chart equiv marketing
    'CMS2 500' : { 'CMS2 500' }, # remains
    'CMS2 515' : { 'CCCS 640' }, # chart from john and dawne
    'CMS2 524' : { 'CSNM 6??' },  # chart from john and dawne
    'CMS2 525' : { 'CSNM 610' }, # chart from john and dawne
    'CMS2 531' : { 'CSNM 6??' }, # chart from john and dawne
    'CMS2 532' : { 'CSNM 632' }, # chart from john and dawne
    'CMS2 540' : { 'CSNM 6??' }, # chart from john and dawne (two courses with pending codes)
    'CPL2 510' : { 'CPL2 610' } # chart equiv marketing
}

schedule = {
    "CCCS 610" : "F23", # T&I teaching plan
    "CCCS 620" : "F23", # T&I teaching plan
    "CCCS 630" : "F23", # T&I teaching plan
    "CCCS 640" : "W24", # T&I teaching plan
    "CCCS 650" : "W24", # T&I teaching plan
    "CCCS 660" : "W24", # T&I teaching plan
    "CCCS 670" : "S24", # T&I teaching plan
    "CCCS 680" : "S24", # T&I teaching plan
    "CCCS 690" : "S24", # T&I teaching plan
    "CMIS 530" : "S23", # T&I teaching plan
    "CMIS 549" : "F23", # T&I teaching plan
    'CMIS 543': 'F23', # Suggestion from Elisa
    'CMS2 627' : 'F23' # asked John
}

completions = {
    'DACS' : ({ 'CCCS 610', 'CCCS 620', 'CCCS 630' }, {'CCCS', 'CMIS', 'CMS2', 'CCS2', 'CMR2' }),
    'DDDM' : ({ 'CCCS 640', 'CCCS 650', 'CCCS 660' }, {'CCCS', 'CMIS', 'CMS2', 'CCS2' }),
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


def match(course, patterns):
    if course is None:
        return False
    for p in patterns:
        if p in course:
            return True
    return False

from collections import defaultdict
needs = defaultdict(set)

def printout(label, name, passed, email, dipl):
    content = template.replace('!!NAME!!', name)
    content = content.replace('!!SID!!', f'McGill Student ID {label}')
    content = content.replace('!!EMAIL!!', email)
    content = content.replace('!!DIPL!!', names[dipl])
    content = content.replace('!!APC!!', APC[dipl])
    # when to take the missing courses
    #print(passed)
    status = { course : course in passed for course in required[dipl] }
    total = len(status)
    completed = sum(status.values())
    missing = total - completed
    if missing == 10:
        print(label, 'has not completed any courses')
        return
    if missing == 0: # ready to graduate, no need for a studyplan
        print(label, 'has completed the diploma')
        return
    else:
        pl = 's' if missing > 1 else ''
        mis = str(missing) if missing < 10 else 'any'
        listing = f'You have not yet completed {mis} course{pl} from the diploma.\n\\begin{{enumerate}}[noitemsep,topsep=0pt]\n'
        available = set()
        #print(status)
        for (course, ok) in status.items():
            when = ''
            if ok: # already passed
                available.add(course)
            else: # not yet passed
                needs[course].add(label)
                when = schedule.get(course, None)
                if when is None:
                    subs = suggestions.get(course, set())
                    opt = ''
                    for sub in subs:
                        when = schedule.get(sub, None)
                        if when is not None:
                            term = full(when)                    
                            opt += f'\item {sub} {names[sub]} which will be available for registration in {term}\n'
                    if opt != '':
                        listing += f'\\item {course} {{\\em {names[course]}}} can be substituted by\n\\begin{{itemize}}[noitemsep,topsep=0pt]\n' + opt + '\\end{itemize}\n'
                        if '680' in opt:
                            print('S24', email)
                    else:
                        error = f'ERROR: nothing scheduled for {course} {label} {email}'
                        print(error)
                        listing += f'\\item \\textcolor{{red}}{{{error}}}\n'
                else:
                    term = full(when)                                    
                    listing += f'\\item {course} {{\\em {names[course]}}} \\\\ \\phantom{{indent}} will be available for registration in {term}\n'            
        listing += '\\end{enumerate}'
        prefix = '''\\newpage

\\section*{Course availability}

According to our current records, there are required courses in the
diploma which you have not yet completed.
\\textcolor{red}{check until when they can still get the diploma}

'''
        listing = prefix + listing
        
        content = content.replace('!!LIST!!', listing)
    listing = ''
    if completed > 0:
        pl = 's' if completed > 1 else ''        
        listing = f'You have already completed {completed} course{pl}.\n\\begin{{itemize}}[noitemsep,topsep=0pt]\n'
        spent = set()
        mentioned = set()
        # if they were to transfer
        for cert in completions:
            sublist = ''
            (mandatory, complementary) = completions[cert]
            for course in available:
                if course in spent:
                    continue
                alt = suggestions.get(course, None)
                if alt in mandatory:
                    sublist += f'\\item {course} {{\\em {names[course]}}} can substituted the {{\\bf mandatory}} course {alt} {names[alt]}\n'
                    spent.add(course)
                elif match(alt, complementary) or match(course, complementary):
                    sublist += f'\\item {course} {{\\em {names[course]}}} could be used as a complementary course\n'
                    mentioned.add(course)
            if len(sublist) > 0:
                note = ''
                if mandatory.issubset(spent) and len(mentioned) >= 2:
                    note = '(could be completed)\\\\'
                listing += f'\\item{{\\bf Graduate Certificate in {{\\em {names[cert]}}}}}\n{note}\n\\begin{{itemize}}[noitemsep,topsep=0pt]\n{sublist}\n\\end{{itemize}}\n'
        for course in available - (spent | mentioned):
            listing += f'\\item \\textcolor{{red}}{{ERROR: nothing defined for {course} yet}}\n'
        listing += '\\end{itemize}\n'
        prefix = '''\\newpage

\\section*{Possible program transfers}

If you do not wish to complete the diploma, we recommend you to
consider the option of transferring to the newly created graduate
certificates as an option to completing the diploma.

Please note that whereas a single course could serve as a
complementary course in more than one certificate, you are only
allowed to share one course between two
certificates. \\textcolor{red}{revise the wording with Sue}

'''
        listing = prefix + listing
    content = content.replace('!!CERT!!', listing)    
    with open(f'studyplan-{label}.tex', 'w') as output:
        print(content, file = output)

from sys import argv
debug = 'debug' in argv

files = {
    'DDABI': 'digital.xlsx'
}

import pandas as pd

records = dict()
emails = dict()
activity = defaultdict(set)

recent = [ 2022, 2023, 2024 ]

def active(termcodes):
    for term in termcodes:
        for tok in recent:
            if str(tok) in str(term):
                return True
    return False


for dataset in files:
    data = pd.ExcelFile(files[dataset])
    students  = data.parse(data.sheet_names[0]) 
    for index, row in students.iterrows():
        email = str(row[465])
        studentID = str(row[0])
        if studentID not in emails and email != 'nan':
            emails[studentID] = email
        if len(studentID) == 9:
            if studentID not in records:
                fullname = str(row[3]).split('/')
                if len(fullname) > 1:
                    lastName = fullname[0] 
                    firstName = fullname[1]
                    name = f'{firstName} {lastName}'
                else:
                    print(fullname)
                    continue
                names[studentID] = name
                records[studentID] = set()
            letterCode = row[5]
            numberCode = int(row[7])
            courseCode = f'{letterCode} {numberCode:03d}'
            finalGrade = str(row[35])
            activity[studentID].add(row[2]) # term code
            if passing(finalGrade):
                records[studentID].add(courseCode)
                # print(studentID, courseCode, finalGrade)
        else:
            if debug:
                print('ignoring', studentID)



actives = 0

for student in records:
    if active(activity[student]):
        printout(student, names[student], records[student], emails.get(student, ''), dataset)
        actives += 1
    else:
        print(student, 'inactive')

print(actives, 'active students')
for course in needs:
    print(len(needs[course]), 'students still need', course)
#    for student in needs[course]:
#        print(course, student)
