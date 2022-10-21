# pip install pandas
# pip install openpyxl
import pandas as pd
from math import fabs
from datetime import datetime
from os.path import exists

SHAREPOINT = '15/08/2022 23:55:00'
CUTOFF = datetime.strptime(SHAREPOINT, '%d/%m/%Y %H:%M:%S')

THRESHOLD = 0.02
DEFAULT = 'This course consists of a community of learners of which you are an integral member; your active participation is therefore essential to its success. This means: attending class; visiting myCourses, doing the assigned readings/exercises before class; and engaging in class discussions/activities.'

wrong = set()

def ascii(text):
    if '&' in text and '\\&' not in text:
        text = text.replace('&', '\\&') # LaTeX not accounted for
    if '%' in text and '\\%' not in text:
        text = text.replace('%', '\\%')
    text = text.replace('$', '\\$') # not contemplating math mode
    text = text.replace('_', '') # no underscores are permitted here
    text = text.replace('#', '') # not needed, really
    text = ''.join(i for i in text if ord(i) < 128) # no unicode
    clean = []
    for line in text.split('\n'):
        words = []
        for word in line.split():
            if 'http' in word and '://' in word: # url
                if '\\url' not in word: # needs to be linked
                    # no parenthesis
                    word = word.replace('(', '')
                    word = word.replace(')', '')
                    word = f'\\url{{{word}}}'
            words.append(word)
        clean.append(' '.join(words))
    joint = '\n'.join(clean)
    joint = joint.replace('#s#', '\n\\begin{itemize}')
    joint = joint.replace('#i#', '\n\\item ')
    joint = joint.replace('#e#', '\n\\end{itemize}')
    return joint.replace('. ', '.\n\n') # paragraphs

def contact(text):
    clean = []
    text = ascii(text)
    text = text.replace('Full name:','')
    text = text.replace('E-mail address:','')
    text = text.replace('<', '')
    text = text.replace(',', '')
    text = text.replace('; ', '')    
    text = text.replace('>', '')
    text = text.replace('(tbc)', '')
    if '. ' in text: # LaTeX space cancelling
        text = text.replace('. ', '.\\ ') 
    for word in text.split():
        if '@' in word:
            link = f'\\href{{mailto:{word}}}{{{word}}}' 
            clean.append(link)
        else:
            clean.append(word)
    return ' '.join(clean)

def group(line):
    line = ascii(line)
    lines = []
    for entry in line.split('\n'):
        cols = entry.split(';')
        if len(cols) > 4: # format not respected, extra columns cut off
            # print('WAY TOO MANY COLUMNS', entry)
            cols = cols[:4]
        while len(cols) < 4:
            cols.append('') # empty missing columns
        lines.append(cols)
    return lines

# load the course information sheet
info = pd.read_csv('courses.csv')

# load the TA information
TAinfo = pd.ExcelFile('Teaching_Assistants.xlsx')
TAs  = TAinfo.parse(TAinfo.sheet_names[0])
TAh = [h.strip() for h in TAs.columns.values.tolist()]
tacl = TAh.index('Course')
tacn = TAh.index('Code')
tas = TAh.index('Section')
tat = TAh.index('Teaching/Course Assistant')
tan = TAh.index('Candidate')

assistant = dict()
for index, row in TAs.iterrows():
    name = row[tan]
    if not  isinstance(name, str):
        break
    name = name.lstrip().strip()
    code = row[tacl].strip()
    number = row[tacn]
    section = row[tas]
    kind = row[tat]
    if 'TA 120' in name:
        name = name.replace('TA 120', '')
    if 'low registration' in name:
        continue
    details = f'\item[Assistant ({kind})]{{{name}}}'
    number = '{:03d}'.format(number)
    print(code, number, section, details)
    assistant[f'{code} {number} {section}'] = details

allbymyself = '' # no TA, no CA (say nothing for now)
    
# load the template
with open('outline.tex') as source:
    template = source.read()
template = template.replace('!!TERM!!', 'Fall 2022') # update the term/year
    
# load the outline responses
responses = pd.ExcelFile('outline.xlsx')
forms  = responses.parse(responses.sheet_names[0])

completion= pd.to_datetime(forms.iloc[:, 2], format = '%m-%d%-y %H:%M:%S')
b = completion < CUTOFF
a = completion >= CUTOFF
# drop the first five columns
forms = forms.iloc[: , 5:]
# split into two for the people who do not read teams
before = forms.loc[b]
after = forms.loc[a]

print('Responses from Forms pre-sharepoint', before.shape)
print('Responses from Forms post-sharepoint', after.shape)
header = [h.strip() for h in forms.columns.values.tolist()]

HWMIN = 'Required minimum computer hardware specifications'
HWREC = 'Recommended computer hardware specifications'
SWMIN = 'Required software and services'
SWREC = 'Recommended software and services'
ADMIN = 'Administrative account'
IMIN = 'Required minimum internet connection specifications'
IREC = 'Recommeneded internet connection specifications'

for i in range(len(header)):
    current = header[i]
    if 'admin' in current.lower():
        header[i] = ADMIN
    elif 'hardware' in current:
        if 'minimum' in current:
            header[i] = HWMIN
        else:
            header[i] = HWREC
    elif 'software' in current:
        if 'required' in current.lower():
            header[i] = SWMIN
        else:
            header[i] = SWREC
    elif 'internet' in current:
        if 'minimum' in current:
            header[i] = IMIN
        else:
            header[i] = IREC

t = header.index('Course title')
n = header.index('Course number')
s = header.index('Section number')
i = header.index('Instructor(s)')
h = header.index('Office hours')
d = header.index('Course description')
o = header.index('Learning outcomes')
hwmin = header.index(HWMIN)
hwrec = header.index(HWREC)
swmin = header.index(SWMIN)
swrec = header.index(SWREC)
admin = header.index(ADMIN)
imin = header.index(IMIN)
irec = header.index(IREC)
req = header.index('Required readings')
opt = header.index('Additional optional course material')
m = header.index('Instructional methods')
a = header.index('% for Attendance and active participation')
e = header.index('Explanation for Attendance and active participation')
g = header.index('Other graded items')

# these have double quotes
edited = pd.read_csv('sharepoint.csv', header = [0], engine = 'python')

#for h1, h2 in zip(header, edited.columns):
#    print(h1, h2,'\n')

print('Responses from Sharepoint', edited.shape)
edited.columns = before.columns # same header
data = pd.concat([before, edited, after], axis = 0, ignore_index = True)
print('Combined responses', data.shape)
data.fillna('', inplace = True)

fields = dict()

for index, response in data.iterrows():
    error = ''
    code = response[n].strip() # course number
    if len(code) < 3:
        print('Skipping an empty row')
        continue
    section = str(response[s]).strip() # course section 
    if len(section) == 0 and '-' in code:
        parts = code.split('-')
        code = parts[0].strip()
        section = parts[1]
    if 'CRN' in section:
        # error = '\nExtra details provided in section number\n'
        if '(' in section:
            section = section.split('(').pop(0).strip()
    if 'Fall' in section:
        error = '\nActual section number is missing\n'
        print(code, 'lacks section number, omitting')
        continue
    if '-' in code: # someone wrote <YCIT 001 - 001>
        parts = code.split('-')
        section = parts[-1].strip()
        code = parts[0].strip()
    lettercode = None
    numbercode = None
    if ' ' in code:
        lettercode, numbercode = code.split(' ')
    else:
        lettercode = code[:3]
        numbercode = code[4:]
    outline = template.replace('!!CODE!!', code)
    outline = outline.replace('!!SECTION!!', section)
    print('Retrieving CA/TA for', code, section)
    outline = outline.replace('!!ASSISTANT!!', assistant.get(f'{code} {section}', allbymyself))
    code = code.replace(' ', '') # no spaces in the filename
    section = str(section)
    while len(section) < 3:
        section = '0' + section
    output = f'{code}-{section}.tex'
    if exists(output):
        print('Reprocessing', output)
    else:
        print('Processing', output)
    outline = outline.replace('!!NAME!!', ascii(response[t])) # course title
    outline = outline.replace('!!CODE!!', code)
    outline = outline.replace('!!SECTION!!', section)
    outline = outline.replace('!!INSTRUCTOR!!', contact(response[i].strip())) # instructor
    hours = response.get(h, '')
    if len(hours) == 0:
        hours = 'Upon request'
    outline = outline.replace('!!HOURS!!', hours.strip())
    details = None
    if lettercode is not None and numbercode is not None:
        print('Extracting details for', lettercode, numbercode)
        details = info.loc[(info['Code'] == lettercode) & (info['Number'] == int(numbercode))]
    else:
        error += '\nCourse code specification not found'
    if details is None or details.empty:
        error += '\nCourse details are not in the domain catalogue, corresponding fields will not be populated\n'
    else:
        graduate = lettercode[-1] == '2' or numbercode[0] == '6'
        if graduate:
            print(lettercode, numbercode, 'is a graduate course')
        prereq = details['Pre-requisites'].iloc[0]
        if pd.isna(prereq):
            outline = outline.replace('!!PREREQ!!', 'None')        
        else:
            outline = outline.replace('!!PREREQ!!', prereq)
        coreq = details['Co-requisites'].iloc[0]
        if pd.isna(coreq):
            outline = outline.replace('!!COREQ!!', 'None')
        else:
            outline = outline.replace('!!COREQ!!', coreq)
        amount = details['Credit amount'].iloc[0]
        kind = ''
        if details['Credit type'].iloc[0] == 'credits':
            if graduate:
                outline = outline.replace('!!GRADING!!', '\\input{graduate.tex}')
                kind = 'Graduate-level credit course'
            else:
                outline = outline.replace('!!GRADING!!', '\\input{undergraduate.tex}')
                kind = 'Undergraduate-level credit course'
            outline = outline.replace('!!CREDITS!!', f'{amount} credits')
            outline = outline.replace('!!TRAINING!!', '\\input{training.tex}')
            outline = outline.replace('!!FINAL!!', '\\input{minerva.tex}')
            outline = outline.replace('!!PROFILE!!', '\\input{profilem.tex}')            
        else:
            kind = 'Non-credit course'
            outline = outline.replace('!!CREDITS!!', f'{amount} CEUs')
            outline = outline.replace('!!GRADING!!', '\\input{noncredit.tex}')
            outline = outline.replace('!!TRAINING!!', '')
            outline = outline.replace('!!FINAL!!', '\\input{athena.tex}')
            outline = outline.replace('!!PROFILE!!', '\\input{profilea.tex}')                       
        outline = outline.replace('!!KIND!!', f'{kind}')
        h = details['Contact hours'].iloc[0]
        if pd.isna(h): # credit-side default is 39
            h = 39
        outline = outline.replace('!!CONTACT!!', f'{h} hours')
        h = details['Approximate assignment hours'].iloc[0]
        if pd.isna(h):
            outline = outline.replace('!!ASSIGNMENT!!','') # nothing goes here
        else:
            outline = outline.replace('!!ASSIGNMENT!!',
                                      f'\\item[Independent study hours]{{Approximately {h} hours }}')
    outline = outline.replace('!!DESCRIPTION!!', ascii(response[d]))
    outline = outline.replace('!!OUTCOMES!!', ascii(response[o]))
    method = ascii(response[m])
    if len(method) == 0:
        method = 'Teaching and learning approach is experiential, collaborative, and problem-based'
    outline = outline.replace('!!METHODS!!', method)
    # hardware
    thwmin = ascii(response[hwmin])
    if len(thwmin.strip()) == 0:
        thwmin = 'No specific computer hardware is required.'
    thwrec = ascii(response[hwrec])
    if len(thwrec) > 0:
        thwrec = '\\subsubsection{Recommended computer hardware}\n\n' + thwrec
    outline = outline.replace('!!HARDWARE!!', thwmin + thwrec)                
    # admin 
    if ascii(response[admin]) == 'yes':
        outline = outline.replace('!!ADMIN!!', '\\input{admin.tex}')
    else:
        outline = outline.replace('!!ADMIN!!', '')
    # software 
    tswmin = ascii(response[swmin])
    if len(tswmin.strip()) == 0:
        tswmin = 'No specific software, operating system, or online service is required.'    
    tswrec = ascii(response[swrec])
    if len(tswrec.strip()) > 0:
        tswrec = '\\subsubsection{Recommended software and services}\n\n' + tswrec
    outline = outline.replace('!!SOFTWARE!!', tswmin + tswrec)        
    # internet
    timin = ascii(response[imin])
    if len(timin.strip()) == 0:
        timin = 'There are no specific requirements regarding the type of internet connection for this course.'        
    tirec = ascii(response[irec])
    if len(tirec) > 0:
        tirec = '\\subsubsection{Recommended internet connection}\n\n' + tirec         
    outline = outline.replace('!!INTERNET!!', timin + tirec)
    # READINGS
    required = ascii(response[req])
    if len(required) == 0:
        required = 'Readings and assignments provided through myCourses'
    outline = outline.replace('!!READINGS!!', ascii(required))
    optional = ascii(response[opt])
    if len(optional) > 0:
        optional = f'\\subsection{{Optional Materials}}\n\n{optional}\n'
    outline = outline.replace('!!OPTIONAL!!', ascii(optional))
    try:
        attendance = int(response[a]) # should be a number
    except:
        attendance = 0 # zero if blank    
    expl = ascii(response[e])
    if attendance > 0 and len(expl) == 0:
        expl = DEFAULT
    graded = [ (attendance,
                'Attendance and active participation',
                'See myCourses for more information', expl )]
    total = float(attendance)
    for entries in group(response[g]):
        if len(entries) >= 2:
            perc = entries[0]
            contrib = 0
            value = 0
            try:
                value = float(perc)
            except:
                pass
            total += value
            name = entries[1]
            due = entries[2] if len(entries) > 2 else 'To be defined'
            detail = ascii(entries[3]) if len(entries) > 3 else 'To be made available on myCourses'
            graded.append( (perc, name, due, detail) )
    if fabs(total - 100) > THRESHOLD:
        error = f'Parsing identified {total} percent for the grade instead of 100 percent.'
    items = '\\\\\n\\hline\n'.join([ f'{ip} & {it} & {idl} & {idesc}' for (ip, it, idl, idesc) in graded ])
    outline = outline.replace('!!ITEMS!!', items)
    sessions = [ ascii(response[header.index(f'Session {k}')]) for k in range(1, 14) ]
    content = '\n'.join([ f'\\item{{{ascii(r)}}}' if len(r) > 0 else '' for r in sessions ])
    outline = outline.replace('\\item{!!CONTENT!!}', ascii(content))
    outline = outline.replace('!!INFO!!', '\\textcolor{blue}{Complementary information to be inserted soon.}')    
    if len(error) > 0:
        error = f'\\textcolor{{red}}{{{error}}}'
        if output not in wrong and exists(output):
            print('Omitting a broken version for', output, 'since an error-free one exists')
            continue
        wrong.add(output)
        print('Storing a version of', output, 'that contains errors:, error')
    else:
        if output in wrong:
            wrong.remove(output)
        print('Storing an error-free version of', output)
    outline = outline.replace('!!ERROR!!', error)
    with open(output, 'w') as target:
        print(outline, file = target)
    
