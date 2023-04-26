# pip install pandas
# pip install openpyxl
import pandas as pd
from math import fabs
from datetime import datetime
from os.path import exists

skip = False # hey yaz
debug = False
THRESHOLD = 0.05
DEFAULT = 'This course consists of a community of learners of which you are an integral member; your active participation is therefore essential to its success. This means: attending class; visiting myCourses, doing the assigned readings/exercises before class; and engaging in class discussions/activities.'

wrong = set()

def o2l(t):
    if ";" not in t:
        return t
    s = '#s#\n'
    for part in t.split(';'):
        part = part.strip()
        if len(part) > 0:
            s += '#i# ' + part + '\n'
    return s + '#e#';

def capsfix(text):
    words = text.split()
    fixed = []
    for word in words:
        if word.isupper():
            fixed.append(word.capitalize())
        else:
            fixed.append(word)
    return ' '.join(fixed)

def ascii(text):
    text = text.replace('%23', '#') # sharepoint export breaks the #    

    text = text.replace('á', "\\'{a}")
    text = text.replace('é', "\\'{e}")
    text = text.replace('í', "\\'{\\i}")
    text = text.replace('ó', "\\'{o}")
    text = text.replace('ú', "\\'{u}")

    text = text.replace('à', "\\`{a}")
    text = text.replace('è', "\\`{e}")
    text = text.replace('ì', "\\`{\\i}")
    text = text.replace('ò', "\\`{o}")
    text = text.replace('ù', "\\`{u}")

    text = text.replace('ä', '\\"{a}')
    text = text.replace('ë', '\\"{e}')
    text = text.replace('ï', '\\"{\\i}')
    text = text.replace('ö', '\\"{o}')
    text = text.replace('ü', '\\"{u}')

    text = text.replace('Á', "\\'{A}")
    text = text.replace('É', "\\'{E}")
    text = text.replace('Í', "\\'{I}")
    text = text.replace('Ó', "\\'{U}")
    text = text.replace('Ú', "\\'{U}")

    text = text.replace('À', "\\`{A}")
    text = text.replace('È', "\\`{E}")
    text = text.replace('Ì', "\\`{I}")
    text = text.replace('Ò', "\\`{O}")
    text = text.replace('Ù', "\\`{U}")
    
    text = text.replace('Ä', '\\"{A}')
    text = text.replace('Ë', '\\"{E}')
    text = text.replace('Ï', '\\"{I}')
    text = text.replace('Ö', '\\"{O}')
    text = text.replace('Ü', '\\"{U}')

    text = text.replace('ñ', '\\~{n}')
    text = text.replace('Ñ', '\\~{N}')

    text = text.replace('ç', '\\c{c}')
    text = text.replace('Ç', '\\c{C}')

    # a fix for jean-philippe
    text = text.replace('<', ' ')
    text = text.replace('>', ' ')    
    if '#i#' in text and '#s#' not in text:
        text = text.replace('#i#', '#s#\n#i#', 1) # just the first one
    if '#s#' in text and '#e#' not in text: # fix for reza
        text = text + '\n#e#\n' # assume the list keeps going until the end            
    text = text.replace('#s#', '\n\\begin{itemize}')
    text = text.replace('#i#', '\n\\item ')
    text = text.replace('#e#', '\n\\end{itemize}\n\n')
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
    joint = '\n'.join(clean).strip().lstrip()
    if len(joint) == 0:
        return ''
    if joint[0] == '"' and joint[-1] == '"':
        joint = joint[1:-1]
    return joint.replace('.  ', '.\n\n') # paragraphs

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
            link = f'\\href{{mailto:{word}}}{{{word}}} \\\\ ' 
            clean.append(link)
        else:
            clean.append(word)
    return ' '.join(clean)

def notanumber(s): # alejandro's fake line breaks
    return not s.split()[0].isdigit()

def group(line):
    line = line.strip()
    if len(line) > 0 and line[0] == line[-1] and line[0] == '"':
        line = line[1:-1]
    lines = []
    for entry in line.split('\n'):
        if len(entry) == 0:
            continue
        cols = entry.split(';')
        if len(cols) > 0:
            if notanumber(cols[0]):
                if len(lines) > 0: # somewhere to append
                    while len(cols) > 0 and notanumber(cols[0]):
                        lines[-1][-1] += ' ' + cols.pop(0)
                    if len(cols) > 0: # the last two columns (hours and minutes)
                        lines[-1] += cols
            else:
                lines.append(cols)
    return lines

# load the course information sheet
info = pd.read_csv('courses.csv')

TAdef = 'Teaching_Assistants.xlsx'
TAfile = { 'Summer 2023': 'Teaching_Assistants_S23.xlsx'}
TAsheet = { 'Fall 2022': 'TAs_fall_2022',
            'Winter 2023': 'TAs_winter_2023',
            'Summer 2023': 'TAs_Summer_2023' }            

# load the TA information
from collections import defaultdict
assistant = defaultdict(list)
for term in TAsheet:
    TAinfo = pd.ExcelFile(TAfile.get(term, TAdef))
    TAs  = TAinfo.parse(TAsheet[term])
    TAh = [h.strip() for h in TAs.columns.values.tolist()]
    tacl = TAh.index('Course') if 'Course' in TAh else TAh.index('Course code')
    tacn = TAh.index('Code') if 'Code' in TAh else TAh.index('Number')
    tas = TAh.index('Section')
    tat = TAh.index('Teaching/Course Assistant')
    tan = TAh.index('Candidate')
    te = TAh.index('McGill Email')
    tok = TAh.index('Workday Status')
    for index, row in TAs.iterrows():
        name = row[tan]
        email = str(row[te]).lstrip().strip()
        if '@' not in email:
            email = '' # blank out the unavailable
        else:
            email = f'({email})'
        if not isinstance(name, str):
            break
        name = name.lstrip().strip()
        code = str(row[tacl]).strip()
        if len(code) != 4:
            print('Invalid course code', code)
            continue
        number = int(row[tacn])
        section = int(row[tas])
        kind = row[tat]
        ast = 'Teaching Assistant' if kind == 'TA' else 'Course Assistant'
        if 'TA 120' in name:
            name = name.replace('TA 120', '')
        if 'low registration' in name:
            continue
        details = f'\item[{ast}]{{{name} {email}}}'
        print(term, code, number, section, details, tok)
        number = '{:03d}'.format(number)
        section = '{:03d}'.format(section)
        if tok == 'hired' or tok == 'offer sent':
            assistant[f'{term} {code} {number} {section}'].append(details)

allbymyself = [ '' ] # no TA, no CA (say nothing for now)
    
# load the template
with open('outline.tex') as source:
    template = source.read()
    
# load the outline responses
data = None
from sys import argv
fixed = None
if 'sharepoint' not in argv:
    responses = pd.ExcelFile('outline.xlsx')
    data  = responses.parse(responses.sheet_names[0])
else:
    data = pd.read_csv('sharepoint.csv', header = [0], engine = 'python')
    fixed = 'Summer 2023'

header = [h.strip().lower() for h in data.columns.values.tolist()]
HWMIN = 'required minimum computer hardware specifications'
HWREC = 'recommended computer hardware specifications'
SWMIN = 'required software and services'
SWREC = 'recommended software and services'
ADMIN = 'administrative account'
IMIN = 'required minimum internet connection specifications'
IREC = 'recommended internet connection specifications'

fixed = 'Summer 23'
if fixed is None:            
    when = header.index('term')

t = header.index('course title') 
n = header.index('course number') if 'course number' in header \
    else (header.index('title') if 'title' in header else header.index('titre')) # sharepoint likes to be french, sometimes
s = header.index('section number')
prof = header.index('instructor(s)')
h = header.index('office hours')
d = header.index('course description')
o = header.index('learning outcomes')
req = header.index('required readings')
opt = header.index('additional optional course material')
m = header.index('instructional methods')
a = header.index('% for attendance and active participation')
e = header.index('explanation for attendance and active participation')
graded = header.index('other graded items')
adinfo = header.index('additional information')

hl = len(header)

for i in range(hl):
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

hwmin = header.index(HWMIN)
hwrec = header.index(HWREC)
swmin = header.index(SWMIN)
swrec = header.index(SWREC)
admin = header.index(ADMIN)
imin = header.index(IMIN)
irec = header.index(IREC)

print('Responses', data.shape)
data.fillna('', inplace = True)

fields = dict()

for index, response in data.iterrows():
    lr = len(response)
    assert lr == hl # multiline bug detection 
    error = ''
    # when is this taught
    term = fixed if fixed is not None else response[when].strip() 
    if term == '':
        term = 'Fall 2022' # default since we did not ask for the term in the start
    shortterm = term[0] + term[-2:]
    code = response[n].strip() # course number
    code = code.replace('-', ' ') # for Hugue
    code = ' '.join(code.split()) # also for Hugue
    if len(code) < 3:
        print('Skipping an empty row')
        continue
    sections = str(response[s]).strip() # course section
    if len(sections) == 0:
        error = '\nMissing section number\n'
    if 'CRN' in sections:
        # error = '\nExtra details provided in section number\n'
        if '(' in sections:
            sections = sections.split('(').pop(0).strip()
    if 'Fall' in sections:
        error = '\nActual section number is missing\n'
        print(code, 'lacks section number, omitting')
        continue
    if '-' in code: # someone wrote <YCIT 001 - 001>
        parts = code.split('-')
        sections = parts[-1].strip()
        code = parts[0].strip()
    lettercode = None
    numbercode = None
    if ' ' in code:
        parts = code.split()
        lettercode = parts.pop(0)
        numbercode = parts.pop(0)
    else:
        lettercode = code[:3]
        numbercode = code[4:]
    if ',' in sections:
        sections = sections.replace(',', ' ')
    for section in sections.split():
        section = '{:03d}'.format(int(section))
        outline = template.replace('!!TERM!!', term) 
        outline = outline.replace('!!CODE!!', code)
        outline = outline.replace('!!SECTION!!', section)
        print('Retrieving CA/TA for', code, section)
        assigned = assistant.get(f'{term} {code} {section}', allbymyself)
        assistants = '\n\n'.join(assigned)
        outline = outline.replace('!!ASSISTANT!!', assistants)
        code = code.replace(' ', '') # no spaces in the filename
        section = str(section)
        while len(section) < 3:
            section = '0' + section
        if len(code) != 7:
            print(f'Wrong code length, making {code} {section} into TEST123-000')
            lettercode = 'TEST'
            numbercode = '123'
            code = lettercode + numbercode
            section = '000'
        output = f'{shortterm}-{code}-{section}.tex'
        if exists(output):
            print('Reprocessing', output)
        else:
            print('Processing', output)
        ct = ascii(response[t]) # course title            
        outline = outline.replace('!!NAME!!', capsfix(ct))
        outline = outline.replace('!!CODE!!', code)
        outline = outline.replace('!!SECTION!!', section)
        outline = outline.replace('!!INSTRUCTOR!!', contact(response[prof].strip())) # instructor
        hours = response.get(h, '')
        if len(hours) == 0:
            hours = 'Upon request'
        outline = outline.replace('!!HOURS!!', hours.strip())
        additionalDetails = None
        if lettercode is not None and numbercode is not None:
            print('Extracting additional details for', lettercode, numbercode)
            additionalDetails = info.loc[(info['Code'] == lettercode) & (info['Number'] == int(numbercode))]
        else:
            error += '\nCourse code specification not found'
        if additionalDetails is None or additionalDetails.empty:
            error += '\nCourse details not in the domain catalogue, corresponding fields will not be populated\n'
        else:
            graduate = lettercode[-1] == '2' or numbercode[0] == '6' or numbercode[0] == '5'
            if graduate:
                print(lettercode, numbercode, 'is a graduate course')
            prereq = additionalDetails['Pre-requisites'].iloc[0]
            if pd.isna(prereq):
                outline = outline.replace('!!PREREQ!!', 'No pre-requisites')        
            else:
                outline = outline.replace('!!PREREQ!!', prereq)
            coreq = additionalDetails['Co-requisites'].iloc[0]
            if pd.isna(coreq):
                outline = outline.replace('!!COREQ!!', 'No co-requisites')
            else:
                outline = outline.replace('!!COREQ!!', coreq)
            amount = additionalDetails['Credit amount'].iloc[0]
            amount = round(int(amount)) # no .0
            kind = ''
            if additionalDetails['Credit type'].iloc[0] == 'credits':
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
            h = additionalDetails['Contact hours'].iloc[0]
            if h is None or h == 'None' or pd.isna(h): # credit-side default is 39
                h = 39
            else:
                h = round(int(h)) # no .0
            outline = outline.replace('!!CONTACT!!', f'{h} hours')
            h = additionalDetails['Approximate assignment hours'].iloc[0]
            if h is None or h == 'None' or pd.isna(h):
                outline = outline.replace('!!ASSIGNMENT!!','') # nothing goes here
            else:
                h = round(int(h)) # no .0
                outline = outline.replace('!!ASSIGNMENT!!',
                                          f'\\item[Independent study hours]{{Approximately {h} hours }}')
        outline = outline.replace('!!DESCRIPTION!!', ascii(response[d]))
        outline = outline.replace('!!OUTCOMES!!', ascii(response[o]))
        method = ascii(str(response[m]))
        if len(method) == 0:
            method = 'Teaching and learning approach is experiential, collaborative, and problem-based'
        outline = outline.replace('!!METHODS!!', method)
        # hardware
        thwmin = ascii(o2l(response[hwmin]))
        if len(thwmin.strip()) == 0:
            thwmin = 'No specific computer hardware is required.'
        thwrec = ascii(o2l(response[hwrec]))
        if len(thwrec) > 0:
            thwrec = '\\subsubsection{Recommended computer hardware}\n\n' + thwrec
        outline = outline.replace('!!HARDWARE!!', thwmin + thwrec)                
        # admin 
        if ascii(response[admin]) == 'yes':
            outline = outline.replace('!!ADMIN!!', '\\input{admin.tex}')
        else:
            outline = outline.replace('!!ADMIN!!', '')
        # software 
        tswmin = ascii(o2l(response[swmin]))
        if len(tswmin.strip()) == 0:
            tswmin = 'No specific software, operating system, or online service is required.'    
        tswrec = ascii(o2l(response[swrec]))
        if len(tswrec.strip()) > 0:
            tswrec = '\\subsubsection{Recommended software and services}\n\n' + tswrec
        outline = outline.replace('!!SOFTWARE!!', tswmin + tswrec)        
        # internet
        timin = ascii(o2l(response[imin]))
        if len(timin.strip()) == 0:
            timin = 'There are no specific requirements regarding the type of internet connection for this course.'        
        tirec = ascii(o2l(response[irec]))
        if len(tirec) > 0:
            tirec = '\\subsubsection{Recommended internet connection}\n\n' + tirec         
        outline = outline.replace('!!INTERNET!!', timin + tirec)

        # READINGS
        required = ascii(str(response[req]))
        if len(required) == 0:
            required = '\\subsection{Readings}Readings and assignments provided through myCourses.'
        else:
            required = f'\\subsection{{Required Readings}}\n\n{required}\n'
        outline = outline.replace('!!READINGS!!', required)

        optional = ascii(str(response[opt]))
        if len(optional) > 0:
            optional = f'\\subsection{{Optional Materials}}\n\n{optional}\n'
        outline = outline.replace('!!OPTIONAL!!', optional)

        # ADDITIONAL INFO
        ai = ascii(response[adinfo])
        if len(ai) > 0:
            aic = f'{ai}\n'
            outline = outline.replace('!!ADDITIONAL!!', aic)
        else:
            outline = outline.replace('!!ADDITIONAL!!', '') # blank        
        try:
            attendance = int(response[a]) # should be a number
        except:
            attendance = 0 # zero if blank    
        expl = ascii(response[e])
        if attendance > 0 and len(expl) == 0:
            expl = DEFAULT
        assessments = []
        if attendance > 0:
            assessments.append((attendance,
                                'Attendance and active participation',
                                'See myCourses for more information', expl ))
        total = float(attendance)
        items = ascii(response[graded])
        if debug:
            idx = 0
            for r in response:
                rs = str(r)
                if len(rs) > 10:
                    rs = rs[:10]
                print(idx, header[idx], '\n  ', rs, '\n')
                idx += 1
            print('###', items)
        for entries in group(items):
            if len(entries) >= 2:
                perc = entries[0].replace('%', '').replace('\\', '')
                contrib = 0
                value = 0
                try:
                    value = float(perc)
                except:
                    pass
                if value > 0:
                    total += value
                    name = ascii(entries[1])
                    due = entries[2] if len(entries) > 2 else 'To be defined'
                    assessmentDetail = ascii(entries[3]) if len(entries) > 3 else 'To be made available on myCourses'
                    assessments.append( (perc, name, due, assessmentDetail) )
        if fabs(total - 100) > THRESHOLD:
            error = f'\nParsing identified {total} percent for the grade instead of 100 percent.'
        items = '\\\\\n\\hline\n'.join([ f'{ip} & {it} & {idl} & {idesc}' for (ip, it, idl, idesc) in assessments ])
        outline = outline.replace('!!ITEMS!!', items)
        sessions = [ ascii(response[header.index(f'session {k}')]) for k in range(1, 14) ]
        sessions = [ s.strip() for s in sessions ]
        content = '\n'.join([ f'\\item{{{ascii(r)}}}' if len(r) > 0 else '' for r in sessions ])
        outline = outline.replace('\\item{!!CONTENT!!}', ascii(content))
        outline = outline.replace('!!INFO!!', '\\textcolor{blue}{Complementary information to be inserted soon.}')    
        if len(error) > 0:
            error = f'\\textcolor{{red}}{{{error}}}'
            if skip and output not in wrong and exists(output):
                print('Omitting a broken version for', output, 'since an error-free one exists')
                continue
            wrong.add(output)
            print('Storing a version of', output, 'that contains errors:', error)
        else:
            if output in wrong:
                wrong.remove(output)
            print('Storing an error-free version of', output)
        outline = outline.replace('!!ERROR!!', error)
        with open(output, 'w') as target:
            print(outline, file = target)
