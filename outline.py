# pip install pandas
# pip install openpyxl
import pandas as pd
from math import fabs
from datetime import datetime
from os.path import exists


# load the course information sheet
info = pd.read_csv('courses.csv')
TAdef = 'Teaching_Assistants.xlsx'

# update the term here
fixed = 'Winter 2025'
TAsheet = {   'Fall 2022': 'TAs_fall_2022',
            'Winter 2023': 'TAs_winter_2023',
            'Winter 2024': 'TAs_Winter_2024',
            'Winter 2025': 'TAs_Winter_2025',
            'Winter 2025': 'TAs_Summer_2025',
            'Summer 2023': 'TAs_Summer_2023',
            'Summer 2024': 'TAs_Summer_2024',            
              'Fall 2023': 'TAs_Fall_2023',
              'Fall 2024': 'TAs_Fall_2024'} 

skip = False 
debug = False
THRESHOLD = 0.05
wrong = set()

def unquote(text):
    while True:
        if text is None or len(text) == 0:
            return ''
        text = text.strip()
        if len(text) > 1 and text[0] == '"' and text[-1] == '"':
            text = text[1:-1]
        else:
            return text

def o2l(t):
    if ";" not in t:
        return t
    s = '\\vspace*{-2mm}#s#\n'
    for part in t.split(';'):
        part = unquote(part.strip())
        if len(part) > 0:
            s += '#i# ' + part + '\n'
    return s + '#e#';

def capsfix(text):
    words = text.split()
    fix = []
    for word in words:
        if word.isupper():
            fix.append(word.capitalize())
        else:
            fix.append(word)
    return ' '.join(fix)

def ascii(text):
    text = unquote(text)
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

    if '#i#' in text:
        if '#s#' not in text and '#ns#' not in text:
            text = text.replace('#i#', '#s#\n#i#', 1) # just the first one, assume bullets
    if '#s#' in text and '#e#' not in text: # fix for reza
        text = text + '\n#e#\n' # assume the list keeps going until the end
    if '#ns#' in text and '#ne#' not in text: # fix for reza
        if '#e#' in text: # mismatch, maybe?
            text = text.replace("#e#", '#ne#');
        else:
            text = text + '\n#ne#\n' # assume the list keeps going until the end                   
    
    text = text.replace('#ns#', '\\begin{enumerate}') # for sam
    text = text.replace('#ne#', '\n\\end{enumerate}\n\n')
    text = text.replace('#s#', '\\begin{itemize}')
    text = text.replace('#e#', '\n\\end{itemize}\n\n')
    text = text.replace('#i#', '\n\\item ')
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
    joint = joint.replace('.  ', '.\n\n') # paragraphs
    if joint[:7] == '\\begin{': # starts with a bulleted list
        joint = '\phantom{skip}\\\\\\vspace*{-2mm}' + joint
    return joint

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
    parts = s.split()
    if len(parts) == 0:
        return True
    return not parts[0].isdigit()

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


def cleanterm(s):
    s = s.strip().lstrip()
    if len(s) == 0:
        return None
    c = None
    digits = None
    if len(s) == 3:
        if 'F' in s:
            c = 'Fall '
        elif 'W' in s:
            c = 'Winter '
        elif 'S' in s:
            c = 'Summer '
        digits = s[1:]
    elif '20' in s and len(s) == 6: # 202309 or similar
        digits = s[:4]
        t = s[4:]
        if t == '09':
            c = 'Fall '
        elif t == '01':
            c = 'Winter '
        elif t == '05':
            c = 'Summer '
    else:
        if 'summer' in s.lower() or 'spring' in s.lower():
            c = 'Summer '
        elif 'winter' in s.lower():
            c = 'Winter '
        elif 'fall' in s.lower() or 'autumn' in s.lower():
            c = 'Fall '
    if c is None:
        print(s, 'is not a valid term')
        quit()
    if digits is None:
        digits = ''.join(i for i in s if i.isdigit())
    if len(digits) == 2:
        digits = '20' + digits
    if len(digits) != 4:
        print(digits, 'is not a valid year')
        quit()
    return c + digits

# load the TA information
from collections import defaultdict
assistant = defaultdict(list)
for term in TAsheet:
    TAinfo = pd.ExcelFile(TAdef)
    try:
        TAs  = TAinfo.parse(TAsheet[term])
    except:
        continue
    print(TAs.columns.values.tolist())
    TAh = [h.strip() for h in TAs.columns.values.tolist()]
    tacl = TAh.index('Course') if 'Course' in TAh else TAh.index('Course code')
    tacn = None
    try:
        tacn = TAh.index('Code') if 'Code' in TAh else TAh.index('Number')
    except:
        tacn = None
    tas = TAh.index('Section')
    tat = TAh.index('Teaching/Course Assistant')
    tan = TAh.index('Candidate') if 'Candidate' in TAh else TAh.index('Candidate Name')
    te = TAh.index('Mcgill email') if 'Mcgill email' in TAh else TAh.index('McGill Email') 
    tok = TAh.index('Workday Status')

    for index, row in TAs.iterrows():
        name = row[tan]
        email = str(row[te]).lstrip().strip()
        if '@' not in email:
            email = '' # blank out the unavailable
        else:
            email = f' ({email})'
        if not isinstance(name, str):
            continue
        name = name.lstrip().strip()
        code = str(row[tacl]).strip().lstrip()
        number = None
        if len(code) > 4:
            code = code.replace(' ', '')
            number = int(code[4:])
            code = code[:4]
        if len(code) != 4:
            print('Invalid course code', code)
            continue
        if number is None:
            number = int(row[tacn])
        section = int(row[tas])
        kind = row[tat]
        ast = 'Teaching Assistant' if kind == 'TA' else 'Course Assistant'
        if 'TA 120' in name:
            name = name.replace('TA 120', '')
        if 'TBC' in name:
            print('Confirmation pending for TA', name)
            name = '\\textcolor{red}{To be confirmed}'
            email = ''
        name = name.strip().lstrip()
        if 'low registration' in name or len(name) == 0:
            continue
        details = f'\item[{ast}]{{{name} {email}}}'
        number = '{:03d}'.format(number)
        section = '{:03d}'.format(section)
        status = str(row[tok])
        # if  'hired' in status or 'accepted' in status: # Nadia asked to disable this and go by the names
        assistant[f'{term} {code} {number} {section}'].append(details)
        print(term, code, number, section, details, status)

allbymyself = [ '' ] # no TA, no CA (say nothing for now)
    
# load the template
with open('outline.tex') as source:
    template = source.read()
    
# load the outline responses
data = None
from sys import argv
if 'msforms' in argv:
    responses = pd.ExcelFile('outline.xlsx')
    data  = responses.parse(responses.sheet_names[0])
else:
    data = pd.read_csv('sharepoint.csv', header = [0], engine = 'python')

header = [h.strip().lower() for h in data.columns.values.tolist()]
HWMIN = 'required minimum computer hardware specifications'
HWREC = 'recommended computer hardware specifications'
SWMIN = 'required software and services'
SWREC = 'recommended software and services'
ADMIN = 'administrative account'
IMIN = 'required minimum internet connection specifications'
IREC = 'recommended internet connection specifications'

when = None
if 'term' in header: 
    when = header.index('term')
    
t = header.index('course title') 
# sharepoint likes to be french, sometimes
n = header.index('course number') if 'course number' in header \
    else (header.index('title') if 'title' in header else header.index('titre'))

s = header.index('section number')
prof = header.index('instructor(s)')
oh = header.index('office hours')
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
    if 'session' in current.lower():
        header[i] = current.lower().replace('session', 's').replace(' ', '')
    elif 'admin' in current.lower():
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
    term = cleanterm(response[when].strip()) if when is not None else None
    if term is None:
        term = fixed # default since we did not ask for the term in the start
    shortterm = term[0] + term[-2:]
    code = response[n].strip().lstrip() # course number
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
        lettercode = code[:4]
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
        ct = ascii(response[t]).lower().title() # course title
        outline = outline.replace('!!NAME!!', capsfix(ct))
        outline = outline.replace('!!CODE!!', code)
        outline = outline.replace('!!SECTION!!', section)
        outline = outline.replace('!!INSTRUCTOR!!', contact(response[prof].strip())) # instructor
        hours = response.get(oh, '')
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
            amount = float(additionalDetails['Credit amount'].iloc[0])
            amount = '{:.1f}'.format(amount)
            if '0' == amount[-1]:
                amount = amount[:-2] # no .0
            kind = ''
            if additionalDetails['Credit type'].iloc[0] == 'credits':
                outline = outline.replace('!!COURSEEVAL!!', '\\input{mercury.tex}')                
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
                outline = outline.replace('!!COURSEEVAL!!', '\\input{limesurvey.tex}')                                
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
                                          f'\\item[Independent study hours]{{Approximately {h} hours in total}}')
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
        thwmin = '\\subsubsection{Required computer hardware}\n\n' + thwmin            
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
        tswmin = '\\subsubsection{Required software and services}\n' + tswmin
        tswrec = ascii(o2l(response[swrec]))
        if len(tswrec.strip()) > 0:
            tswrec = '\\subsubsection{Recommended software and services}\n' + tswrec
        outline = outline.replace('!!SOFTWARE!!', tswmin + tswrec)        
        # internet
        timin = ascii(o2l(response[imin]))
        if len(timin.strip()) == 0:
            timin = 'There are no specific requirements regarding the type of internet connection for this course.'        
        tirec = ascii(o2l(response[irec]))
        if len(tirec) > 0:
            tirec = '\\subsubsection{Recommended internet connection}\n' + tirec         
        outline = outline.replace('!!INTERNET!!', timin + tirec)

        # READINGS
        required = ascii(str(response[req]))
        if len(required) == 0:
            required = '\\subsection{Readings}Readings and assignments provided through myCourses.'
        else:
            required = f'\\subsection{{Required Readings}}\n{required}\n'
        outline = outline.replace('!!READINGS!!', required)

        optional = ascii(str(response[opt]))
        if len(optional) > 0:
            optional = f'\\subsection{{Optional Materials}}\n{optional}\n\\newpage'
        outline = outline.replace('!!OPTIONAL!!' , optional)

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
            expl = '\\input{part.tex}'
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
        items = '\\\\\n\\hline\n'.join([ f'{it} & {idl} & {ip} & {idesc}' for (ip, it, idl, idesc) in assessments ])
        outline = outline.replace('!!ITEMS!!', items)
        # at most 15 sessions as of now (Nabil and Diana Oka)
        sessions = [ ascii(response[header.index(f's{k}')]) for k in range(1, 16) ] # <---- NEED MORE SESSIONS???
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
