# pip install pandas
# pip install openpyxl
import pandas as pd
from math import fabs

THRESHOLD = 0.02
DEFAULT = 'This course consists of a community of learners of which you are an integral member; your active participation is therefore essential to its success. This means: attending class; visiting myCourses, doing the assigned readings/exercises before class; and engaging in class discussions/activities.'

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
    return '\n'.join(clean)

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
    line = line.replace('//', '\n') # this is for Mike
    line = line.replace('- ', '\n') # this is for Mohammad    
    line = ascii(line)
    lines = []
    if '\n' not in line: # a single-line answer
        tokens = line.split(';')
        while len(tokens) > 4:
            lines.append(tokens[:4]) # supposed to have four fields per line
            tokens = tokens[4:]
        lines.append(tokens)
        return lines
    else:
        for entry in line.split('\n'):
            lines.append(entry.split(';'))
    return lines

# load the course information sheet
info = pd.read_csv('courses.csv')

# load the template
with open('outline.tex') as source:
    template = source.read()
template = template.replace('!!TERM!!', 'Fall 2022') # update the term/year
    
# load the outline responses
responses = pd.ExcelFile('outline.xlsx')
forms  = responses.parse(responses.sheet_names[0])
# drop the first five columns
forms = forms.iloc[: , 5:]
print('Responses from MS Forms', forms.shape)
header = [h.strip() for h in forms.columns.values.tolist()]
t = header.index('Course title')
n = header.index('Course number')
s = header.index('Section number')
i = header.index('Instructor(s)')
h = header.index('Office hours')
d = header.index('Course description')
o = header.index('Learning outcomes')
req = header.index('Required course material')
opt = header.index('Optional course material')
m = header.index('Instructional methods')
a = header.index('% for Attendance and active participation')
e = header.index('Explanation for Attendance and active participation')
g = header.index('Other graded items')
edited = pd.read_csv('sharepoint.csv', header = [0], engine = 'python') # these have double quotes
print('Responses from Sharepoint', edited.shape)
edited.columns = forms.columns # same header
data = pd.concat([forms, edited], axis = 0, ignore_index = True)
print('Combined responses', data.shape)
data.fillna('', inplace = True)

fields = dict()
for index, response in data.iterrows():
    error = ''
    code = response[n].strip() # course number
    print('Processing', code)
    section = response[s].strip() # course section
    if len(section) == 0 and '-' in code:
        parts = code.split('-')
        code = parts[0].strip()
        section = parts[1]
    if 'CRN' in section:
        error = '\nExtra details provided in section number\n'
        if '(' in section:
            section = section.split('(').pop(0).strip()
    if 'Fall' in section:
        error = '\nActual section number is missing\n'
        section = 'XXX'
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
    code = code.replace(' ', '') # no spaces in the filename
    output = f'{code}-{section}.tex'
    outline = outline.replace('!!NAME!!', ascii(response[t])) # course title
    outline = outline.replace('!!CODE!!', code)
    outline = outline.replace('!!SECTION!!', section)
    outline = outline.replace('!!INSTRUCTOR!!', contact(response[i].strip())) # instructor
    # pending: insert information on TA/CA into !!ASSISTANT!!
    outline = outline.replace('!!ASSISTANT!!', '\\textcolor{blue}{Teaching/course assistant information to be inserted soon.}')
    hours = response.get(h, '')
    if len(hours) == 0:
        hours = 'Upon request'
    outline = outline.replace('!!HOURS!!', hours.strip())
    details = None
    if lettercode is not None and numbercode is not None:
        details = info.loc[(info['Code'] == lettercode) & (info['Number'] == int(numbercode))]
    else:
        error += '\nCourse code specification not found'
    if details is None or details.empty:
        error += '\nCourse details are not in the domain catalogue, corresponding fields will not be populated\n'
    else:
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
        if details['Credit type'].iloc[0] == 'credits':
            outline = outline.replace('!!CREDITS!!', f'{amount} credits')
        else:
            outline = outline.replace('!!CREDITS!!', f'{amount} CEUs')
        h = details['Contact hours'].iloc[0]
        outline = outline.replace('!!CONTACT!!', f'{h} hours')
        h = details['Approximate assignment hours'].iloc[0]
        outline = outline.replace('!!ASSIGNMENT!!', f'{h} hours')

    outline = outline.replace('!!DESCRIPTION!!', ascii(response[d]))
    outline = outline.replace('!!OUTCOMES!!', ascii(response[o]))
    method = ascii(response[m])
    if len(method) == 0:
        method = 'Teaching and learning approach is experiential, collaborative, and problem-based'
    outline = outline.replace('!!METHODS!!', method)
    required = ascii(response[req])
    if len(required) == 0:
        required = 'Readings and assignments provided through myCourses'
    else:
        required = required.replace('. ', '.\n\n')        
    outline = outline.replace('!!REQUIRED!!', ascii(required))
    optional = ascii(response[opt])
    if len(optional) > 0:
        optional = optional.replace('. ', '.\n\n')
        optional = f'\\subsection{{Optional Materials}}\n\n{optional}\n'
    outline = outline.replace('!!OPTIONAL!!', ascii(optional))
    attendance = response[a]
    expl = ascii(response[e])
    if len(expl) == 0:
        expl = DEFAULT
    graded = [ (attendance,
                'Attendance and active participation',
                'See myCourses for more information', expl )]
    total = float(attendance)
    for entries in group(response[g]):
        if len(entries) >= 2:
            perc = entries[0]
            contrib = 0
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
    outline = outline.replace('!!ERROR!!', error)
    with open(output, 'w') as target:
        print(outline, file = target)
    
