import pandas as pd
from math import fabs

import pandas as pd
from math import fabs

THRESHOLD = 0.02
DEFAULT = 'This course consists of a community of learners of which you are an integral member; your active participation is therefore essential to its success. This means: attending class; visiting myCourses, doing the assigned readings/exercises before class; and engaging in class discussions/activities.'

def ascii(text):
    return ''.join(i for i in text if ord(i) < 128)

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
    for word in text.split():
        if '@' in word:
            link = f'\\href{{mailto:{word}}}{{{word}}}' 
            clean.append(link)
        else:
            clean.append(word)
    return ' '.join(clean)

def group(line):
    line = line.replace('%', '') # no percent symbols are wanted here
    line = line.replace('_', '') # no underscores are permitted here
    line = line.replace('#', '\#')
    line = line.replace('&', '\&')
    line = line.replace('//', '\n') # this is for Mike
    line = line.replace('- ', '\n') # this is for Mohammad
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


def outlineHandler(file_location):
    # load the template
    with open('assets/outline.tex') as source:
        template = source.read()
    template = template.replace('!!TERM!!', 'Fall 2022') # update the term/year
        
    # load the outline responses
    responses = pd.ExcelFile(file_location)
    data  = responses.parse(responses.sheet_names[0])
    data.fillna('', inplace = True)
    fields = dict()

    header = [h.strip() for h in data.columns.values.tolist()]
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

    for index, response in data.iterrows():
        error = ''
        code = response[n].strip() # course number
        section = response[s].strip() # course section
        if len(section) == 0 and '-' in code:
            parts = code.split('-')
            code = parts[0]
            section = parts[1]
        if 'CRN' in section:
            error = 'Extra details provided in section number\n'
            if '(' in section:
                section = section.split('(').pop(0).strip()
        if 'Fall' in section:
            error = 'Actual section number is missing\n'
            section = 'XXX'
        outline = template.replace('!!CODE!!', code)
        outline = outline.replace('!!SECTION!!', section)    
        code = code.replace(' ', '') # no spaces in the filename
        output = f'tmp/{code}-{section}.tex'
        outline = outline.replace('!!NAME!!', response[t].strip()) # course title
        outline = outline.replace('!!CODE!!', code)
        outline = outline.replace('!!SECTION!!', section)
        outline = outline.replace('!!INSTRUCTOR!!', contact(response[i].strip())) # instructor
        # pending: insert information on TA/CA into !!ASSISTANT!!
        outline = outline.replace('!!ASSISTANT!!', '\n\\\\\n\\textcolor{blue}{Course assistant information to be inserted soon.}')
        hours = response[h].strip()
        if len(hours) == 0:
            hours = 'Upon request'
        outline = outline.replace('!!HOURS!!', hours)
        outline = outline.replace('!!DESCRIPTION!!', response[d].strip())
        outline = outline.replace('!!OUTCOMES!!', response[o].strip().replace('&', '\\&'))
        method = response[m].strip().replace('&', '\\&')
        if len(method) == 0:
            method = 'Teaching and learning approach is experiential, collaborative, and problem-based'
        outline = outline.replace('!!METHODS!!', method)
        required = response[req].strip().replace('&', '\\&')
        if len(required) == 0:
            required = 'Readings and assignments provided through myCourses'
        else:
            required = required.replace('. ', '.\n\n')        
        outline = outline.replace('!!REQUIRED!!', ascii(required))
        optional = response[opt].strip().replace('&', '\\&')
        if len(optional) > 0:
            optional = optional.replace('. ', '.\n\n')
            optional = f'\\subsection{{Optional Materials}}\n\n{optional}\n'
        outline = outline.replace('!!OPTIONAL!!', ascii(optional))
        attendance = response[a]
        expl = response[e].strip().replace('&', '\\&')
        if len(expl) == 0:
            expl = DEFAULT
        graded = [ (attendance, 'Attendance and active participation', None, expl )]
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
        sessions = [ response[header.index(f'Session {k}')].replace('&', '\\&').replace('#', '\\#') for k in range(1, 14) ]
        content = '\n'.join([ f'\\item{{{r}}}' if len(r) > 0 else '' for r in sessions ])
        outline = outline.replace('\\item{!!CONTENT!!}', content)
        outline = outline.replace('!!INFO!!', '\\textcolor{blue}{Complementary information to be inserted soon.}')    
        if len(error) > 0:
            error = f'\\textcolor{{red}}{{{error}}}'
        outline = outline.replace('!!ERROR!!', error)
        with open(output, 'w') as target:
            print(outline, file = target)
        
