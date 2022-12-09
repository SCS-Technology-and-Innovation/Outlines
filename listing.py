from bs4 import BeautifulSoup
from random import randint
from time import sleep
from os import path, remove
import requests

NA = 'N/A'
verbose = True

class Course():
    def __init__(self, c, n, a, u):
        self.code = c
        self.number = n
        self.title = NA
        self.status = 'active'
        if a is None:
            self.amount = NA
            print(f'Credit amount missing for {c} {n}')
        else:
            self.amount = a
        if u is None:
            self.unit = NA
            print(f'Credit type missing for {c} {n}')
        else:
            self.unit = u
        self.decsc = NA
        self.contact = NA
        self.assign = NA
        self.prereq = set()
        self.coreq = set()
        self.incl = set()
        self.urls = set()

    def __str__(self):
        p = '\n'. join(self.prereq)
        c = '\n'. join(self.coreq)
        i = '\n'. join(self.incl)
        u = '\n'. join(self.urls)        
        return f'{self.code},{self.number},"{self.title}",{self.status},{self.amount},{self.unit},"{self.descr}",{self.contact},{self.assign},"{p}","{c}","{i}","{u}"'

    def __repr__(self):
        return str(self)

header = 'Code,Number,Title,Status,Credit amount,Credit type,Description,' \
    + 'Contact hours,Approximate assignment hours,Pre-requisites,Co-requisites,' \
    + 'Programs including the course,URL'
    
hdr = { 'User-Agent' :
        'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:97.0) Gecko/20100101 Firefox/97.0' }

PREFIX = 'https://cce.mcgill.ca//itimetable/cpd/ProgramDetails/'

locations = { 'datasci.html' :
              'https://www.mcgill.ca/continuingstudies/area-of-study/data-science',
              'inftech.html' :
              'https://www.mcgill.ca/continuingstudies/area-of-study/information-technology',
              # the website is so bad...
              'inftech2.html':
              'https://www.mcgill.ca/continuingstudies/area-of-study/information-technology?tid=738&page=0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C0%2C1',
              'cybersec.html' : PREFIX + '131' ,
              'compit.html' : PREFIX + '134',
              'busint.html': PREFIX + '106',
              'busana.html' :
              'https://www.mcgill.ca/continuingstudies/program/professional-development-certificate-business-analysis'}

skip = [ 'https://www.mcgill.ca/continuingstudies/registration',
         'https://mcgill.ca/continuingstudies/registration', 
         'https://www.mcgill.ca/continuingstudies/how-register' ]

records = dict()

for listing in locations:
    print('Parsing', listing)
    if not path.exists(listing):
        sleep(randint(1, 5))
        r = requests.get(locations[listing], hdr)
        content = ' '.join(r.text.split()) # tons of whitespace
        with open(listing, 'w') as target:
            print(content, file = target)
        print('Downloaded', listing)            
    with open(listing) as source:
        text = source.read()
    soup = BeautifulSoup(text, 'html.parser')
    if 'timetable' in locations[listing]:
        content = soup.find('div', { 'class' : 'container body-content' })
        for row in content.find_all('tr'):
            columns = row.find_all('td')
            if len(columns) == 0:
                print('Ignoring an empty table row')
                continue
            column = columns.pop(0)
            fields = column.get_text().split()
            if len(fields) >= 2:
                code = fields.pop(0)
                number = fields.pop(0)
                print('Parsing', code, number)
                label = (code, number)
                c = records.get(label, None)
                column = columns.pop(0)
                url = PREFIX + column.find('a')['href']
                filename = f'{code}{number}.html'
                if not path.exists(filename):
                    sleep(randint(1, 5))
                    r = requests.get(url, hdr)
                    content = ' '.join(r.text.split()) # tons of whitespace
                    with open(filename, 'w') as target:
                        print(content, file = target)
                if c is not None:
                    c.urls.add(url)
                name = ' '.join(column.get_text().split())
                name = name.replace(code, '') # no idea why one course repeats these
                name = name.replace(number, '') # ^^^
                name = name.strip().lstrip()
                # abbreviations mess things up
                name = name.replace('Dev.', 'Development')
                name = name.replace('&', 'and')
                name = name.replace('Tech.', 'Technology')
                name = name.replace('Cloud Computer', 'Cloud Computing') # this is an error on a website
                if c is not None:
                    assert name == c.title
                column = columns.pop(0)
                cor = False
                co = None
                s = []
                if c is None: # the business intelligence diploma courses just show up
                    c = Course(code, number, 3, 'credits')
                    c.descr = '' # we do not have these just now
                    c.title = name
                    c.contact = 39
                    records[label] = c
                for field in column.get_text().split():
                    field = field.replace(',', '').strip().lstrip()
                    if 'Co-' in field:
                        cor = True
                    elif 'Pre' in field:
                        continue # default
                    elif len(field) == 4 and field.isupper():
                        co = field # this is a course code
                    elif len(field) == 3 and field.isnumeric():
                        nu = field # this is a course number
                        assert co is not None
                        if cor:
                            if c is not None:
                                c.coreq.add(f'{co} {nu}')
                            cor = False
                        else:
                            if c is not None:
                                c.prereq.add(f'{co} {nu}')
                            co = None
                    else:
                        s.append(field)
                if len(s) > 0:
                    if c is not None:
                        c.prereq.add(' '.join(s)) # written notes
    for course in soup.find_all('div', { 'class' : 'course' }):
        t = course.find('div', { 'class': 'title' })
        if t is None:
            print('NO INFO FOUND IN COURSE BLOCK, aborting.')
            print(course)
            quit()
        name = t.get_text().strip().lstrip()
        cc = course.find('span', { 'class': 'credits' })
        amount = None
        unit = None
        if cc is None:
            if 'CEUs' in name:
                unit = 'CEUs'
                i = name.index('(') + 1
                j = name.index('CEUs') 
                amount = int(name[i:j])
                name = name[:(i-2)]
            elif 'credits' in name:
                unit = 'credits'
                i = name.index('(') + 1
                j = name.index('credits') 
                amount = int(name[i:j])
                name = name[:(i-3)]
        else:
            cr = cc.get_text().lstrip().strip()
            name = name.replace(cr, '')
            cr = cr.strip().lstrip()[1:-1].split() # remove ()
            amount = cr.pop(0)
            unit = cr.pop(0)
        name = name.strip().lstrip().split()
        code = name.pop(0)
        number = name.pop(0)
        label = (code, number)
        if label in records:
            if verbose:
                print(f'Additional data for {code} {number}')
        else:
            if verbose:
                print(f'Located data for {code} {number}')            
            records[label] = Course(code, number, amount, unit)            
        c = records[label]
        name = ' '.join(name)
        name = name.replace(code, '') # no idea why one course repeats these
        name = name.replace(number, '') # ^^^
        name = name.strip().lstrip()        
        if name[-1] == '.':
            name = name[:-1].strip()
        c.title = name
        d = course.find('div', { 'class': 'desc' }).get_text()
        assert '"' not in d
        c.descr = d.strip().lstrip()
        more = course.find('div', { 'class': 'lnk' })
        url = None
        if more is not None:
            url = more.find('a')['href'] 
        if url is not None and url not in skip:
            c.urls.add(url)
            filename = f'{code}{number}.html'
            if not path.exists(filename):
                sleep(randint(1, 5))
                r = requests.get(url, hdr)
                content = ' '.join(r.text.split()) # tons of whitespace
                with open(filename, 'w') as target:
                    print(content, file = target)
            with open(filename) as source:
                info = source.read()
                details = BeautifulSoup(info, 'html.parser')
                h1 = details.find('h1').get_text()
                if 'How to Register' in h1 or 'Registration' in h1:
                    assert url in skip
                    remove(filename) # no need to keep this
                    print(f'Detail page does not exist for {code} {number}')
                    c.status = 'inactive?'                    
                elif 'Course Search Results' in h1: # page no longer exists
                    print(f'Detail page link {url} is broken for {code} {number}')
                    c.status = 'inactive?'                    
                else: # there ARE details
                    assert code in h1 and number in h1
                    n = h1.replace(code, '')
                    n = n.replace(number, '')
                    n = n.replace(' - ', '')                
                    n = n.replace('&', 'and')
                    n = ' '.join(n.split())
                    if code not in h1:
                        print('CODE MISMATCH', code, h1)
                    if number not in h1:
                        print('NUMBER MISMATCH', number, h1)                    
                    tn = c.title.replace('&', 'and')
                    if tn not in n and n not in tn:
                        print(f'Name mismatch for {code} {number}:',
                              f'\n\t<{tn}>\n\tversus\n\t<{n}>')
                    profile = details.find('div', { 'id' : 'courseProfile' })
                    if profile is not None:
                        cv = None
                        av = None
                        profile = profile.get_text()
                        s = 'PD contact hours' 
                        if s in profile: # IIBA
                            i = profile.index(s) + len(s) + 2
                            cv = int(profile[i:(i + 3)])                            
                        else:
                            ch = [ 'hours of classroom instruction',
                                   'hours in class', 'hours class', 
                                   'hours of in-class instruction', 'contact hours',
                                   'hours of lectures']
                            for s in ch:
                                if s in profile:
                                    i = profile.index(s)
                                    cv = int(profile[(i - 3):i])
                                    break
                        ah = [ 'hours of assignments', 
                               'hours of course readings and assignments',
                               'hours of independent study']
                        for s in ah:
                            if s in profile:
                                i = profile.index(s)
                                av = int(profile[(i - 3):i])
                                break
                        if cv is not None and c.contact != NA:
                            assert c.contact == cv
                        else:
                            c.contact = cv
                        if av is not None and c.assign != NA:
                            assert c.assign == av
                        else:
                            c.assign = av
                        pr = details.find('div',
                                          { 'class' : 'courseProfileRequiredPrerequisites' })
                        if pr is not None:
                            for req in pr.find_all('li'):
                                specs = req.get_text()
                                if '(' and ')' in specs: # another course
                                    i = specs.index('(') + 1
                                    j = specs.index(')') 
                                    c.prereq.add(specs[i:j])
                                else:
                                    assert ',' not in specs
                                    assert '"' not in specs
                                    assert ';' not in specs                            
                                    c.prereq.add(specs)
                        ps = details.find('div', { 'id' : 'courseProfileCertificates' })
                        if ps is not None:
                            for prog in ps.find_all('li'):
                                pn = prog.get_text()
                                if ':' in pn:
                                    pn = pn.split(':')[0]
                                c.incl.add(pn.strip().lstrip())

with open('courses.csv', 'w') as target:
    print(header, file = target)
    for course in records:
        print(records[course], file = target)
