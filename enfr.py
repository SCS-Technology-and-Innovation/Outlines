from bs4 import BeautifulSoup
from random import randint
from time import sleep
from os import path, remove
import requests

hdr = { 'User-Agent' :
        'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:97.0) Gecko/20100101 Firefox/97.0' }

domain = "https://www.mcgill.ca"
aos = { 'en': domain + "/continuingstudies/areas-study",
        'fr': domain + "/continuingstudies/fr/domaines-detudes" }
loc = {'en': 'aos.html', 'fr': 'aosf.html' }

language = { 'en': set(), 'fr': set() }
mapping = dict()
areaname = dict()
programname = dict()

for af in aos:
    if not path.exists(loc[af]):        
        r = requests.get(aos[af], hdr)
        sleep(randint(1, 5))
        with open(loc[af], 'w') as target:
            print(' '.join(r.text.split()), file = target)
    with open(loc[af]) as source:
        soup = BeautifulSoup(source.read(), 'html.parser')
        content = soup.find('div', { 'class' : 'grid' })
        areas = content.find_all('div', { 'class' : 'item areaofstudy' })
        names = content.find_all('div', { 'class' : 'title' })
        for (a, n) in zip(areas, names):
            if n is None:
                continue
            name = n.string.strip()
            aurl = domain + a.find('a')['href']
            label = aurl.split('/')[-1]
            fl = label + '.html'
            if not path.exists(fl):        
                r = requests.get(aurl, hdr)
                sleep(randint(1, 5))        
                with open(fl, 'w') as target:
                    print(' '.join(r.text.split()), file = target)
            with open(fl) as src:
                aso = BeautifulSoup(src.read(), 'html.parser')
                for ac in aso.find_all('div', { 'class' : 'items' }):
                    for program in ac.find_all('div', { 'class' : 'title' }):
                        pd = program.find('a')
                        if pd is None:
                            continue
                        purl = domain + pd['href']
                        if pd is None:
                            continue
                        pn = pd.string
                        if pn is None:
                            continue
                        programname[purl] = pn.strip()
                        areaname[purl] = name
                        if af == 'en':
                            assert purl is not None
                            language[af].add(purl)
                        pl = purl.split('/')[-1] + '.html'
                        if not path.exists(pl):        
                            pr = requests.get(purl, hdr)
                            sleep(randint(1, 5))        
                            with open(pl, 'w') as target:
                                print(' '.join(pr.text.split()), file = target)
                        with open(pl) as src:
                            pso = BeautifulSoup(src.read(), 'html.parser')
                            alt = pso.find('div', { 'id' : 'accessibility' })
                            try:
                                other = domain + alt.find('a')['href']
                                mapping[purl] = other
                            except:
                                pass # unmapped

areamap = dict()
NA = '(translation unavailable)'
info = set()

for l in language:
    for p in language[l]:
        if p not in mapping:
            continue
        o = mapping[p]
        f = [ programname.get(p, NA), programname.get(o, NA) ]
        if l == 'fr':
            f = f[::-1] # reverse
        if p in areaname and o in areaname:
            if l == 'fr':
                areamap[areaname[o]] = areaname[p]
            else:
                areamap[areaname[p]] = areaname[o]
        info.add(','.join(f))

for ae in areamap:
    info.add(f'{ae},{areamap[ae]}')

with open('dictionary.csv', 'w') as target:
    for i in sorted(list(info)):
        print(i, file = target)
            
            
            
                      
