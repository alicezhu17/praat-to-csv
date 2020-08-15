from xlwt import Workbook 
import numpy as np
import os

nameoffile = "example.TEXTGRID"

inFile = open(nameoffile,'r')
filelines = []
for line in inFile:
    filelines.append(line.strip())
inFile.close()

tiernames = []
for i, line in enumerate(filelines):
    if 'name = ' in line:
        tiernames.append(line.strip('name = ').strip('""'))


def extract_tier(filelines,linename):
    '''Extract unprocessed tier from filelines with name linename
    '''
    start = None
    end = None
    for i, line in enumerate(filelines):
        if 'name = ' in line and linename in line:
            start = i
        elif start and 'name = ' in line:
            end = i
            break    
    tier = filelines[start:end]
    return tier

def strip_interval_tier(tier):
    '''Input: interval tier from extract_tier()
    Returns [(Word,starttime,endtime),(Word,starttime,endtime),(Word,starttime,endtime)...]
    '''
    wordtup = []
    for b, item in enumerate(tier):
        if 'xmin = ' in item and tier[b+2] != 'text = ""' and 'intervals: size = ' not in tier[b+2]:
            now = tier[b+2].strip('text = ').strip('""').strip()
            wordtup.append((now,float(tier[b].strip('xmin = ')),float(tier[b+1].strip('xmax = '))))
    return wordtup

def strip_point_tier(tier):
    '''Input: point tier from extract_tier()
    Returns [(time,landmark),(time,landmark),(time,landmark)...]
    '''
    lmtime = []
    for a, item in enumerate(tier):
        if 'number = ' in item:
            lmtime.append((float(item.strip('number = ')),tier[a+1].strip('mark = ').strip('""')))
    return lmtime

def strip_tiers(filelines,tiernames):
    '''
    Uses extract_tier(), strip_interval_tier(), strip_point_tier()
    Returns list of stripped tiers
    '''
    tierslist = []
    wordtier = extract_tier(filelines,tiernames[0])
    wordtup = strip_interval_tier(wordtier)
    tierslist.append(wordtup)
    
    for elem in tiernames[1:]:
        lmtier = extract_tier(filelines,elem)
        lmtime = strip_point_tier(lmtier)
        tierslist.append(lmtime)
    return tierslist


def word_dic(tierslist,tiernames):
    '''
    Returns dictionary keyed by (Word,start,end) tuple
    Values are dictionary keyed by point tiers value list of (time,lm) tuples
    {(Word,start,end):{Tier1: [(time,lm),(time,lm),(time,lm)], Tier2: [(time,lm),(time,lm),(time,lm)]},etc}
    '''
    worddic = {elem:{k:[] for k in tiernames[1:]} for elem in tierslist[0]}
    for i, tier in enumerate(tierslist[1:]):
        name = tiernames[i+1]
        for lm in tier:
            for word in worddic:
                if lm[0]>=word[1] and lm[0]<=word[2]:
                    worddic[word][name].append(lm)
                    break
    return worddic
    
    
tierslist = strip_tiers(filelines,tiernames)
worddic = word_dic(tierslist,tiernames)    
print(worddic, '\n')    



    
wb = Workbook()
sheet1 = wb.add_sheet('Textgrid')
numtiers = len(tiernames)

i = 0
for wordi, elem in enumerate(worddic):
    sheet1.write(numtiers*wordi*2,0,tiernames[0])
    j = 1
    for tup in elem:
        sheet1.write(numtiers*wordi*2,j,tup)
        j += 1
        
    for namei,name in enumerate(tiernames):
        if namei != 0:
            sheet1.write(numtiers*wordi*2+2*namei-1,0,name)
            sheet1.write(numtiers*wordi*2+2*namei,0,'Time')
            j = 1
            for lmtime in worddic[elem][name]:
                sheet1.write(numtiers*wordi*2+2*namei-1,j,lmtime[1])
                sheet1.write(numtiers*wordi*2+2*namei,j,lmtime[0])
                j += 1











wb.save('example.xls')