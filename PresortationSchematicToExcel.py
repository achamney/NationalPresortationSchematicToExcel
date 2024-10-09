#!/usr/bin/env python3

import sys
import openpyxl
import os
from pypdf import PdfReader

# magic numbers for the character numbers between each column. Eg Level1 stations are between char 4 and 56.
levelRanges = [4,56,87,118,150]
chars = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
postCodeLtoProvCode = {'A':1, 
                       'B':2,
                       'C':3,
                       'E':4,
                       'G':5,'H':5,'J':5,
                       'K':6,'L':6,'M':6,'N':6,'P':6,
                       'R':7,
                       'S':8,
                       'T':9,
                       'V':10}
postCodes = []
for i in range(0,26):
    for j in range(0,10):
        for k in range(0,26):
            postCodes.append(f"{chars[i]}{j}{chars[k]}")

def getLevelText(rowText, lvl):
    return rowText[levelRanges[lvl]:levelRanges[lvl+1]].replace("CONT'D./SUITE","")
def getProvCode(l4, postCode):
    l4Strip = sanitize(l4)
    try:
        if len(l4Strip) >0:
            return postCodeLtoProvCode[l4Strip[0:1]]
        else :
            return postCodeLtoProvCode[postCode[0:1]]
    except Exception as e:
        print(l4)
        print(postCode)
        return 0
    
def sanitize(inText):
    return inText.replace("-|","").replace("|","").replace("  "," ").replace("-\\","")

# Keep a running list of the stations at L1,2,3,4. 
#If there are many "----" or no "|" at the begining of the column, then there is no station for this FSA
def populateCurLevels(curLevelName, curLineIndex, allLines):
    curLine = allLines[curLineIndex]
    outCurLevelNames = [curLevelName[0],curLevelName[1],curLevelName[2],curLevelName[3]]
    for lvl in range(0,4):
        lvlName = getLevelText(curLine, lvl)
        
        #Test to see if there is any station name for this FSA, or if it should be removed
        begLvlName = lvlName[0:2] # The first few characters indicate whether the station is repeated from above or if there's a new station name 
        if (lvl > 0 or "----" in lvlName) and "\\" not in begLvlName and "|" not in begLvlName and "/" not in begLvlName:
            outCurLevelNames[lvl] = ""
        if lvl == 0 and "-\   " in lvlName:
            outCurLevelNames[0] = ""

        lvlName = sanitize(lvlName).strip()
        if len(lvlName) > 2 and "--" not in lvlName: # New station name found. Set the outCurLevelName for this lvl
            nextLevel = getLevelText(f"{allLines[curLineIndex+1]}",lvl).strip()
            if lvl == 0:
                outCurLevelNames[lvl] = lvlName
            elif len(nextLevel)>2:
                #Station names are typically on two lines, so we have to look forward to grab the full station name
                outCurLevelNames[lvl] = f"{lvlName} {sanitize(nextLevel)}".replace("  "," ")
    return outCurLevelNames

def run(filename=None):
    if len(sys.argv) < 2 and filename is None:
        print("""Usage:
    PresortationSchematicToExcel.py /path/to/pdf
Returns:
    A file called fsa.xlsx with all FSAs and their corresponding stations""")
        return
    parsedCodes = []
    curLevelName = ["","","",""]
    uniqueStations = []
    for i in range(0,11):
        uniqueStations.append([[""],[""],[""],[""]])
    reader = PdfReader(sys.argv[1] if filename is None else filename)
    number_of_pages = len(reader.pages)
    
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["FSA","Level1","Level2","Level3","Level4","Code_Level1","Code_Level2","Code_Level3","Code_Level4","Code_Prov","FullCodePC"])
    for i in range(0,number_of_pages):
        page = reader.pages[i]
        text = page.extract_text()
        lines = text.split("\n")
        for j in range(0,len(lines)):
            thisLine = lines[j]
            postCode = thisLine[1:4] # 1:4 because the first character is always a space

            if postCode in postCodes and postCode not in parsedCodes:
                curLevelName = populateCurLevels(curLevelName, j, lines)
                    
                provCode = getProvCode(curLevelName[3],postCode)
                for lvl in range(0,4):
                    if curLevelName[lvl] not in uniqueStations[provCode][lvl]:
                        uniqueStations[provCode][lvl].append(curLevelName[lvl])

                l1=uniqueStations[provCode][0].index(curLevelName[0])
                l2=uniqueStations[provCode][1].index(curLevelName[1])
                l3=uniqueStations[provCode][2].index(curLevelName[2])
                l4=uniqueStations[provCode][3].index(curLevelName[3]) 
                sheet.append([postCode,curLevelName[0],
                            curLevelName[1],
                            curLevelName[2],
                            curLevelName[3],
                            l1,l2,l3,l4,
                            provCode,
                            provCode + l4/100+l3/10000+l2/1000000+l1/1000000000])
                parsedCodes.append(postCode)

    workbook.save("fsa.xlsx")
    print(len(parsedCodes))
    return f"{os.getcwd()}{os.sep}fsa.xlsx"

if __name__ == '__main__':
    print(run())  