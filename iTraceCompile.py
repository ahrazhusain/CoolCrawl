
# I am Awesome
from xml.dom import minidom
import glob
import xml.etree.ElementTree as ET
import sys
import xlsxwriter
import xlwt
import fnmatch
import os
from datetime import datetime
import time
import pytz
import calendar
import csv

utc=pytz.utc
eastern=pytz.timezone('US/Eastern')
fmt='%Y-%m-%d %H:%M:%S'

oFiles = 0

#CSV master sheet creation
csvfile =  open('iTraceXls/master.csv', 'w')
fieldnames = ["File", "Type", "x", "y", "left-validation", "right-validation", "left-pupil-diameter", "Fixation", "right-pupil-diameter", "tracker-time", "system-time", "nano-time", "line_base_x", "Line", "Col", "SCE-Count", "hows", "types", "fullyQualifiedNames", "line_base_y", "Difficulty", "Expertise", "PID", "TaskID"]
writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
writer.writeheader()


#get all subdirectories
def SubDirPath (d):
    return filter(os.path.isdir, [os.path.join(d,f) for f in os.listdir(d)])


#Function to convert to GMT
def timeformat(dS):
    tZone=dS[-3:]
    #print "Zone:", tZone
    dS=dS.rpartition(".")[0]
    try:
        dt = datetime.strptime(dS, fmt)
    #print  dt
    except ValueError:
        #print "Cannot strip time: ", dS
        pass
    
    date_eastern=eastern.localize(dt,is_dst=None)
    date_utc=date_eastern.astimezone(utc)

    return date_utc

#Function to convert time to epoch
def timeToEpoch(dT):
    
    dT=timeformat(dT)
    #print dT
    epochTime=calendar.timegm(dT.timetuple())
    
    
    return epochTime


#Function to convert epoch to system time
def epochToTime(eT):
    
    systemTime=time.strftime(fmt,  time.gmtime(eT))
    
    
    return systemTime



#Function to get start and end tracker time: iTrace

def itraceData(xmldoc):
    iTimeT=[0,0]
    itemlist = xmldoc.getElementsByTagName('response')
    t=0
    for s in itemlist :
        try:
            if t==0:
                iStartT = int(s.attributes['tracker-time'].value)/1000
                #print "\niTrace start: ", iStart
                t=1
            iEndT = int(s.attributes['tracker-time'].value)/1000
        
        except KeyError:
            #print "Keyerror! "
            pass
    #print "iTrace end: ", iEnd

    iTimeT[0]=iStartT
    iTimeT[1]=iEndT
    
    return iTimeT




#-------------------- MAIN -----------------------

print "\n\n\n"



for folder in SubDirPath(os.getcwd()):
    AllFound = []
    FilesFound = []
    mFilesFound = []
    #print "\n\n\n_____ $$$>>>> In folder: ", folder
    for root, dirnames, filenames in os.walk(folder):
        for filename in fnmatch.filter(filenames, '*.xml'):
            if ".metadata" not in root.lower():
                #print "metadata not in file", root
                FilesFound.append(os.path.join(root, filename))
            else:
                #print "metadata file", root
                continue



#print "Searching for iTrace Files "
    AllFound = FilesFound
    FilesFound = []
    for f in AllFound:
        #mF="local-"+str(sys.argv[1])
        itF="gaze-responses-task"
        if itF.lower() in f.lower():
            #print "Found iTrace file: ", f
            FilesFound.append(f)
        elif "local-".lower() in f.lower():
            #print "Found iTrace file: ", f
            mFilesFound.append(f)
    
    
    
    
    #print "\nTotal Found ",len(FilesFound), " iTrace files...\n"
    
    
    for f in FilesFound:
        tree=0
        print "Parsing file", f
        TaskId = str(f[-31])
        mP=os.path.dirname(os.path.abspath(f))
        tmp = f.replace(os.getcwd(),"")
        tmp = tmp.replace(os.path.basename(os.path.basename(f)),"")
        tmp = tmp.split('/')
        PID = int(filter(str.isdigit, tmp[1]))
        print "Task ",TaskId, "PID:", PID
        
        #Initialize sheet
        #response file="3.txt" type="text" x="877" y="942" left-validation="1.0" right-validation="1.0" left-pupil-diameter="2.432342529296875" fixation="false" right-pupil-diameter="2.3959503173828125" tracker-time="1409165309383" system-time="1409166878037" nano-time="2561621322266" line_base_x="425" line="6" col="50" hows="DECLARE;DECLARE" types="METHOD;TYPE" fullyQualifiedNames="net.sf.jabref.imports.BibtexParser.parseString();net.sf.jabref.imports.BibtexParser" line_base_y="930"/
        
        #XML sheet creation
        saveName = "iTraceXls/PID"+str(PID)+"-Task"+str(TaskId)+".xlsx"
        workbook = xlsxwriter.Workbook(saveName)
        sheet = workbook.add_worksheet()

        n=0
        sheet.write(n, 0, "File")
        sheet.write(n, 1, "Type")
        sheet.write(n, 2, "x")
        sheet.write(n, 3, "y")
        sheet.write(n, 4, "left-validation")
        sheet.write(n, 5, "right-validation")
        sheet.write(n, 6, "left-pupil-diameter")
        sheet.write(n, 7, "Fixation")
        sheet.write(n, 8, "right-pupil-diameter")
        sheet.write(n, 9, "tracker-time")
        sheet.write(n, 10, "system-time")
        sheet.write(n, 11, "nano-time")
        sheet.write(n, 12, "line_base_x")
        sheet.write(n, 13, "Line")
        sheet.write(n, 14, "Col")
        sheet.write(n, 15, "SCE-Count")
        sheet.write(n, 16, "hows")
        sheet.write(n, 17, "types")
        sheet.write(n, 18, "fullyQualifiedNames")
        sheet.write(n, 19, "line_base_y")
        sheet.write(n, 20, "Difficulty")
        sheet.write(n, 21, "Expertise")
        sheet.write(n, 22, "PID")
        sheet.write(n, 23, "TaskID")
        n+=1

        xmldoc = minidom.parse(f)
        
        if xmldoc.getElementsByTagName('itrace-records'):
            responsesList = xmldoc.getElementsByTagName('response')
            for r in responsesList:
                
                try:
                    
                    RF = (r.attributes['file'].value)
                    
                    Type = (r.attributes['type'].value)
                    
                    X = (r.attributes['x'].value)

                    Y = (r.attributes['y'].value)
                  
                    LV = (r.attributes['left-validation'].value)
                    
                    RV = (r.attributes['right-validation'].value)
                    
                    LPD = (r.attributes['left-pupil-diameter'].value)
                  
                    Fix = (r.attributes['fixation'].value)
                    
                    RPD = (r.attributes['right-pupil-diameter'].value)
              
                    TT = (r.attributes['tracker-time'].value)
                   
                    ST = (r.attributes['system-time'].value)
                    
                    NT = (r.attributes['nano-time'].value)
                    
                    LBX = (r.attributes['line_base_x'].value)
               
                    Line = (r.attributes['line'].value)
                    
                    COL = (r.attributes['col'].value)
                   
                    LBY = (r.attributes['line_base_y'].value)
                    HOWS = (r.attributes['hows'].value)
                    TYPES = (r.attributes['types'].value)
                    FQN = (r.attributes['fullyQualifiedNames'].value)
                
                    HOWS = HOWS.split(";")
                    TYPES = TYPES.split(";")
                    FQN = FQN.split(";")
                    [x.encode('UTF8') for x in HOWS]
                    [x.encode('UTF8') for x in TYPES]
                    [x.encode('UTF8') for x in FQN]
                    SCE = 0
                    #print "rf is :", RF

                        
                except KeyError:
                    #print "Keyerror! "
                    HOWS = [""]
                    TYPES = [""]
                    FQN = [""]
                    SCE = -1
                    
                    

            

                
                Difficulty = "Easy"
                if int(TaskId) == 2:
                    Difficulty = "Difficult"
                Expertise = "Professional"
                if int(PID) > 20:
                    Expertise = "Amateur"

                
                #print "HOWS is ", (HOWS)
                for i in range(len(HOWS)):
                    #print i, " SPLit: ", HOWS[i], " | ", TYPES[i], " | ", FQN[i]
                    
                    
                    

                    sheet.write(n, 0, RF)
                    sheet.write(n, 1, Type)
                    sheet.write(n, 2, X)
                    sheet.write(n, 3, Y)
                    sheet.write(n, 4, LV)
                    sheet.write(n, 5, RV)
                    sheet.write(n, 6, LPD)
                    sheet.write(n, 7, Fix)
                    sheet.write(n, 8, RPD)
                    sheet.write(n, 9, TT)
                    sheet.write(n, 10, ST)
                    sheet.write(n, 11, NT)
                    sheet.write(n, 12, LBX)
                    sheet.write(n, 13, Line)
                    sheet.write(n, 14, COL)
                    if SCE == -1:
                        SCE = 0
                        sheet.write(n, 15, SCE)
                    
                    else:
                        sheet.write(n, 15, i+1)
                        SCE = i+1
                    sheet.write(n, 16, HOWS[i])
                    sheet.write(n, 17, TYPES[i])
                    sheet.write(n, 18, FQN[i])
                    sheet.write(n, 19, LBY)
                    sheet.write(n, 20, Difficulty)
                    sheet.write(n, 21, Expertise)
                    sheet.write(n, 22, PID)
                    sheet.write(n, 23, TaskId)
                    n+=1

                    writer.writerow({"File": RF, "Type": Type, "x": X, "y": Y, "left-validation": LV, "right-validation": RV, "left-pupil-diameter": LPD, "Fixation": Fix, "right-pupil-diameter": RPD, "tracker-time": TT, "system-time": ST, "nano-time": NT, "line_base_x": LBX, "Line": Line, "Col": COL, "SCE-Count":SCE, "hows": HOWS[i], "types": TYPES[i], "fullyQualifiedNames": FQN[i], "line_base_y": LBY, "Difficulty": Difficulty, "Expertise": Expertise, "PID": PID, "TaskID": TaskId})
                    
        workbook.close()

workbook1.close()

