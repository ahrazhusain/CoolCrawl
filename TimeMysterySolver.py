
# I am Awesome
from xml.dom import minidom
import glob
import xml.etree.ElementTree as ET
import sys
import xlwt
import fnmatch
import os
from datetime import datetime
import time
import pytz
import calendar

utc=pytz.utc
eastern=pytz.timezone('US/Eastern')
fmt='%Y-%m-%d %H:%M:%S'

oFiles = 0

#XML sheet creation
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")


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


#Function to find time overlapping

def intersection(x1, x2, y1, y2):
    
    status=[False,0,0]
    global oFiles
    intervals = [[x1,x2],[y1,y2]]
    #print "Comparing intervals: ", intervals
    
    overlapping = [ [x,y] for x in intervals for y in intervals if x is not y and x[1]>y[0] and x[0]<y[0] ]
    for x in overlapping:
        #print '{0} overlaps with {1}'.format(x[0],x[1])
        status[1]=(max(x1,y1))
        status[2]=(min(x2,y2))
        #print "--> Overlap time: ",status[1], " to ", status[2]
        oFiles=oFiles+1
        status[0]= True
    #print "--> No Overlap! "
    return status

#Function to get start and end tracker time: iTrace

def itraceTrackerTimes(xmldoc):
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


#Function to get start and end system time: iTrace

def itraceSystemTimes(xmldoc):
    iTimeS=[0,0]
    itemlist = xmldoc.getElementsByTagName('response')
    t=0
    for s in itemlist :
        try:
            if t==0:
                iStartS = int(s.attributes['system-time'].value)/1000
                #print "\niTrace start: ", iStart
                t=1
            iEndS = int(s.attributes['system-time'].value)/1000
        
        except KeyError:
            #print "Keyerror! "
            pass
    #print "iTrace end: ", iEnd

    iTimeS[0]=iStartS
    iTimeS[1]=iEndS
    
    return iTimeS




#Function to get start and end time: Mylyn

def mylynTimes(root):
    mTime=[0,0]
    t=0
    for k in root.findall('InteractionEvent'):
        if t==0:
            mStart = k.get('StartDate')
            #print "Mylyn start :",mStart, "| Epoch: ",timeToEpoch(mStart)
            t=1
        mEnd = k.get('EndDate')
    #print "Mylyn End :",mEnd, "| Epoch: ",timeToEpoch(mEnd)
    
    mTime[0]=mStart
    mTime[1]=mEnd
    
    return mTime


#Display comparisions in timings for iTrace amd corresponding Mylyn files
def compareTimes(task,iTimeT, mTime):
    #print "\n--------------------- TASK",task," --------------------------------"
    #print "\niTrace start: ", epochToTime(iTimeT[0]), "| Epoch: ",iTimeT[0]
    #print "Mylyn start:  ",timeformat(mTime[0]).strftime(fmt), "| Epoch: ",timeToEpoch(mTime[0])
    #print "iTrace end:   ", epochToTime(iTimeT[1]), "| Epoch: ",iTimeT[1]
    #print "Mylyn End :   ",timeformat(mTime[1]).strftime(fmt), "| Epoch: ",timeToEpoch(mTime[1])
    status= intersection(iTimeT[0], iTimeT[1], timeToEpoch(mTime[0]), timeToEpoch(mTime[1]))
    #print "--> Overlap: ", status
    #print "total Overlapping files: ", oFiles
    #print "--------------------------------------------------------------\n"
    
    return status

#-------------------- MAIN -----------------------



print "\nPID      Task #        i/Start Time           i/End Time               m/Start Time            m/End Time\n____________________________________________________________________________________________________________"
sheet1.write(0, 0, "PID")
sheet1.write(0, 1, "Task #")
sheet1.write(0, 2, "Mylyn Start Time")
sheet1.write(0, 3, "Mylyn End Time")
sheet1.write(0, 4, "iTrace Start Time(Tracker)")
sheet1.write(0, 5, "iTrace End Time(Tracker)")
sheet1.write(0, 6, "TrackerTime Overlap")
sheet1.write(0, 7, "iTrace Start Time(System)")
sheet1.write(0, 8, "iTrace End Time(System)")
sheet1.write(0, 9, "SystemTime Overlap")
sheet1.write(0, 10, "iTrace File")
sheet1.write(0, 11, "Mylyn File")


n=1;

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


    typesKinds=[]
    howsKinds=[]
    allKinds=[]
    
    i=0;
    m=0;
    
    
    
    
    #print "\nTotal Found ",len(FilesFound), " iTrace files...\n"
    
    
    for f in FilesFound:
        tree=0
        #print "Parsing file", n," : ",f
        xmldoc = minidom.parse(f)
        
        if xmldoc.getElementsByTagName('itrace-records'):
            #print "itrace file: ",f
            i=i+1;
            iTimeT=itraceTrackerTimes(xmldoc)
            iTimeS=itraceSystemTimes(xmldoc)
            #print "iTrace:", iTimeT
            mF="local-"+str(f[-31])
            mP=os.path.dirname(os.path.abspath(f))
            #print "Corresponding mylyn file: ",mP,mF
            for f1 in mFilesFound:
                #if mP.lower() in f1.lower():
                if mF.lower() in f1.lower():
                    if "._" not in f1.lower():
                        #print "Found Mylyn file: ", f1
                        
                        xmldoc1 = minidom.parse(f1)
                        
                        if xmldoc1.getElementsByTagName('InteractionHistory'):
                            #print "in Mylyn file: ",f1
                            m=m+1;
                            tree = ET.parse(f1)
                            root = tree.getroot()
                            mTime=mylynTimes(root)
                            statusT=compareTimes(str(f[-31]),iTimeT, mTime)
                            statusS=compareTimes(str(f[-31]),iTimeS, mTime)
                            print "ID",int(filter(str.isdigit, os.path.basename(os.path.dirname(f)))),"    ",str(f[-31]),"    ",epochToTime(iTimeT[0]),"    " ,epochToTime(iTimeT[1]),"    ",epochToTime(timeToEpoch(mTime[0])),"    " ,epochToTime(timeToEpoch(mTime[1]))
                          
                            
                            
                            sheet1.write(n, 0, int(filter(str.isdigit, os.path.basename(os.path.dirname(f)))))
                            sheet1.write(n, 1, str(f[-31]))
                            sheet1.write(n, 2, epochToTime(timeToEpoch(mTime[0])))
                            sheet1.write(n, 3, epochToTime(timeToEpoch(mTime[1])))
                            sheet1.write(n, 4, epochToTime(iTimeT[0]))
                            sheet1.write(n, 5, epochToTime(iTimeT[1]))
                            sheet1.write(n, 6, statusT[2]-statusT[1])
                            sheet1.write(n, 7, epochToTime(iTimeS[0]))
                            sheet1.write(n, 8, epochToTime(iTimeS[1]))
                            sheet1.write(n, 9, statusS[2]-statusS[1])
                            sheet1.write(n, 10, str(f).replace(os.getcwd(),""))
                            sheet1.write(n, 11, str(f1).replace(os.getcwd(),""))
            
                            
                            
                            
                            n=n+1




#else:
#print "Cannot verify if mylyn file!"
#else:
#print "Path issues"


#else:
#print "Useless file: ",f






#print "\n\nTOTAL itrace files : ",i
#print "\nTOTAL mylyn files : ",m



book.save("timeMysterySolver.xls")

print "--> Total Overlapping files: ", oFiles
