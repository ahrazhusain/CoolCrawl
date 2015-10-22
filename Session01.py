import xlrd
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
import csv

utc=pytz.utc
eastern=pytz.timezone('US/Eastern')
fmt='%Y-%m-%d %H:%M:%S'
fmt1='%Y-%m-%d %H:%M:%S.%f'
oFiles=0




#Function to convert epoch to system time
def epochToTime(eT):
    
    systemTime=time.strftime(fmt,  time.gmtime(eT))
    
    
    return systemTime

#Function to convert time to epoch
def timeToEpoch(dT):
    
    dT = datetime.strptime(dT, fmt)
    #print type(dT)
    epochTime=calendar.timegm(dT.timetuple())
    
    return epochTime


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
def newTimeToEpoch(dT):
    
    dT=timeformat(dT)
    #print dT
    epochTime=calendar.timegm(dT.timetuple())
    
    
    return epochTime


#Function to convert epoch to system time
def epochToTime(eT):
    
    systemTime=time.strftime(fmt,  time.gmtime(eT))
    
    
    return systemTime

#Function to find time overlapping

def intersection(intervals):
    
    status=[False,0,0,"0:00:00"]
    global oFiles
    intervals = [[intervals[0],intervals[1]],[intervals[2],intervals[3]]]
    #print "Comparing intervals: ", intervals
    
    overlapping = [ [x,y] for x in intervals for y in intervals if x is not y and x[1]>y[0] and x[0]<y[0] ]
    for x in overlapping:
        #print '{0} overlaps with {1}'.format(x[0],x[1])
        status[1]=epochToTime(max(intervals[0][0],intervals[1][0]))
        status[2]=epochToTime(min(intervals[0][1],intervals[1][1]))
        status[3]=datetime.strptime(status[2], fmt) - datetime.strptime(status[1], fmt)
        #print "--> Overlap time: ",status[1], " to ", status[2]
        #print "--> Overlap duration: ",status[3]
        oFiles=oFiles+1
        status[0]= True
    #print "--> No Overlap! "
    return status

#get all subdirectories
def SubDirPath (d):
    return filter(os.path.isdir, [os.path.join(d,f) for f in os.listdir(d)])


#Excited or Not and info
def PupilStudy (meanT, v):
    status = [False, 0, False ]
    if (v > 0):
        status[0] = True
    
    status[1]= PupilDegree(meanT, v)
    if (status[1]>1.0 ):
        status[2]=True

    
    return status #return status for info

#Linear Degree
def PupilDegree (meanT, v):
    
    status= (v/meanT)*10
    
    return status-10

#Calculate mean pupil for a task
def PupilMean (itemlist):
    
    meanPupil = 0.00
    t=0
    for s in itemlist :
        
        meanPupil+=(float(s.attributes['left-pupil-diameter'].value)+float(s.attributes['right-pupil-diameter'].value))/2.0
        t+=1
    
    meanPupilT=meanPupil/float(t)
    
    return meanPupilT



#Extract correspnding overlaps for actions and illusions
def sessionInfoGen(xmldoc,root,sheet1):
    #impressionInfo=[gazeTime, actionTime, gazeInstances, actionInstances,0]
    itemlist = xmldoc.getElementsByTagName('response')
    eventsT=0
    t=0
    wasAction=0
    n=1

    meanPupilT = PupilMean (itemlist)
    meanPupil=0.00
    print "Mean is: ", meanPupilT
    
    timeBackup = int(itemlist[0].attributes['tracker-time'].value)/1000
    #print "Found at", timeFound
    
    for s in itemlist :
        
        
        timeFound = int(s.attributes['tracker-time'].value)/1000
        #print "Found at", timeFound
        
        if(timeFound==timeBackup):
        
            meanPupil+=(float(s.attributes['left-pupil-diameter'].value)+float(s.attributes['right-pupil-diameter'].value))/2.0
            t=t+1
            eventsT=0
        else:
            Pdegree=round(PupilDegree(meanPupilT, float(meanPupil/t)),2)
            #print "Impression at time: ", timeBackup, " of degree ", Pdegree
            timeBackup=timeFound
            meanPupil=(float(s.attributes['left-pupil-diameter'].value)+float(s.attributes['right-pupil-diameter'].value))/2.0
            t=1
            
            
            for k in root.findall('InteractionEvent'):
                
                mStart = k.get('StartDate')
                mEnd = k.get('EndDate')
                
                r=range(int(newTimeToEpoch(mStart)),int(newTimeToEpoch(mEnd)))
                #print "Mylyn Start :",newTimeToEpoch(mStart),"Mylyn End :",newTimeToEpoch(mEnd)
                if timeFound in r:
                    #print "events:", eventsT
                    if(eventsT==0):
                        #print "-------------------------------->>>>>> Action at time: ", timeFound
                        wasAction=1
                        eventsT=1
            
            print "Impression at time: ", timeBackup, " of degree ", Pdegree," was action: ", wasAction
            sheet1.write(n, 0, timeBackup)
            sheet1.write(n, 1, Pdegree)
            sheet1.write(n, 2, wasAction)
            n+=1
            wasAction=0

    return sheet1

#______________________________MAIN_________________________________________________





#Scanning all xml files
AllFound = []
iFilesFound = []
mFilesFound = []

for folder in SubDirPath(os.getcwd()):

    #print "\n\n\n_____ $$$>>>> In folder: ", folder
    for root, dirnames, filenames in os.walk(folder):
        for filename in fnmatch.filter(filenames, '*.xml'):
            if ".metadata" not in root.lower():
                #print "metadata not in file", root
                iFilesFound.append(os.path.join(root, filename))
            else:
                #print "metadata file", root
                continue



#print "Searching for iTrace Files "
    AllFound = iFilesFound
    iFilesFound = []
    for f in AllFound:
        #mF="local-"+str(sys.argv[1])
        itF="gaze-responses-task"
        if itF.lower() in f.lower():
            #print "Found iTrace file: ", f
            iFilesFound.append(f)
        elif "local-".lower() in f.lower():
            #print "Found iTrace file: ", f
            mFilesFound.append(f)





#open excel file and read
workbook = xlrd.open_workbook('TimeMappings.xls')
worksheet = workbook.sheet_by_name('Sheet 1')
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0

fileNo=1

while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    #print 'Row:', curr_row
    curr_cell = -1
    d=[]
    while curr_cell < num_cells:
        curr_cell += 1
        # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
        cell_type = worksheet.cell_type(curr_row, curr_cell)
        cell_value = worksheet.cell_value(curr_row, curr_cell)
        
        if (curr_cell>1 and curr_cell<6):
            d.append(int(timeToEpoch(str(cell_value))))
        elif (curr_cell>5 and curr_cell<8):
            d.append(str(cell_value))
        #print '	',curr_cell," ", cell_type, ':', cell_value
    status=intersection(d)

    if(status[0]==True):
        task=int(str(d[4][-31]))
        id=str(os.path.dirname(d[4]))
        id=int(filter(str.isdigit, id))
        print "Comparision",fileNo," ",d[4]," AND ",d[5]," FOR TASK",task,"for user",id,"\n"
        fileNo+=1
        toFind=str(d[4])
        #print toFind
        """j=0
        
        try:
            meanPupilT= meanPupilT/i
            #print "Mean pupil for task: ", meanPupilT

        except ZeroDivisionError:
            #print "Mean pupil for task: 0.00"
            pass

        impressionInfo = [0,0,0,0,0] #Gaze time, Action time, Gaze instances, Action Instances, total actions?
        """
        for iFile in iFilesFound:
            #print "-->",toFind, " in ", iFile,"\n"
            if toFind.lower() in iFile.lower():
                #print "Found itrace: ",iFile
                mP=str(d[5])
                #print "MP:", mP
                
                for mFile in mFilesFound:
                    #print "Comparing",mP.lower() ," and " , mFilesFound
                    if mP.lower() in mFile.lower():
                        #print "Found Mylyn file: ", mFile
                        xmldoc = minidom.parse(iFile)#Parse itrace
                        tree = ET.parse(mFile)#parse mylyn
                        root = tree.getroot()
                        #print type(root)


                        if xmldoc.getElementsByTagName('itrace-records'):
                            #print type(root)
                            
                            #XML sheet creation
                            book = xlwt.Workbook(encoding="utf-8")
                            sheet1 = book.add_sheet("Sheet 1")
                            
                            #Initializing sheet values
                            n=1

                            sheet1.write(0, 0, "Time(s)")
                            sheet1.write(0, 1, "Degree")
                            sheet1.write(0, 2, "Action?")
    
                            sheet1=sessionInfoGen(xmldoc,root,sheet1)
                            
                            bookName="session_01_Data/User"+str(id)+"-task"+str(task)+".xls"
                            book.save(bookName)


                                            



        #print "==>> File data:- Gaze time:",impressionInfo[0],"Action time:",impressionInfo[1]," Gaze instances:",impressionInfo[2],"Action instances:",impressionInfo[3],"Totoal Actions:",impressionInfo[4]

        j=1
        if(j==0):
            print "No match!"

        print "\n======================================================"






