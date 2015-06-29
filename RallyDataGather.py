import sys, getopt
from pyral import Rally, rallySettings
import unicodedata
import re
import datetime
from datetime import date
import sqlite3
'''
INPUT
devReportList

TABLES
syncData {
    UserName integer UNIQUE, 
    LastSync text
}
iterSummary {
    IterID integer UNIQUE, 
    IterName text, 
    StartDate text, 
    EndDate text, 
    Stories integer, 
    IncompleteStories integer, 
    IncompleteStoryPercent real, 
    Points integer, 
    IncompletePoints integer, 
    IncompletePointPercent real
}
userCapacity {
    IterID integer, 
    UserID text, 
    Capacity integer, 
    TotalActual real, 
    TotalEstimate real, 
    AERatio real, 
    AENorm real, 
    ACRatio real, 
    ECRatio real, 
    unique(IterID, UserID)
}
userTasks {
    IterID integer, 
    UserID text, 
    TaskID integer UNIQUE, 
    Estimate real, 
    Actual real, 
    TaskName text, 
    State text, 
    Project text, 
    Artifact text
}
'''

class RallyData:
    def __init__(self, dbConn, dbCur):
        self.errors = ''
        self.conn = dbConn
        self.cursor = dbCur
        self.rally = self.rallyLogin()
        self.devList = None

    def rallyLogin(self):
        server = 'rally1.rallydev.com'
        user = 'caden.morikuni@spirent.com'
        password = 'sp1LSZA4UBNm'
        project = 'MBH - HNL'
        workspace = 'default'
        rally = Rally(server, user, password, workspace=workspace, project=project)
        rally.enableLogging('mypyral.log')
        return rally

    def gatherData(self, devList):
        self.devList = devList
        self.createTables()

        # Query SyncData table for last updates
        devSyncDict = dict()
        devQuery = None
        for dev in devList:
            devSyncDict[dev] = None
            if devQuery is None:
                devQuery = 'UserName="' + dev + '"'
            else:
                devQuery = devQuery + ' OR UserName="' + dev + '"'

        for row in self.cursor.execute("SELECT * FROM syncData WHERE " + devQuery):
            key = row[0]
            if key in devSyncDict.keys():
                devSyncDict[key] = row[1]
            else:
                print "ERROR: could not find " + key + " in developer list."

        for dev in devList:
                # Write to DB
                self.cursor.execute("INSERT OR REPLACE INTO syncData\
                    (UserName, LastSync) \
                    values(?,?)", \
                    (dev,str(date.today())))
                self.conn.commit()

        iterCapDict = self.getTasks(devSyncDict)
        self.getIterations(iterCapDict['Iterations'])
        self.getCapacities(iterCapDict['Capacities'])

    def getTasks(self, devSyncDict):
        iterCapDict = dict()
        iterDict = dict()
        capDict = dict()
        for key in devSyncDict:
            query = 'Owner.UserName = "' + key + '@spirent.com"' 
            if devSyncDict[key] is not None:
                query = query + ' AND LastUpdateDate > "' + devSyncDict[key] + '"'
            query = query + ' AND Iteration != null'

            if key not in capDict:
                capDict[key] = dict()
            userCapDict = capDict[key]

            response = self.rally.get('Task', fetch='Project,WorkProduct,LastUpdateDate,Name,Owner,Iteration,Actuals,Estimate,State', project=None, query=query)
            if response.errors:
                sys.stdout.write("\n".join(errors))
                sys.exit(1)

            for task in response:
                if task.Iteration is not None:
                    iteration = task.Iteration
                    iterId = iteration.ObjectID

                    if iterId not in iterDict:
                        iterDict[iterId] = dict()
                        iterDict[iterId]['UserCapacities'] = self.getUserIterCaps(iteration.UserIterationCapacities)
                        iterDict[iterId]['NumDevs'] = iterDict[iterId]['UserCapacities']['NumDevs']
                        iterDict[iterId]['Project'] = task.Project.Name
                        iterDict[iterId]['Name'] = iteration.Name
                        iterDict[iterId]['StartDate'] = iteration.StartDate
                        iterDict[iterId]['EndDate'] = iteration.EndDate
                        iterDict[iterId]['Stories'] = 0
                        iterDict[iterId]['Points'] = 0
                        iterDict[iterId]['IncompleteStories'] = 0
                        iterDict[iterId]['IncompletePoints'] = 0

                    if iterId not in userCapDict:
                        userCapDict[iterId] = dict()
                        userCapDict[iterId]['IterName'] = iteration.Name
                        userCapDict[iterId]['UserName'] = task.Owner.Name
                        userCapDict[iterId]['Capacity'] = iterDict[iterId]['UserCapacities'][key]
                        userCapDict[iterId]['TotalActuals'] = 0
                        userCapDict[iterId]['TotalEstimate'] = 0

                    # Modify Actual/Estimate based on None or 0
                    act = task.Actuals
                    est = task.Estimate
                    if est is None:
                        est = 0
                        act = 0
                    if act is None or act == 0:
                        act = est

                    if task.State == 'Completed':
                        userCapDict[iterId]['TotalActuals'] += act
                        userCapDict[iterId]['TotalEstimate'] += est

                    # Write to DB
                    self.cursor.execute("INSERT OR REPLACE INTO userTasks \
                        (IterID, UserID, IterName, UserName, TaskID, Estimate, Actual, TaskName, State, LastUpdateDate, Project, Artifact, ArtifactName) \
                        values(?,?,?,?,?,?,?,?,?,?,?,?,?)", \
                        (iteration.ObjectID,key,iteration.Name,task.Owner.Name,task.FormattedID, \
                            est,act,task.Name,task.State,task.LastUpdateDate, \
                            task.Project.Name,task.WorkProduct.FormattedID,task.WorkProduct.Name))
                    self.conn.commit()

        iterCapDict['Iterations'] = iterDict
        iterCapDict['Capacities'] = capDict
        return iterCapDict

    def getIterations(self, iterDict):
        for it in iterDict:
            self.getArtifact('UserStory', iterDict[it], it)
            self.getArtifact('Defect', iterDict[it], it)

            # Write to DB
            iterInfo = iterDict[it]
            if iterInfo['Stories'] > 0:
                incStoriesPer = float(iterInfo['IncompleteStories'])/float(iterInfo['Stories']) * 100
            else:
                incStoriesPer = 0
            if iterInfo['Points'] > 0:
                incPointsPer = float(iterInfo['IncompletePoints'])/float(iterInfo['Points']) * 100
            else:
                incPointsPer = 0
            self.cursor.execute("INSERT OR REPLACE INTO iterSummary \
                (IterID, IterName, StartDate, EndDate, NumDevs, Stories, IncompleteStories, IncompleteStoryPercent, Points, IncompletePoints, IncompletePointPercent) \
                values(?,?,?,?,?,?,?,?,?,?,?)", \
                (it,iterInfo['Name'],iterInfo['StartDate'],iterInfo['EndDate'],iterInfo['NumDevs'], \
                    iterInfo['Stories'],iterInfo['IncompleteStories'],incStoriesPer, \
                    iterInfo['Points'],iterInfo['IncompletePoints'],incPointsPer))
            self.conn.commit()

    def getCapacities(self, capDict):
        for capKey in capDict:
            for iterKey in capDict[capKey]:
                iterName = capDict[capKey][iterKey]['IterName']
                userName = capDict[capKey][iterKey]['UserName']
                cap = capDict[capKey][iterKey]['Capacity']
                t_act = capDict[capKey][iterKey]['TotalActuals']
                t_est = capDict[capKey][iterKey]['TotalEstimate']
                ave = 0
                avc = 0
                evc = 0
                if t_est != 0:
                    ave = t_act/t_est * 100
                else:
                    ave = 0
                ave_norm = 100-ave
                if cap is None:
                    print 'Skipping ' + iterName + " -> " + capKey + ': capacity is None'
                    continue
                else:
                    avc = t_act/cap * 100
                    evc = t_est/cap * 100

                # Write to DB
                self.cursor.execute("INSERT OR REPLACE INTO userCapacity \
                    (IterID, UserID, UserName, Capacity, TotalActual, TotalEstimate, AERatio, AENorm, ACRatio, ECRatio) \
                    values(?,?,?,?,?,?,?,?,?,?)", \
                    (iterKey,capKey,userName,cap,t_act,t_est,ave,ave_norm,avc,evc))
                self.conn.commit()

    # Get UserStories from iteration
    def getArtifact(self, artifactType, sumDict, iteration):
        query = 'Iteration.ObjectID = "' + str(iteration) +'"'
        response = self.rally.get(artifactType, fetch='ObjectID,Iteration,Name,ScheduleState,PlanEstimate,RevisionHistory', project=sumDict['Project'], query=query)
        if response.errors:
            sys.stdout.write("\n".join(errors))
            sys.exit(1)

        for artifact in response:
            # Gather Summary Info
            if artifactType == 'UserStory' and artifact.PlanEstimate is not None:
                points = int(artifact.PlanEstimate)
                if artifact.ScheduleState == 'Incomplete':
                    points = self.getEstChange(artifact)
                    sumDict['IncompleteStories'] += 1
                    sumDict['IncompletePoints'] += points
                if points > 0:
                    sumDict['Stories'] += 1
                    sumDict['Points'] += points

    def getEstChange(self, artifact):
        points = 0
        if artifact.PlanEstimate is not None:
            points = artifact.PlanEstimate
        if artifact.PlanEstimate == 0:
            if artifact.ScheduleState == 'Incomplete':
                # Use revision history to get info
                for rev in artifact.RevisionHistory.Revisions:
                    desc = rev.Description
                    m = re.findall('(?:PLAN\sESTIMATE\schanged\sfrom\s\[)(\d)', desc)
                    if len(m) > 0:
                        return int(m[0])
        return points

    def getUserIterCaps(self, uicList):
        numDevs = 0
        uicDict = dict()
        for uic in uicList:
            try:
                uname = uic.User.UserName.split('@')[0]
                
                # Count number of devs participating in this sprint
                if uic.Capacity > 0:
                    numDevs += 1

                # Ignore dev if not queried for
                if uname not in self.devList:
                    continue
                uicDict[uname] = uic.Capacity
            except:
                print "ERROR: UIC issue in " + uic.Iteration.Project.Name + ":" + uic.Iteration.Name
                continue
        uicDict['NumDevs'] = numDevs
        return uicDict

    def createTables(self):
        # lastSync should be the last completed iteration endDate
        self.cursor.execute("create table if not exists syncData(UserName integer UNIQUE, LastSync text)")
        self.cursor.execute("create table if not exists iterSummary(IterID integer UNIQUE, IterName text, StartDate text, EndDate text, NumDevs integer, Stories integer, IncompleteStories integer, IncompleteStoryPercent real, Points integer, IncompletePoints integer, IncompletePointPercent real)")
        self.cursor.execute("create table if not exists userCapacity(IterID integer, UserID text, UserName text, Capacity integer, TotalActual real, TotalEstimate real, AERatio real, AENorm real, ACRatio real, ECRatio real, unique(IterID, UserID))")
        self.cursor.execute("create table if not exists userTasks(IterID integer, UserID text, IterName text, UserName text, TaskID integer UNIQUE, Estimate real, Actual real, TaskName text, State text, LastUpdateDate text, Project text, Artifact text, ArtifactName text)")
        self.conn.commit()

def main(argv):
    startDate = None
    endDate = None
    devReportList = None
    try:
        opts, args = getopt.getopt(argv, "hd:", ["devs="])
    except getopt.GetoptError:
        print 'python RallyDataGather.py -d <dev list>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == "-h":
            print 'python RallyDataGather.py -d <dev list>'
            sys.exit()
        elif opt in ("-d", "--devs"):
            devReportList = arg.split(',')

    # DEBUGGING
    if devReportList is None:
        devReportList = ["ben.yoshino","greg.kodama","greg.ofiesh","daniel.dubois","gayan.abeysundara","brandon.tom","felma.duque","kelli.sawai","kent.kanja","scott.shimokawa","rance.yamamoto"]

    conn = sqlite3.connect('RallyData.db')
    cursor = conn.cursor()
    if devReportList is not None:
        print str(devReportList)
        rallyDict = RallyData(conn, cursor).gatherData(devReportList)
    else:
        print "ERROR: Devs need to be set."

if __name__ == '__main__':
    main(sys.argv[1:])