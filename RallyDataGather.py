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
    OID integer UNIQUE, 
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
    UserID integer, 
    UserName text, 
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
    UserID integer, 
    TaskID integer UNIQUE, 
    TaskEst real, 
    Actual real, 
    TaskName text, 
    State text, 
    Project text, 
    UserStory text
}
'''

class RallyData:
    def __init__(self, dbConn, dbCur):
        self.errors = ''
        self.conn = dbConn
        self.cursor = dbCur
        self.rally = self.rallyLogin()

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

        iterDict = self.getTasks(devSyncDict)
        self.getIterations(iterDict)

    def getTasks(self, devSyncDict):
        iterDict = dict()
        for key in devSyncDict:
            print devSyncDict[key]
            query = 'Owner.UserName = "' + key + '@spirent.com"' 
            if devSyncDict[key] is not None:
                query = query + ' AND LastUpdateDate > "' + devSyncDict[key] + '"'
            query = query + ' AND Iteration != null'

            response = self.rally.get('Task', fetch='Project,WorkProduct,LastUpdateDate,Name,Owner,Iteration,Actuals,Estimate,State', project=None, query=query)
            if response.errors:
                sys.stdout.write("\n".join(errors))
                sys.exit(1)

            for task in response:
                if task.Iteration is not None:
                    iteration = task.Iteration
                    if iteration.FormattedID not in iterDict:
                        iterDict[iteration.FormattedID] = dict()
                        iterDict[iteration.FormattedID]['Project'] = task.Project.Name
                        iterDict[iteration.FormattedID]['Name'] = iteration.Name
                        iterDict[iteration.FormattedID]['StartDate'] = iteration.StartDate
                        iterDict[iteration.FormattedID]['EndDate'] = iteration.EndDate
                        iterDict[iteration.FormattedID]['Stories'] = 0
                        iterDict[iteration.FormattedID]['Points'] = 0
                        iterDict[iteration.FormattedID]['IncompleteStories'] = 0
                        iterDict[iteration.FormattedID]['IncompletePoints'] = 0

                    print iteration.FormattedID
                    print task.Owner.ObjectID
                    print task.FormattedID
                    print task.Estimate
                    print task.Actuals
                    print task.Name
                    print task.State
                    print task.Project.Name
                    print task.WorkProduct.Name
                    # Write to DB
                    # self.cursor.execute("INSERT OR REPLACE INTO userTasks \
                    #     (IterID, UserID, TaskID, Estimate, Actual, TaskName, State, Project, Artifact) \
                    #     values(?,?,?,?,?,?,?,?,?)", \
                    #     (iteration.ObjectID,task.Owner.ObjectID,task.ObjectID, \
                    #         task.Estimate,task.Actuals,task.Name,task.State, \
                    #         task.Project.Name,task.WorkProduct.Name))
                    # self.conn.commit()
        return iterDict

    def getIterations(self, iterDict):
        for it in iterDict:
            self.getArtifact('UserStory', iterDict[it], it)
            self.getArtifact('Defect', iterDict[it], it)

            # Write to DB
            iterInfo = iterDict[it]
            incStoriesPer = float(iterInfo['IncompleteStories'])/float(iterInfo['Stories']) * 100
            incPointsPer = float(iterInfo['IncompletePoints'])/float(iterInfo['Points']) * 100
            self.cursor.execute("INSERT OR REPLACE INTO iterSummary \
                (FormattedID, IterName, StartDate, EndDate, Stories, IncompleteStories, IncompleteStoryPercent, Points, IncompletePoints, IncompletePointPercent) \
                values(?,?,?,?,?,?,?,?,?,?)", \
                (it,iterInfo['Name'],iterInfo['StartDate'],iterInfo['EndDate'], \
                    iterInfo['Stories'],iterInfo['IncompleteStories'],incStoriesPer, \
                    iterInfo['Points'],iterInfo['IncompletePoints'],incPointsPer))
            self.conn.commit()

    # Get UserStories from iteration
    def getArtifact(self, artifactType, sumDict, iteration):
        query = 'Iteration.FormattedID = "' + str(iteration) +'"'
        response = self.rally.get(artifactType, fetch='FormattedID,Iteration,Name,ScheduleState,PlanEstimate,RevisionHistory', project=sumDict['Project'], query=query)
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
            if artifact.ScheduleState == 'Incomplete':
                # Use revision history to get info
                for rev in artifact.RevisionHistory.Revisions:
                    desc = rev.Description
                    m = re.findall('(?<=PLAN\sESTIMATE\schanged\sfrom)*\[(\d)', desc)
                    if len(m) > 0:
                        return int(m[0])
        return points

    def createTables(self):
        # lastSync should be the last completed iteration endDate
        self.cursor.execute("create table if not exists syncData(UserName integer UNIQUE, LastSync text)")
        self.cursor.execute("create table if not exists iterSummary(FormattedID integer UNIQUE, IterName text, StartDate text, EndDate text, Stories integer, IncompleteStories integer, IncompleteStoryPercent real, Points integer, IncompletePoints integer, IncompletePointPercent real)")
        self.cursor.execute("create table if not exists userCapacity(IterID integer, UserID integer, UserName text, Capacity integer, TotalActual real, TotalEstimate real, AERatio real, AENorm real, ACRatio real, ECRatio real, unique(IterID, UserID))")
        self.cursor.execute("create table if not exists userTasks(IterID integer, UserID integer, TaskID integer UNIQUE, TaskEst real, Actual real, TaskName text, State text, Project text, UserStory text)")
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
        devReportList = ["ben.yoshino","greg.kodama"]

    conn = sqlite3.connect('RallyData.db')
    cursor = conn.cursor()
    if devReportList is not None:
        print str(devReportList)
        rallyDict = RallyData(conn, cursor).gatherData(devReportList)
    else:
        print "ERROR: Devs need to be set."

if __name__ == '__main__':
    main(sys.argv[1:])