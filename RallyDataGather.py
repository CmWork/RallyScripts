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
start date
end date

TABLES
iterSummary {
    OID, IterName, Stories, IncompleteStories, IncompleteStoryPercent, Points, IncompletePoints, IncompletePointPercent, StartDate, EndDate
}

userCapacity {
    IterID, UserName, UserID, Capacity, TotalActual, TotalEstimate, AERatio, AENorm, ACRatio, ECRatio
}

userTasks {
    IterID, UserID, TaskEst, Actual, TaskName
}
'''

class RallyData:
    def __init__(self, dbConn, dbCur):
        self.errors = ''
        self.conn = dbConn
        self.cursor = dbCur
        self.rally = self.rallyLogin()
        self.iterDict = dict()

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

        print devSyncDict
        self.getTasks(devSyncDict)

    def getTasks(self, devSyncDict):
        for key in devSyncDict.keys():
            query = 'Owner.UserName = "' + key + '@spirent.com"' 
            if devSyncDict[key] is not None:
                query = query + ' AND LastUpdateDate > "' + devSyncDict[key] + '"'

            response = self.rally.get('Task', fetch='Project,WorkProduct,LastUpdateDate,Name,Owner,Iteration,Actuals,Estimate,State', project=None, query=query)
            if response.errors:
                sys.stdout.write("\n".join(errors))
                sys.exit(1)

            for task in response:
                print task.Owner.Name
                print task.Owner.ObjectID
                print task.LastUpdateDate
                print task.Project.Name
                print task.WorkProduct.Name
                if task.WorkProduct.ObjectID not in self.iterDict:
                    self.iterDict[task.WorkProduct.ObjectID] = task.WorkProduct

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
        self.cursor.execute("create table if not exists iterSummary(OID integer UNIQUE, IterName text, StartDate text, EndDate text, Stories integer, IncompleteStories integer, IncompleteStoryPercent real, Points integer, IncompletePoints integer, IncompletePointPercent real)")
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