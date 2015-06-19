import sys, getopt
from pyral import Rally, rallySettings
import unicodedata
import re
import xlwt
from time import localtime, strftime
import sqlite3

# RunCmd: python <script> --conf=<cfg no ".cfg" extension>
quarters = {'Q1':['01-01', '03-31'], 'Q2':['04-01', '06-30'], 'Q3':['07-01', '09-30'], 'Q4':['10-01', '12-31']}

class RallyStories:
    def __init__(self):
        self.errors = ''
        self.db = sqlite3.connect('rally.db')
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

    # Get UserStories from iteration
    def getArtifact(self, artifactType, projName, iteration):
        query = 'Iteration.Name = "' + iteration +'"'
        response = self.rally.get(artifactType, fetch='FormattedID,Iteration,Name,ScheduleState,PlanEstimate,RevisionHistory', project=projName, query=query)
        if response.errors:
            sys.stdout.write("\n".join(errors))
            sys.exit(1)

        artList = list()
        for artifact in response:
            artDict = dict()
            artDict['Name'] = artifact.Name
            artDict['ID'] = artifact.FormattedID
            artDict['State'] = artifact.ScheduleState
            artDict['Est'] = artifact.PlanEstimate
            artList.append(artDict)

            # Gather Summary Info
            if artifactType == 'UserStory' and artDict['Est'] is not None:
                points = int(artDict['Est'])
                if artDict['State'] == 'Incomplete':
                    points = self.getEstChange(artifact)
                    self.summary['IncompleteStories'] += 1
                    self.summary['IncompletePoints'] += points
                self.summary['Stories'] += 1
                self.summary['Points'] += points
        return artList

    def getSprintStories(self, projectName, sprintName):
        sprintDict = dict()
        sprintDict["story"] = self.getArtifact('UserStory', projectName, sprintName)
        sprintDict["defect"] = self.getArtifact('Defect', projectName, sprintName)
        return sprintDict

    # Get project user stories
    def sprintStories(self, projectName, sprintName):
        self.summary = dict()
        self.summary['Stories'] = 0
        self.summary['Points'] = 0
        self.summary['IncompleteStories'] = 0
        self.summary['IncompletePoints'] = 0

        self.sprint = dict()
        self.sprint['Summary'] = self.summary
        self.sprint['Stories'] = self.getSprintStories(projectName, sprintName)

        return self.sprint

# cipHealthDict[projName]
#   iterationDict[iterName]
#       iterDataDict[stories] = # of stories
#       iterDataDict[points] = # of points
#       iterDataDict[defects] = # of defects
#       iterDataDict[incStories] = incomplete stories
#       iterDataDict[incPoints] =  incomplete points
#       iterDataDict[accPoints] =  accepted points

    def getArtifact(self, artifactType, projName, query):
        response = self.rally.get(artifactType, fetch='FormattedID,Iteration,Name,ScheduleState,PlanEstimate,RevisionHistory', project=projName, query=query)
        if response.errors:
            sys.stdout.write("\n".join(errors))
            sys.exit(1)

        artDict = dict()
        for artifact in response:
            print artifact.FormattedID

    # Iteration end date must lie between start and end date
    def buildIterDateQuery(self, sDate, eDate):
        dateQuery = '(Iteration.EndDate > ' + sDate + ') AND (Iteration.EndDate < ' + eDate + ')'
        return dateQuery

    def getProjStories(self, proj, dateQuery):
        iterationDict = dict()
        iterationDict["story"] = self.getArtifact('UserStory', proj, dateQuery)
        iterationDict["defect"] = self.getArtifact('Defect', proj, dateQuery)
        return iterationDict

    def cipHealth(self, projectList, year, qtr):
        cipHealthDict = dict()
        quarter = quarters[qtr]
        sDate = year + '-' + quarter[0]
        eDate = year + '-' + quarter[1]
        dateQuery = self.buildIterDateQuery(sDate, eDate)
        projList = projectList.split(',')
        for proj in projList:
            cipHealthDict[proj] = self.getProjStories(proj, dateQuery)
        return cipHealthDict

class ExportData:
    def __init__(self, iteration, storiesDict):
        self.artifacts = storiesDict['Stories']
        self.summary = storiesDict['Summary']
        self.printToExcel(iteration)
            
    def printToExcel(self, iteration):
        sBold = xlwt.easyxf('font: bold on')
        wb = xlwt.Workbook()

        # Summary Data
        wsSum = wb.add_sheet("Summary")
        colLabel = ['Stories', 'Incomplete Stories', 'Incomplete Stories %', 'Points', 'Incomplete Points', 'Incomplete Points %']
        row = 0
        col = 0

        # Print Column Headers
        for label in colLabel:
            wsSum.write(row, col, label, sBold)
            col = col + 1
        row = row + 1

        # Print Summary
        data_stories = self.summary['Stories']
        data_incStories = self.summary['IncompleteStories']
        data_points = self.summary['Points']
        data_incPoints = self.summary['IncompletePoints']
        wsSum.write(row, 0, data_stories)
        wsSum.write(row, 1, data_incStories)
        wsSum.write(row, 2, float(data_incStories)/float(data_stories) * 100)
        wsSum.write(row, 3, data_points)
        wsSum.write(row, 4, data_incPoints)
        wsSum.write(row, 5, float(data_incPoints)/float(data_points) * 100)

        # Iteration Data
        colLabel = ['FormattedName', 'State', 'ID', 'Name', 'Est']
        for artifactType in self.artifacts.keys():
            artifactList = self.artifacts[artifactType]
            wsArt = wb.add_sheet(artifactType)
            row = 0
            col = 0

            # Print Column Headers
            for label in colLabel:
                wsArt.write(row, col, label, sBold)
                col = col + 1
            row = row + 1

            # Print Task Data
            for artifact in artifactList:
                formName = artifact['ID'] + ": " + artifact['Name']
                if artifact['Est']:
                    formName = formName + ' (' + str(int(artifact['Est'])) + 'pts)'
                wsArt.write(row, 0, formName)
                wsArt.write(row, 1, artifact['State'])
                wsArt.write(row, 2, artifact['ID'])
                wsArt.write(row, 3, artifact['Name'])
                wsArt.write(row, 4, artifact['Est'])
                row = row + 1

        wb.save('EosReport_' + iteration + '.xls')

def main(argv):
    # INPUT: ProjectList, StartDate, EndDate
    projectName = None
    try:
        opts, args = getopt.getopt(argv, "hp:y:q:", ["projectList=","year=","qtr="])
    except getopt.GetoptError:
        print 'python ProjectHealth.py -p <project list> -y <year YYYY> -e <quarter Q1-Q4>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == "-h":
            print 'python ProjectHealth.py -p <project list> -y <year YYYY> -e <quarter Q1-Q4>'
            sys.exit()
        elif opt in ("-p", "--projectList"):
            projectList = arg
        elif opt in ("-y", "--year"):
            year = arg
        elif opt in ("-q", "--quarter"):
            qtr = arg
    cipHealth = RallyStories().cipHealth(projectList, year, qtr)
    #sprintDict = RallyStories().sprintStories(projectName, sprintName)
    #ExportData(sprintName, cipHealth)

if __name__ == '__main__':
    main(sys.argv[1:])