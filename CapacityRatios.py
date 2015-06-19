import sys, getopt
from pyral import Rally, rallySettings
import unicodedata
import re
import xlwt
from time import localtime, strftime
import sqlite3

# RunCmd: python <script> --conf=<cfg no ".cfg" extension>

class RallyStories:
    def __init__(self):
        self.errors = ''
        self.conn = sqlite3.connect('EosReport.db')
        self.cursor = self.conn.cursor()
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

    def buildUserQuery(self, userList, startDate, endDate):
        dateQuery = ''
        if startDate is not None and endDate is not None:
            dateQuery = ' AND ((' + str(startDate) + ' <= LastUpdateDate) AND (' + str(endDate) + ' >= LastUpdateDate))'
        if startDate is not None:
            dateQuery = ' AND (' + str(startDate) + ' <= LastUpdateDate)'
        if endDate is not None:
            dateQuery = ' AND (' + str(endDate) + ' >= LastUpdateDate)'

        userQuery = ''
        for (idx, user) in enumerate(userList):
            if '@spirent.com' not in user:
                user = user + '@spirent.com'
            if idx == 0:
                userQuery = '(Owner.UserName = "' + user.strip().lower() + '")'
            elif idx == 1:
                userQuery = userQuery + ' OR (Owner.UserName = "' + user.strip().lower() + '")'
            else:
                userQuery = '(' + userQuery + ') OR (Owner.UserName = "' + user.strip().lower() + '")'

        if dateQuery != '':
            userQuery = '(' + userQuery + ')'
        return userQuery + dateQuery

    def userCapacity(self, userList, startDate, endDate):
        query = self.buildUserQuery(userList, startDate, endDate)
        response = self.rally.get('Task', fetch='ObjectID,Name,Owner,Iteration,LastUpdateDate,Actuals,Estimate,State', project=None, query=query)
        if response.errors:
            sys.stdout.write("\n".join(response.errors))
            sys.exit(1)

        print response
        iterDict = dict()
        #{iterName: {iter: Iteration, users: [userList]}}
        userDict = dict()
        for task in response:
            if task.Owner is None:
                continue
            userName = task.Owner.Name
            if userName not in userDict.keys():
                userDict[userName] = dict()
            #userDict[userName][]

            print task.Owner.Name
            print task.Owner.DisplayName
            print task.Owner.UserName

        # user: {capacity: val, actual: val, estimate: val, ae: val, aeNorm: val, ac: val, ec: val}
        # only for completed tasks
        # prorate capacity based on dates entered

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
    def getArtifact(self, artifactType, projName, iteration, toWrite):
        query = 'Iteration.Name = "' + iteration +'"'
        response = self.rally.get(artifactType, fetch='FormattedID,Iteration,Name,ScheduleState,PlanEstimate,RevisionHistory', project=projName, query=query)
        if response.errors:
            sys.stdout.write("\n".join(errors))
            sys.exit(1)

        artList = list()
        artifact = None
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

        # Write to DB
        if toWrite and artifact is not None:
            stories = self.summary['Stories']
            incStories = self.summary['IncompleteStories']
            points = self.summary['Points']
            incStoriesPer = float(incStories)/float(stories) * 100
            incPoints = self.summary['IncompletePoints']
            incPointsPer = float(incPoints)/float(points) * 100
            self.cursor.execute("INSERT OR REPLACE INTO iterSummary (OID, IterName, Stories, IncompleteStories, IncompleteStoryPercent, Points, IncompletePoints, IncompletePointPercent) values(?,?,?,?,?,?,?,?)", (artifact.Iteration.oid,artifact.Iteration.Name,stories,incStories,incStoriesPer,points,incPoints,incPointsPer))
            self.conn.commit()
        return artList

    def getSprintStories(self, projectName, sprintName, toWrite):
        sprintDict = dict()
        sprintDict["story"] = self.getArtifact('UserStory', projectName, sprintName, toWrite)
        sprintDict["defect"] = self.getArtifact('Defect', projectName, sprintName, toWrite)
        return sprintDict

    def getUserIterCaps(self, uicList):
        uicDict = dict()
        for uic in uicList:
            uicDict[uic.User.Name] = uic.Capacity
        return uicDict

    def getSprintCapacity(self, projectName, sprintName, toWrite):
        query = 'Iteration.Name = "' + sprintName +'"'
        response = self.rally.get('Task', fetch='Name,Owner,Iteration,Actuals,Estimate,State', project=projectName, query=query)
        if response.errors:
            sys.stdout.write("\n".join(response.errors))
            sys.exit(1)

        iteration = None
        userIterCap = dict()
        userCapDict = dict()
        for task in response:
            if task.Owner is None:
                continue
            if not iteration:
                iteration = task.Iteration
            if task.Owner.Name not in userCapDict.keys():
                userCapDict[task.Owner.Name] = dict()
                userCapDict[task.Owner.Name]['taskList'] = list()
                userCapDict[task.Owner.Name]['capacity'] = None
                userCapDict[task.Owner.Name]['totalAct'] = 0
                userCapDict[task.Owner.Name]['totalEst'] = 0

                if not userIterCap:
                    userIterCap = self.getUserIterCaps(task.Iteration.UserIterationCapacities)
                if not userCapDict[task.Owner.Name]['capacity']:
                    if task.Owner.Name in userIterCap.keys():
                        userCapDict[task.Owner.Name]['capacity'] = userIterCap[task.Owner.Name]

            act = task.Actuals
            est = task.Estimate
            taskInfoDict = dict()
            taskInfoDict['name'] = task.Name
            if est is None:
                act = 0
                est = 0
            if act is None or act == 0:
                act = est
            taskInfoDict['act'] = act
            taskInfoDict['est'] = est
            userCapDict[task.Owner.Name]['taskList'].append(taskInfoDict)
            if task.State == 'Completed':
                userCapDict[task.Owner.Name]['totalAct'] = userCapDict[task.Owner.Name]['totalAct'] + taskInfoDict['act']
                userCapDict[task.Owner.Name]['totalEst'] = userCapDict[task.Owner.Name]['totalEst'] + taskInfoDict['est']

        # Save Cap to DB
        for user in userCapDict.keys():
            uc = userCapDict[user]
            cap = uc['capacity']
            t_act = uc['totalAct']
            t_est = uc['totalEst']
            if t_est != 0:
                ave = t_act/t_est * 100
            else:
                ave = 0
            ave_norm = 100-ave
            if cap is None:
                print 'Skipping ' + iteration.Name + " -> " + user + ': capcacity is None'
                continue
            else:
                avc = t_act/cap * 100
                evc = t_est/cap * 100
            self.cursor.execute("INSERT OR REPLACE INTO userCapacity (IterID, UserName, Capacity, TotalActual, TotalEstimate, AvE, AvE_Norm, AvC, EvC) values(?,?,?,?,?,?,?,?,?)", (iteration.oid,user,cap,t_act,t_est,ave,ave_norm,avc,evc))
            self.conn.commit()

        return userCapDict

    def createTables(self):
        self.cursor.execute("create table if not exists iterSummary(OID integer UNIQUE, IterName text, Stories integer, IncompleteStories integer, IncompleteStoryPercent real, Points integer, IncompletePoints integer, IncompletePointPercent real)")
        self.cursor.execute("create table if not exists userCapacity(IterID integer, UserName text, Capacity integer, TotalActual real, TotalEstimate real, AvE real, AvE_Norm real, AvC real, EvC real, unique(IterID, UserName))")
        self.conn.commit()

    # Get project user stories
    def sprintStories(self, projectName, sprintName, toWrite):
        if toWrite:
            self.createTables()

        self.summary = dict()
        self.summary['Stories'] = 0
        self.summary['Points'] = 0
        self.summary['IncompleteStories'] = 0
        self.summary['IncompletePoints'] = 0

        self.sprint = dict()
        self.sprint['Summary'] = self.summary
        self.sprint['Stories'] = self.getSprintStories(projectName, sprintName, toWrite)
        self.sprint['Capacity'] = self.getSprintCapacity(projectName, sprintName, toWrite)

        return self.sprint

    def outputGraph(self):
        chartDict = self.buildChartDict()
        
        f = open('EndOfSprintReport.html','w')
        header = '''
        <!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
        <html>
            <head>
            <title>EOS Charts</title>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
            <script src="http://code.jquery.com/jquery-1.9.1.js" type="text/javascript"></script>
            <script src="http://code.highcharts.com/highcharts.js" type="text/javascript"></script>
            <script src="http://code.highcharts.com/highcharts-more.js"></script>
            <script src="http://code.highcharts.com/modules/exporting.js" type="text/javascript"></script>
            </head>
            <body>
        '''
        f.write(header)

        #<script type="text/javascript">
        #        $(function () {
        #            $('#&&ITER_NAME_SHORT&&').highcharts({
        #                title: {text: '&&ITER_NAME&&'},
        #                xAxis: [{
        #                    categories: ['Points', 'Incomplete Points', 'Stories', 'Incomplete Stories'],
        bodyTemp = '''
            <script type="text/javascript">
                $(function () {
                    $('#&&ITER_NAME_SHORT&&').highcharts({
                        title: {text: '&&ITER_NAME&&'},
                        xAxis: [{
                            categories: ['Inc Points %', 'Inc Stories %'],
                            id: 'iterPer'
                        },{
                            categories: ['Points', 'Stories'],
                            id: 'iterData'
                        },{
                            categories: [&&USER_LIST&&],
                            id: 'users'
                        }],
                        yAxis: [{
                            title: {
                                text: 'Against Capacity'
                            },
                            labels: {
                                format: '{value} %'
                            },
                            tickInterval: 20,
                            plotLines: [{
                                value: 75,
                                color: 'red',
                                dashStyle: 'shortdash',
                                width: 2,
                                label: {
                                    text: 'Minimum Capacity Threshold (75%)'
                                }
                            }]
                        },{
                            title: {
                                text: 'Iteration %'
                            },
                            labels: {
                                format: '{value} %'
                            },
                            opposite: true
                        },{
                            title: {
                                text: 'Iteration Counts'
                            },
                            opposite: true
                        }],
                        plotOptions: {
                            columnrange: {
                                dataLabels: {
                                    enabled: true
                                }
                            },
                            column: {
                                dataLabels: {
                                    enabled: true
                                }
                            }
                        },
                        series: [{
                            type: 'column',
                            name: 'Iteration %',
                            data: [&&ITER_PER&&],
                            xAxis: 0,
                            yAxis: 1,
                            color: 'rgba(0,255,255,0.5)'
                        },{
                            type: 'column',
                            name: 'Iteration Data',
                            data: &&ITER_DATA&&,
                            xAxis: 1,
                            yAxis: 2,
                            visible: false,
                            stacking: 'normal'
                        },{
                            type: 'columnrange',
                            name: 'Developer Data',
                            data: &&DEV_DATA&&,
                            xAxis: 2
                        }]
                    })
                });
            </script>
            <div id="&&ITER_NAME_SHORT&&" style="min-width: 310px; height: 400px; margin: 0 auto"></div>
        '''
        for oid in sorted(chartDict.keys(), reverse=True):
            itBody = bodyTemp
            chartInfo = chartDict[oid]
            for key in chartInfo:
                itBody = itBody.replace('&&'+key+'&&', chartInfo[key], 2)
            f.write(itBody)

        trailer = '''
            </body>
        </html>
        '''
        f.write(trailer)
        f.close()

    def buildChartDict(self):
        chartDict = dict()
        self.cursor.execute('SELECT OID, IterName, IncompletePointPercent, IncompleteStoryPercent, Points, IncompletePoints, Stories, IncompleteStories from iterSummary')
        # Points, IncompletePoints, Stories, IncompleteStories
        iterInfo = self.cursor.fetchall()
        sumHelperDict = dict()
        for itInfo in iterInfo:
            infoDict = dict()
            infoDict['ITER_NAME_SHORT'] = itInfo[1].replace(' ', '_', 5)
            infoDict['ITER_NAME'] = itInfo[1]
            infoDict['ITER_PER'] = str(itInfo[2:4]).strip('()')
            infoDict['ITER_DATA'] = str(self.formatIterData(itInfo[4:]))

            q = (itInfo[0],)
            userInfo = self.cursor.execute('SELECT IterId, UserName, AvC, EvC, Capacity, TotalActual, TotalEstimate from userCapacity as uc WHERE uc.IterId = ?', q)
            ulist = list()
            capList = list()
            for uInfo in userInfo:
                userName = str(uInfo[1])
                ulist.append(userName)
                cap = self.formatDevData(uInfo[2], uInfo[3])
                capList.append(cap)

                # Summary Helper
                if uInfo[1] not in sumHelperDict.keys():
                    sumHelperDict[userName] = dict()
                    sumHelperDict[userName]['Capacity'] = 0
                    sumHelperDict[userName]['Actual'] = 0
                    sumHelperDict[userName]['Estimate'] = 0
                sumHelperDict[userName]['Capacity'] += uInfo[4]
                sumHelperDict[userName]['Actual'] += uInfo[5]
                sumHelperDict[userName]['Estimate'] += uInfo[6]

            infoDict['USER_LIST'] = str(ulist).strip('[]')
            infoDict['DEV_DATA'] = str(capList)

            chartDict[itInfo[0]] = infoDict

        summaryDict = dict()
        summaryDict['ITER_NAME_SHORT'] = 'Summary'
        summaryDict['ITER_NAME'] = 'Summary'
        summaryDict['ITER_PER'] = '0,0'
        summaryDict['ITER_DATA'] = '[0,0]'
        sumUserList = list()
        sumCapList = list()
        for userKey in sorted(sumHelperDict.keys()):
            sumUserList.append(userKey)
            cap = sumHelperDict[userKey]['Capacity']
            act = sumHelperDict[userKey]['Actual']
            est = sumHelperDict[userKey]['Estimate']
            avc = act/cap * 100
            evc = est/cap * 100
            sumCap = self.formatDevData(avc, evc)
            sumCapList.append(sumCap)
        summaryDict['USER_LIST'] = str(sumUserList).strip('[]')
        summaryDict['DEV_DATA'] = str(sumCapList)
        chartDict['999999999999'] = summaryDict
        return chartDict

    def formatIterData(self, dataArr):
        incPoints = dataArr[1]
        point = dataArr[0] - incPoints
        incStory = dataArr[3]
        story = dataArr[2] - incStory
        dataList = list()
        incPointDict = dict()
        incPointDict['x'] = 0
        incPointDict['y'] = incPoints
        incPointDict['color'] = 'rgba(255,0,0,0.5)'
        incPointDict['name'] = 'Incomplete Points'
        dataList.append(incPointDict)
        pointDict = dict()
        pointDict['x'] = 0
        pointDict['y'] = point
        pointDict['color'] = 'rgba(10,200,200,0.5)'
        pointDict['name'] = 'Points'
        dataList.append(pointDict)
        incStoryDict = dict()
        incStoryDict['x'] = 1
        incStoryDict['y'] = incStory
        incStoryDict['color'] = 'rgba(255,0,0,0.5)'
        incStoryDict['name'] = 'Incomplete Stories'
        dataList.append(incStoryDict)
        storyDict = dict()
        storyDict['x'] = 1
        storyDict['y'] = story
        storyDict['color'] = 'rgba(10,200,200,0.5)'
        storyDict['name'] = 'Stories'
        dataList.append(storyDict)
        return dataList

    def formatDevData(self, AvC, EvC):
        devDict = dict()
        f_avc = float(format(AvC, '.2f'))
        f_evc = float(format(EvC, '.2f'))
        if abs(f_avc-f_evc) <= 10 or abs(f_evc-f_avc) <= 10:
            devDict['color'] = 'rgba(0,200,0,0.5)'
        elif f_avc > f_evc:  # Underestimate
            devDict['color'] = 'rgba(200,0,0,0.5)'
            devDict['dataLabels'] = dict()
            devDict['dataLabels']['format'] = '{y} U'
        else:                # Overestimate
            devDict['color'] = 'rgba(200,200,0,0.5)'
            devDict['dataLabels'] = dict()
            devDict['dataLabels']['format'] = '{y} O'
        devDict['low'] = min(f_avc, f_evc)
        devDict['high'] = max(f_avc, f_evc)
        return devDict

class ExportData:
    def __init__(self, iteration, storiesDict):
        self.artifacts = storiesDict['Stories']
        self.summary = storiesDict['Summary']
        self.capacity = storiesDict['Capacity']
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

        # Capacity Data
        wsCap = wb.add_sheet('Capacity')
        colLabel = ['User','Capacity','Total Act','Total Est','A/E','A/E Norm','A/C','E/C']
        row = 0
        col = 0

        # Print Column Headers
        for label in colLabel:
            wsCap.write(row, col, label, sBold)
            col = col + 1
        row = row + 1

        # Print Capacity Info
        for user in self.capacity.keys():
            cap = self.capacity[user]['capacity']
            totAct = self.capacity[user]['totalAct']
            totEst = self.capacity[user]['totalEst']
            if totEst != 0:
                ae = totAct/totEst
            else:
                ae = 0

            wsCap.write(row, 0, user)
            wsCap.write(row, 1, cap)
            wsCap.write(row, 2, totAct)
            wsCap.write(row, 3, totEst)
            wsCap.write(row, 4, format(ae, '.2f'))
            wsCap.write(row, 5, format(1-ae, '.2f'))
            if cap is not None:
                wsCap.write(row, 6, format(totAct/cap, '.2f'))
                wsCap.write(row, 7, format(totEst/cap, '.2f'))
            row = row + 1

        wb.save('EosReport_' + iteration + '.xls')

def main(argv):
    userList = None
    startDate = None
    endDate = None
    try:
        opts, args = getopt.getopt(argv, "hu:s:e:", ["userList=","startDate=","endDate="])
    except getopt.GetoptError:
        print 'python CapacityRatios.py -u <userList> -s <startDate> -e <endDate>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == "-h":
            print 'python CapacityRatios.py -u <userList> -s <startDate> -e <endDate>'
            sys.exit()
        elif opt in ("-u", "--userList"):
            userList = arg.split(',')
        elif opt in ("-s", "--startDate"):
            startDate = arg
        elif opt in ("-e", "--endDate"):
            endDate = arg

    if userList is not None:
        # users -> tasks -> iteration -> iter_cap, task_act, task_est
        userCapDict = RallyStories().userCapacity(userList, startDate, endDate)
        # user: {capacity: val, actual: val, estimate: val, ae: val, aeNorm: val, ac: val, ec: valu}
        # only for completed tasks
        # prorate capacity based on dates entered

        #sprintDict = RallyStories().sprintStories(projectName, sprintName, toWrite)
        #ExportData(sprintName, sprintDict)
        #RallyStories().outputGraph()
    else:
        print "UserList not specified."

if __name__ == '__main__':
    main(sys.argv[1:])