import sys, getopt
from pyral import Rally, rallySettings
import unicodedata
import re
import xlwt
from time import localtime, strftime
from datetime import date
import sqlite3

# RunCmd: python <script> --conf=<cfg no ".cfg" extension>
users = ['Abeysundara','Ofiesh','DuBois','Kodama','Yoshino','Duque','Tom','Shimokawa','Yamamoto','Kanja', 'Sawai']
users2 = ['Matsumoto', 'Cordeiro', 'LaMont', 'Yang']
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

    def getIterDates(self, iterName):
        query = 'Name = "' + iterName + '"'
        response = self.rally.get('Iteration', fetch='Name,StartDate,EndDate', project=None, query=query)
        for iteration in response:
            return (iteration.StartDate, iteration.EndDate)
        return (None, None)

    def parseDateString(self, pDate):
        myDate = None
        m = re.match('(\d{4})-(\d{2})-(\d{2})', pDate)
        if m != None:
            myDate = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return myDate

    def getIterPercent(self, repStart, repEnd, iterStart, iterEnd):
        iterPercent = 0
        iterDays = float((iterEnd - iterStart).days)
        startDelta = float((iterStart - repStart).days)
        endDelta = float((repEnd - iterEnd).days)
        if startDelta >= 0 and endDelta >= 0:
            iterPercent = float(1)
        elif startDelta < 0:
            iterPercent = (iterDays + startDelta)/iterDays
        elif endDelta < 0:
            iterPercent = (iterDays + endDelta)/iterDays
        return iterPercent

    def getIterationInfo(self, iteration, userList, repStartDate, repEndDate, iterStartDate, iterEndDate):
        capDict = dict()
        capDict['StartDate'] = str(iterStartDate)
        capDict['EndDate'] = str(iterEndDate)
        capDict['Percent'] = self.getIterPercent(repStartDate, repEndDate, iterStartDate, iterEndDate)
        capDict['Users'] = dict()
        for userCap in iteration.UserIterationCapacities:
            if userCap.User.LastName not in userList:
                continue
            capDict['Users'][userCap.User.Name] = userCap.Capacity
        return capDict

    def buildUserTasksQuery(self, sDate, eDate, userList):
        dateQuery = '((LastUpdateDate > ' + sDate + ') AND (LastUpdateDate < ' + eDate + ')) AND '
        userQuery = ''
        for (idx, user) in enumerate(userList):
            if idx == 0:
                userQuery = '(Owner.LastName = "' + user + '")'
            else:
                userQuery = '(' + userQuery + ' OR (Owner.LastName = "' + user + '"))'
        return dateQuery + userQuery

    def getUserTasks(self, startDate, endDate, userList):
        iterDict = dict()
        userTasksDict = dict()
        query = self.buildUserTasksQuery(startDate, endDate, userList)
        response = self.rally.get('Task', fetch='Name,Owner,WorkProduct,Project,Iteration,LastUpdateDate,State,Description,Notes,Actuals,Estimate', project=None, query=query)
        for task in response:
            if not task.Owner:
                continue
            if task.Owner.Name not in userTasksDict:
                userTasksDict[task.Owner.Name] = list()

            taskDict = dict()
            taskDict['Iteration'] = ''
            taskDict['Capacity'] = 0
            taskDict['Percent'] = 0
            if task.Iteration:
                repStartDate = self.parseDateString(startDate)
                repEndDate = self.parseDateString(endDate)
                iterStartDate = self.parseDateString(task.Iteration.StartDate)
                iterEndDate = self.parseDateString(task.Iteration.EndDate)
                if (iterStartDate >= repStartDate and iterStartDate <= repEndDate) or (iterEndDate >= repStartDate and iterEndDate <= repEndDate):
                    taskDict['Iteration'] = task.Iteration.Name
                    # Get IterInfo
                    if task.Iteration.Name not in iterDict:
                        iterDict[task.Iteration.Name] = self.getIterationInfo(task.Iteration, userList, repStartDate, repEndDate, iterStartDate, iterEndDate)
                    iterInfo = iterDict[taskDict['Iteration']]
                    if task.Owner.Name in iterInfo['Users'].keys():
                        taskDict['Capacity'] = iterInfo['Users'][task.Owner.Name]
                        if taskDict['Capacity'] is None:
                            taskDict['Capacity'] = 0
                    taskDict['Percent'] = iterInfo['Percent']
                else:
                    continue

            taskDict['Name'] = task.Name
            taskDict['USNumber'] = ''
            taskDict['USName'] =  ''
            if task.WorkProduct:
                taskDict['USNumber'] = task.WorkProduct.FormattedID
                taskDict['USName'] = task.WorkProduct.Name
            taskDict['Project'] = ''
            if task.Project:
                taskDict['Project'] = task.Project.Name
            task.LastUpdateDate = task.LastUpdateDate.split('T')[0]
            taskDict['Date'] = task.LastUpdateDate
            taskDict['State'] = task.State
            taskDict['Description'] = task.Description
            taskDict['Notes'] = task.Notes
            taskDict['Actuals'] = task.Actuals
            taskDict['Estimate'] = task.Estimate
            userTasksDict[task.Owner.Name].append(taskDict)
        return (userTasksDict, iterDict)

    # Get project user stories
    def userTasks(self, sDate, eDate, iterName):
        userTasksIteration = self.getUserTasks(sDate, eDate, users, iterName)
        return userTasksIteration

class ExportData:
    def __init__(self, userTasks, iterCaps, title):
        self.userTasks = userTasks
        self.iterCaps = iterCaps
        self.printToExcel(title)

    def sortIterInfoByUser(self):
        userCapDict = dict()
        for iterName in self.iterCaps:
            iterInfo = self.iterCaps[iterName]
            percent = iterInfo['Percent']
            for user in iterInfo['Users'].keys():
                userName = user
                capacity = iterInfo['Users'][userName]
                if capacity is None:
                    capacity = 0
                if userName not in userCapDict:
                    userCapDict[userName] = capacity * percent
                userCapDict[userName] = userCapDict[userName] + (capacity * percent)
        return userCapDict

    def printToExcel(self, title):
        sBold = xlwt.easyxf('font: bold on')
        wb = xlwt.Workbook()

        userCapDict = self.sortIterInfoByUser()

        sumRow = 1
        wsSummary = wb.add_sheet('Summary')
        wsSummary.write(0, 0, title, sBold)
        wsSummary.write(0, 1, 'Capacity', sBold)
        wsSummary.write(0, 2, 'Total Actual', sBold)
        wsSummary.write(0, 3, 'Total Estimate', sBold)
        wsSummary.write(0, 4, 'A/E', sBold)
        wsSummary.write(0, 5, 'A/E Norm', sBold)
        wsSummary.write(0, 6, 'A/C', sBold)
        wsSummary.write(0, 7, 'E/C', sBold)
        rowLabel = ['Date', 'State', 'Project', 'Iteration', 'USNumber', 'USName', 'Task', 'Description', 'Notes', 'Actuals', 'Estimate', 'A/E', 'Capacity', 'Percent', 'ActCap']
        for userKey in self.userTasks.keys():
            user = self.userTasks[userKey]
            wsUser = wb.add_sheet(userKey)
            row = 0
            col = 0

            # Print Column Headers
            for label in rowLabel:
                wsUser.write(row, col, label, sBold)
                col = col + 1
            row = row + 1

            # Print Task Data
            totalAct = 0
            totalEst = 0
            for task in user:
                wsUser.write(row, 0, task['Date'])
                wsUser.write(row, 1, task['State'])
                wsUser.write(row, 2, task['Project'])
                wsUser.write(row, 3, task['Iteration'])
                wsUser.write(row, 4, task['USNumber'])
                wsUser.write(row, 5, task['USName'])
                wsUser.write(row, 6, task['Name'], sBold)
                wsUser.write(row, 7, task['Description'])
                wsUser.write(row, 8, task['Notes'])
                act = task['Actuals']
                est = task['Estimate']
                if est is None:
                    est = 0
                    act = 0
                if act is None or act == 0:
                    act = est
                wsUser.write(row, 9, act)
                wsUser.write(row, 10, est)
                totalAct = totalAct + act
                totalEst = totalEst + est
                if est > 0:
                    actEst = act/est
                    wsUser.write(row, 11, actEst)
                wsUser.write(row, 12, task['Capacity'])
                wsUser.write(row, 13, task['Percent'])
                wsUser.write(row, 14, task['Capacity'] * task['Percent'])
                row = row + 1
            avgActEst = totalAct/totalEst
            wsUser.write(row, 9, totalAct)
            wsUser.write(row, 10, totalEst)
            wsUser.write(row, 11, avgActEst, sBold)

            wsSummary.write(sumRow, 0, userKey)
            wsSummary.write(sumRow, 1, userCapDict[userKey])
            wsSummary.write(sumRow, 2, totalAct)
            wsSummary.write(sumRow, 3, totalEst)
            wsSummary.write(sumRow, 4, format(avgActEst, '.2f'))
            wsSummary.write(sumRow, 5, format(1-avgActEst, '.2f'))
            wsSummary.write(sumRow, 6, format(totalAct/userCapDict[userKey], '.2f'))
            wsSummary.write(sumRow, 7, format(totalEst/userCapDict[userKey], '.2f'))
            sumRow = sumRow + 1
        wb.save('UsersTasks' + '_' + title + '.xls')

def main(argv):
    try:
        opts, args = getopt.getopt(argv, "hy:q:i:", ["year=","quarter=","iteration="])
    except getopt.GetoptError:
        print 'python DevQtrReport.py -y <year> -q <quarter> -i <iteration>'
        sys.exit(2)
    year = None
    qtr = None
    iterName = None
    for opt, arg in opts:
        if opt == "-h":
            print 'python DevQtrReport.py -y <year> -q <quarter>'
            sys.exit()
        elif opt in ("-y", "--year"):
            year = arg
        elif opt in ("-q", "--quarter"):
            qtr = arg
        elif opt in ("-i", "--iteration"):
            iterName = arg
    if not year:
        year = "2015"
        qtr = "Q1"

    title = '2015-Q1'
    sDate = '2015-01-01'
    eDate = '2015-03-31'
    if iterName is None:
        quarter = quarters[qtr]
        sDate = year + '-' + quarter[0]
        eDate = year + '-' + quarter[1]
        title = year + '-' + qtr
    else:
        title = iterName
        (sDate, eDate) = RallyStories().getIterDates(title)

    if sDate is not None and eDate is not None:
        (userTasks, iterCaps) = RallyStories().userTasks(sDate, eDate, iterName)
        ExportData(userTasks, iterCaps, title)
    else:
        print iterName + " doesn't exist or the start and end dates for the iteration aren't defined."

if __name__ == '__main__':
    main(sys.argv[1:])