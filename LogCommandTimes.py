import fnmatch, os, sys, getopt
import re
import sqlite3
import xlwt

def getFiles(dirPath):
    fileList = list()

    for root, dirnames, filenames in os.walk(dirPath):
        #uiList = fnmatch.filter(filenames, 'ui.log')
        uiList = []
        bllList = fnmatch.filter(filenames, 'bll.session.log*')
        logList = uiList + bllList
        for log in logList:
            logFile = os.path.join(root, log)
            fileList.append(logFile)
    return fileList

def parseLogs(fileList):
    sheetCache = dict()
    wb = xlwt.Workbook()
    for log in fileList:
        sheetName = log.split('\\')[1]
        if sheetName not in sheetCache:
            ws = wb.add_sheet(sheetName)
            sheetCache[sheetName] = dict()
            sheetCache[sheetName]['sheet'] = ws
            sheetCache[sheetName]['row'] = 0
        ws = sheetCache[sheetName]['sheet']
        row = sheetCache[sheetName]['row']
        sheetCache[sheetName]['row'] = parseFile(ws, row, log)
    wb.save('SadaLogs.xls')

def parseFile(sheet, row, logfile):
    num = xlwt.easyxf("","##.####")
    sBold = xlwt.easyxf('font: bold on')
    try:
        f = open(logfile, 'r')
    except:
        print 'ERROR: ' + logfile + ': does not exists'
        return

    for line in f.readlines():
        # get timestamp, command name, and execution status
        match = re.match('.*\s(.*Command.*\)).*(state: Complete).*took:\s(.*)\ssec', line)
        if match is not None:
            cmdName = match.group(1)
            stateStr = match.group(2)
            if stateStr.split()[1] == "Complete":
                time = match.group(3)
                if time != '0.000':
                    lsplit = line.split()
                    date = lsplit[0]
                    timestamp = lsplit[1]
                    if 'ApplyTo' in cmdName:
                        sheet.write(row, 0, date, sBold)
                        sheet.write(row, 1, timestamp, sBold)
                        sheet.write(row, 2, cmdName, sBold)
                        sheet.write(row, 3, time, sBold)
                    else:
                        sheet.write(row, 0, date)
                        sheet.write(row, 1, timestamp)
                        sheet.write(row, 2, cmdName)
                        sheet.write(row, 3, time)
                    row += 1
    return row

def main(argv):
    dirPath = None
    try:
        opts, args = getopt.getopt(argv, "hd:", ["dirpath="])
    except getopt.GetoptError:
        print 'python LogCommandTimes.py -d <dir path>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == "-h":
            print 'python LogCommandTimes.py -d <dir path>'
            sys.exit()
        elif opt in ("-d", "--dirpath"):
            dirPath = arg

    if dirPath is not None:
        print str(dirPath)
        fileList = getFiles(dirPath)
        parseLogs(fileList)
    else:
        print "ERROR: Devs need to be set."

if __name__ == '__main__':
    main(sys.argv[1:])