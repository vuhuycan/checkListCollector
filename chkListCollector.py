#!/bin/python3

import subprocess
import openpyxl
import sys
import re
from copy import copy
import argparse
import os


print('////////////////////////////////////////////////////////////////')
print("chkListCollector v2.0r2022.08.05b, Beta. !!!No right reserved!!!")
print('////////////////////////////////////////////////////////////////')
print('')
#v0.1r2022.04.01: support multiple commands in one row (main target is for FaultInsert and RepairVerif checklists)
#v0.1r2022.04.02: fix bug: when there's multiple cmds in one row: print out wrong cmd/result, bad padding/coloring the inserted lines
#v0.1r2022.04.04: now search for "Working dir" and "command" in arbitrary column, not fixed anymore
#v0.1r2022.07.25: fix bug that noDelForVi, mergeAcross and noEnterBash do not work.
#v0.1r2022.08.02: fix bug: cannot regconize command start with space character
#v0.1r2022.08.02b: fix bug: add timeout=3secs to Popen commands, so faulty commands does not halt the program
#v0.1r2022.08.03: add feature: if dir does not contain a user specified string ( dirProtect ), a warning is raised.
#v2.0r2022.08.05: change program structure to oop for esier maintenance and readability.
#                 add CLI, auto remove existed output file.
#v2.0r2022.08.05b: fix bug: insert rows will not keep merged cells format below the inserted row. 
#                           inserting rows does not update new value for sheet.max_row

#----------------------------------------------------
#sometimes, you want to edit this portion to modify tool behavior, but very rare, so I did not make them tool input
#-------------------
#the tool only regconize following commands:
knownCmdList = ('cd', 'ls', 'll', 'vi', 'cat', 'head', 'tail', 'grep', 'egrep', 'sed', 'awk', 'diff', 'vimdiff')
#when the tool meets these one of these commands, it would skip it:
inhibitCmdList = ('vi', 'vimdiff') #, 'cat')

#color code of "working dir":
wkDirColor = 'FFFFC000'
#color code of "command execution":
cmdColor = 4
#color code of "command result":
rsltColor = 9

#----------------------------------------------------

#----------------------------------------------------
#parse inputs
#-------------------
# create a parser object
parser = argparse.ArgumentParser(description = "A Collector for your checklist.")

# add arguments
parser.add_argument("-i","--inFile", nargs = 1, metavar = "file", type = str,
                     help = "Specify your input xlsx file path. \
                             Notes: - Only xlsx file is accepted.\
                                    - All pictures and drawing objects would not be kept in the output checklist.\
                                    These are restrictions of the openpyxl lib.")
parser.add_argument("-o","--outFile", nargs = 1, metavar = "file", type = str,
                     help = "Specify your output xlsx file path.") 
parser.add_argument("-sh","--enterBash", action = 'store_true',
                     help = "If there's error with a command, this switch allow you to enter bash terminal, \
                             allow you to type a new command and replace it in the checklist.") 
parser.add_argument("-m","--mergeAcross", action = 'store_true',
                     help = "This would merge across all excel columns of the command's result in the checklist.") 
parser.add_argument("-k","--keepViRslt", action = 'store_true',
                     help = "This would keep the result of vi command instead of deleting it by default.") 
parser.add_argument("-p","--dirProtector", nargs = '?', metavar = "keyStr", type = str,
                     help = "If specify this option. The working dir must contain the keyStr, \
                             if not, the program raise a warning.") 
parser.add_argument("-L","--resultLimit", nargs = '?', metavar = "num", type = int,
                     help = "By default, the program only dump last 100 lines of stdout.\
                             Use this switch to change it.") 
parser.add_argument("-dbg","--debug", action = 'store_true',
                     help = "Enable debugging.")

# parse the arguments from standard input
args = parser.parse_args()

inputfile = args.inFile[0]
outputfile = args.outFile[0]

enterBash = args.enterBash
mergeEna = args.mergeAcross
noDelForVi = args.keepViRslt
if args.dirProtector is not None:
    dirProtect = args.dirProtector
else: dirProtect = 'dummy'
if args.resultLimit is not None:
    rsltLim = args.resultLimit[0]
else: rsltLim = 100

dbg = args.debug

#TODO: build warning class, separate types of warning: dir not protected, cmd return code not 0, cmd result too long... info: added lines, ...
wrnCnt = 0
#TODO: build log file handling for this program
#----------------------------------------------------


class WorkDir:
  #class attribute (all instances of class WorkDir has same value)
 #color = wkDirColor

  #instance attribute
  def __init__ (self, name=None, row=None, col=None):
    self.name = name 
    self.row = row 
    self.col = col 


class cmd:
    #class attribute (all instances of class WorkDir has same value)
   #color = cmdColor
    if enterBash == True: choice = 'y'
    else: choice = 'no to all'
    
    #instance attribute
    def __init__ (self, name=None, row=None, col=None, wkDir=None):
        self.name = name 
        self.row = row 
        self.col = col 
        
        self.wkDir = wkDir
        self.isCollectable = True
        
        self.RsltStaRow = self.row + 1 
        self.RsltStaCol = self.col 
        self.RsltEndRow = None 
        self.RsltEndCol = None 
        self.neededLines = 0 
        
        self.returncode = 0
        self.stdoutList = []
        self.stderrList = []
        
        self.evalCmd()

    #----------------------------------------------------
    #export returncode, stdout and stderr to screen
    #-------------------
    def printCmdRslt(self):
        print("returncode = ",self.returncode)
        for line in self.stdoutList:
            print(line.decode('utf-8'))
        for line in self.stderrList:
            print(line.decode('utf-8'))
    #----------------------------------------------------
    #check inhibited commands
    #For ex: if this cmd is 'vi', we can't execute it, since it halt the whole program
    #-------------------
    def chkInhibitedCmds(self):
        global wrnCnt
       #if cmd.startswith('vi') : 
       #if any(ref in cmd for ref in inhibitCmdList):
        if any(self.name.startswith(ref) for ref in inhibitCmdList):
            self.isCollectable = False
            print ("!!WARNING!!: At row",self.row,", inhibited cmd: ",self.name)
            wrnCnt = wrnCnt +1
           #aCmd.returncode = 0
           #aCmd.stdoutList = ['']
           #aCmd.stderrList = ['']
           #aCmd.neededLines = 0
           #cmdList.pop(len(cmdList)-1)
    #----------------------------------------------------
    #
    #-------------------
    def delOldRslt(self):
       #col = aCmd.col
       #cmdName = aCmd.name
        #now, to make room for new result, delete remain content from old checklist, below the cmd:
        rowPtr = self.row +1
        cellColor = sheet.cell(rowPtr,self.col).fill.start_color.index
        while ( cellColor == rsltColor ) or ( cellColor == 'FFFFFF00'):
            if self.isCollectable and noDelForVi is True:
                pass
            else:
                sheet.cell(rowPtr,self.col).value = ''
            rowPtr = rowPtr +1
            cellColor = sheet.cell(rowPtr,self.col).fill.start_color.index
        self.RsltEndRow = rowPtr -1
        #calculate how many columns is colored in the "result" range, it is needed later:
        clredCol = self.col
        if rowPtr != self.row+1: #if result is empty, self.RsltEndCol = self.col, else:
            #check at the row below the command, count how many columns is colored with same color (may not be rsltColor variable):
            while (sheet.cell(self.row+1,clredCol+1).fill.start_color.index == sheet.cell(self.row+1,clredCol).fill.start_color.index):
                clredCol = clredCol +1
        self.RsltEndCol = clredCol
        if dbg: print("colored column = ",self.RsltEndCol)
    #----------------------------------------------------

    #----------------------------------------------------
    #execute a command, get results
    # input:  command, workingdir of said command
    # output: returncode, stdout, stderr
    #-------------------
    def runCmd(self): #.name,self.wkDir):
        cmd = self.name
        cmd = re.sub('^ll ','ls -l ',cmd)#bash does not have ll cmd, and somehow, csh does not works with subprocess
        cwd = self.wkDir
        result = subprocess.Popen([cmd],cwd=cwd,shell=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
        try: 
            returncode = result.wait(timeout=3)
            stdoutList = result.stdout.readlines()
            stderrList = result.stderr.readlines()
        except subprocess.TimeoutExpired: 
            result.terminate()
            returncode = "timeout"
            stdoutList = []
            stderrList = []
            pass
        self.returncode = returncode
        #get 100 last lines of stdout/err:
        stdoutLen = len(stdoutList)
        stderrLen = len(stderrList)
        if stdoutLen > rsltLim:
          self.stdoutList = stdoutList[-rsltLim:]
        else:
          self.stdoutList = stdoutList
        if stderrLen > rsltLim:
          self.stderrList = stderrList[-rsltLim:]
        else:
          self.stderrList = stderrList
        self.neededLines = len(self.stdoutList) + len(self.stderrList)
    
    #----------------------------------------------------
    #check if we need to change the command in case return code is not zero
    #-------------------
    def changeCmd(self):
        global wrnCnt

        if self.returncode ==0:
            return

        self.printCmdRslt()
        if self.choice == 'y' or self.choice == 'n' : 
            cmd.choice = input("Returncode not zero. Do you want to enter bash$ to find out why (y/n/yes to all/no to all)?")
            if dbg: print('choice - ',self.choice)
        
        if self.choice == 'yes to all' or self.choice == 'y' :
            print("when you want to leave bash$ and continue, type \"exit\"")
            print("!!!NOTES!!!: \"cd\" cannot be used here. Be carefull with it.")
            proceed = 0
            while (proceed == 0):
                usrIn = input("bash$ ")
                if ( usrIn == "exit"): proceed = 1
                else:
                    self.name = usrIn
                    self.runCmd() #cmd,wkDirList[i]['value'])
                    self.printCmdRslt()
            sheet.cell(self.row,self.col).value = self.name #update the command in checklist
        
        if self.returncode !=0 :
            print("!!WARNING!! returncode is",self.returncode,". Maybe there is error with above cmd.")
            wrnCnt = wrnCnt +1
    #----------------------------------------------------

    #----------------------------------------------------
    #when a command is detected, it's initial attributes are eval by this method
    #-------------------
    def evalCmd(self):
        self.chkInhibitedCmds()
        self.delOldRslt()
        #check if this cmd is 'vi' or 'vim', we can't execute this, since it halt the whole program:
        if self.isCollectable is False: return
        #Execute the command, so we have the precious result :      
        print("At row",self.row," : ",self.name)
        self.runCmd() 
        self.changeCmd() 
    #----------------------------------------------------


#----------------------------------------------------
#check if a cell contain a known command, put it in cmdList
# input:  row/column of an excel cell
#         global vars: knownCmdList, cmdColor, cmdList (list of cmd objects)
# output: True or False, updated cmdList (appended new cmdDict if necessary)
#-------------------
def isThisCellACmd(row,col,wkDir):
    #get cell value and cell color:
    cellVal = sheet.cell(row,col).value
    cellColor = sheet.cell(row,col).fill.start_color.index
    #check if the cell color match "cmdColor" and the cell start with a known command in knownCmdList:
    if (cellVal is None) or (cellColor is None) or (cellColor != cmdColor): return False
    for refCmd in knownCmdList: 
        if re.match('\s*'+refCmd,cellVal) : 
            #put the found command cell value, row and column to cmdDict:
            if dbg: print("found cmd ",cellVal,'row ',row)
            cmdObj = cmd(cellVal,row,col,wkDir)
            cmdList.append(cmdObj)
            if dbg: print(cmdList[-1].name)
            return True
    else:
        if dbg: print("not a cmd in known cmd list")
        return False
#----------------------------------------------------
    

print('Opened ',inputfile)
print('')

#get all worksheet in the file:
workbk = openpyxl.load_workbook(filename = inputfile)
sheetNameList = workbk.sheetnames

for i in sheetNameList:

    sheet = workbk[i]
    print("OPENED SHEET ",i)
    print('')
    wkDirList = [] # <----- THE LIST OF WORKING DIR IN CURRENT SHEET.
#----------------------------------------------------
#search for all working dir in the sheet
#input:  sheet, wkDirList
#output: wkDirList (appended wkDirDict if any)
#Note:   wkDirDict is a dict with keys: 
#         + value: the actual directory
#         + row/col: position of it in the sheet
#-------------------
    print('Start searching for \"workng dir\" keyword :')
    for row in range(sheet.min_row,sheet.max_row+1):
        for col in range(sheet.min_column,10):
            cellVal = sheet.cell(row,col).value
            cellColor = sheet.cell(row,col).fill.start_color.rgb

            #find cell with keyword 'working dir' (no case):
            if isinstance(cellVal, str) and ( isinstance(cellColor,str) or isinstance(cellColor,int) ) :
                if not ( re.match('working dir',cellVal,re.IGNORECASE) and (cellColor == wkDirColor) ): 
                    if dbg: print("Found Working Dir key at row :",row)
                    continue
            else: continue

            #check in the cells to the right of this row, if there's a dir there:
            for col in range(col+1,col + 10):
                wkDir = sheet.cell(row,col).value
                #check if this dir exists in the drive:
                if wkDir is not None and os.path.isdir(wkDir) : break
               #print(wkDir)
               #if os.path.isdir(wkDir): break
            else: 
                print("!!!WARNING!!! Found \"Working Dir\" key at row :",row,"but no actual Dir found!")
                wrnCnt = wrnCnt +1
                continue

            #if above loop breaks, means we found one, add it to our list to use later:
            print ("Found workdir at row:",row,"col:",col,":")
            print (wkDir)
            #check if the working dir contain the string in dirProtect
            if dirProtect != 'dummy':
                if ( not re.search(dirProtect,wkDir) ):
                    print("!!!WARNING!!! The dir does not contain \"",dirProtect,"\"!")
                    wrnCnt = wrnCnt +1
                    proceed = False
                else: proceed = True
                while proceed is False:
                    print("Please type in a new one or ignore (type \"n\" for ignore):")
                    wkDir = input()
                    if wkDir == 'n': 
                        proceed = True
                    elif not  re.search(dirProtect,wkDir)  :
                        print("!!!WARNING!!! The dir does not contain \"",dirProtect,"\"!")
                    elif not os.path.isdir(wkDir):
                        print("!!!WARNING!!! This dir does not exists!")
                    else: proceed = True

            wkDirObj = WorkDir(wkDir,row,col)
            wkDirList.append(wkDirObj)
   #print(wkDirList)
    wkDirList.append(WorkDir('dummy',sheet.max_row,sheet.max_column))
    print('')
    print('')
    print('Finished looking for working dir. Start collecting checklist!')
    print('')
    print('')
    print('')
#----------------------------------------------------

    for i in range(len(wkDirList)-1):#for all found "working dir":
        print('-----------------------------------------------------------')
        print('---COLLECTING RESULT IN',wkDirList[i].name,' ----:')
        print('')
        print('')
       #for row in range(wkDirList[i].row, wkDirList[i+1].row +1):#find cmd between two "working dir"
        row = wkDirList[i].row
        while  row < ( wkDirList[i+1].row +1 ) :
            row = row +1
            cmdList = []
            foundCmd = False
            for cmdCol in range(sheet.min_column,sheet.max_column):
                rslt = isThisCellACmd(row,cmdCol,wkDirList[i].name)
                if rslt is False: continue
                else: foundCmd = True
            if foundCmd == False: continue
            #finished looking for commands in a row. 
            if dbg: 
              for j in cmdList:
                print(j.name)
                print(j.col)
                print(j.row)
            #----------------------------------------------------
            #results of all cmds in current row are now saved in cmdList
            #now we start to dump "result" of those cmds to excel file:
            #before print the result, check if there's enough room for it:
            #TODO write better method, because current algo only works if all cmds in cmdList are in on one row
            maxNeededLines = max([cmdDict.neededLines for cmdDict in cmdList])
            maxEndClrRow = max([cmdDict.RsltEndRow for cmdDict in cmdList]) 
            availLines = maxEndClrRow - row
            if maxNeededLines > availLines:
                shiftLines = maxNeededLines - availLines
                #preseve merged cells format before adding rows
                for merged_cells in sheet.merged_cells.ranges:
                    if merged_cells.min_row > maxEndClrRow:
                        merged_cells.shift(0,shiftLines)
                #add rows
                sheet.insert_rows(maxEndClrRow+1,amount=shiftLines)
                print('At line',maxEndClrRow+1,',inserted',shiftLines,'lines.')
                #update position for all "Working Dir"
                for uprow in range(i+1,len(wkDirList)):
                    wkDirList[uprow].row = wkDirList[uprow].row + shiftLines
            #----------------------------------------------------
            
            #----------------------------------------------------
            #print out the stdout and stderr:
            for aCmd in cmdList:
               #row = aCmd['row']
                col = aCmd.col
                stdoutList = aCmd.stdoutList
                stderrList = aCmd.stderrList
                neededLines = aCmd.neededLines
                endClrRow = aCmd.RsltEndRow
                clredCol = aCmd.RsltEndCol
                if endClrRow < row + neededLines:
                    #color new rows by a fait green:
                    for clrow in range(endClrRow+1,row +neededLines +1):
                        for clcol in range(col,clredCol+1):
                            sheet.cell(clrow,clcol).fill = copy(sheet.cell(row +1,clcol).fill)
                            sheet.cell(clrow,clcol).font = copy(sheet.cell(row +1,clcol).font)
                #merge across, so conditional formating in the first line is applied to the whole line:
                if mergeEna is True :
                    for mgrow in range(row+1,row+neededLines+1):
                        sheet.merge_cells(start_row=mgrow, start_column=col, end_row=mgrow, end_column=clredCol)
                #now there's enough room, formating is beautiful. It's time to dump those result out:
                rowWrPtr = row +1
                for line in stdoutList:
                    sheet.cell(rowWrPtr,col).value = line
                    rowWrPtr = rowWrPtr +1
                for line in stderrList:
                    sheet.cell(rowWrPtr,col).value = line
                    rowWrPtr = rowWrPtr +1
            #end collecting in a row
            #----------------------------------------------------
        #end collecting in a "working dir"
        print('')
        print("-----------------------------------------------------------")
        print('')
    #end collecting in a sheet
    print('')
    print("-----------------------------------------------------------")
    print('')
    print('')
#end whole workbook
print("Completed collecting this checklist.")
print("-----------------------------------------------------------")
print('')

if os.path.exists(outputfile):
    os.remove(outputfile)
    print("WARNING! The file",outputfile,"exists. It would be replaced.")
workbk.save(outputfile)
print("Saved to file ",outputfile)

print('')
print("-----------------------------------------------------------")
print(wrnCnt,"WARNING")
print("-----------------------------------------------------------")
print('')
print("Wish you all the best with the project!")
print('')
print("-----------------------------------------------------------")



