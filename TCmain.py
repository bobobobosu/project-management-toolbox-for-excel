import multiprocessing
import numbers
import os
import time
from datetime import datetime
from multiprocessing import Pool
from win10toast import ToastNotifier
import xlwings as xw
import os
from shutil import copyfile
currentPath = os.path.dirname(os.path.abspath(__file__))
xb = xw.Book(os.path.join(currentPath,'TC.xlsb'))
xb.sheets['趨勢'].range('R2').value = 2

core = int(xb.sheets['趨勢'].range('R2').value)
cores = [str(i)+'.xlsb' for i in range(1,core+1)]
xbn=[]
toaster = ToastNotifier()


def llServer(interval):
    global xbn
    global xb
    global toaster

    xb.sheets['趨勢'].range('V2').value = ""
    xb.sheets['趨勢'].range('L2').value = ""


    while True:
        
        # xb.macro('UpdateThisTask')(False)
        # behind = xb.macro('CheckBehind')()
        # if behind >1:
        #     toaster.show_toast(str(behind)+"behind",
        #                             icon_path=None,
        #                             duration=0.5,
        #                             threaded=True)
        time.sleep(interval)
        # xb.save()
        if xb.sheets['趨勢'].range('O2').value ==0 :
            continue

        if not xb.sheets['趨勢'].range('V2').value is None:
            if xb.sheets['趨勢'].range('V2').value == "[Refresh]":
                runCores()
            else:
                print(">Sync")
                print(xb.sheets['趨勢'].range('V2').value)
                current = xb.sheets['趨勢'].range('V2').value
                if (current is None) : continue
                try:
                    listAreas = current.replace('|',',').split(',')
                except:
                    listAreas = [current]
                if len(listAreas) > 0:
                    for area in listAreas:
                        for xbs in xbn:
                            xbs.macro('InstanceSyncUpdateValue')(area)
                            xbs.macro('SetTempSheet')()
                print("-------"+str(datetime.now())+"\n")
            xb.sheets['趨勢'].range('V2').value = ""

        if not xb.sheets['趨勢'].range('L2').value is None:
            print(">Cal")
            print(xb.sheets['趨勢'].range('L2').value)
            current = xb.sheets['趨勢'].range('L2').value
            xb.sheets['趨勢'].range('L2').value = ""
            if (current is None) : continue
            try:
                listAreas = current.replace('|',',').split(',')
            except:
                listAreas = [current]
            print(listAreas)


            tasks = []
            for i in range(core):
                if len(listAreas) ==1:
                    if xw.Range(listAreas[0]).columns.count == 1:
                        if xw.Range(listAreas[0]).count >1:
                            address = xw.Range(listAreas[0]).get_address(True, False, True)

                            rowStartori = int(address.split('!')[1].split(':')[0].split('$')[1])
                            rowStart = int(address.split('!')[1].split(':')[0].split('$')[1])-1
                            rowEnd = int(address.split('!')[1].split(':')[1].split('$')[1])
                            cellsCount = rowEnd-rowStart+1

                            newStart = rowStart+(1+int(cellsCount/core))*(i % core)+1
                            newEnd = rowStart+(1+int(cellsCount/core))*((i % core)+1)
                            if newEnd > rowEnd: newEnd= rowEnd

                            newaddress = address.replace(str(rowStartori),str(int(newStart))).replace(str(rowEnd),str(int(newEnd)))
                            tasks.append([i, [newaddress]])
                        else:
                            tasks.append([i, [xw.Range(listAreas[0]).get_address(True, False, True)]])
                    else:
                        tasks.append([i, [j for j in listAreas if listAreas.index(j) % core == i]])
                else:
                    tasks.append([i,[j for j in listAreas if listAreas.index(j) % core == i]])

            if len(tasks) > 0:
                with Pool(core) as p:
                    p.map(calculatell, tasks)
            # xb.macro('SetTempSheet')()
            ToastNotifier().show_toast("Calculation",
                               "Done!",
                                icon_path=None,
                                duration=2,
                                threaded=True)
            print("-------"+str(datetime.now())+"\n")


def calculatell(args):
    coreBook = xw.Book(os.path.join(currentPath,cores[args[0]]))
    if (len(args[1])==0): return
    toCal = ",".join(args[1])
    print("Calculating " +str(toCal) +" using: " + coreBook.name)
    coreBook.macro('InstanceSyncUpdateFormula')(toCal)
    coreBook.macro('RefreshCalAdd')(toCal)
    coreBook.macro('InstanceSyncBackFormula')(toCal)
    coreBook.macro('SetTempSheet')()


def getBookByName(name):
    for iapp in xw.apps:
        for workbook in iapp.books:
            if workbook.name == name:
                return workbook
    return None
def newExcelInstance():
    app = xw.App(visible=False)
    app.calculation = 'manual'
    return app

  
def saveCores(opened):
    avaliblecores = []
    for openedcores in opened:
        try:
            getBookByName(openedcores).app.books.add()
            avaliblecores.append(getBookByName(openedcores).app.pid)
            getBookByName(openedcores).close()
            print(openedcores+" closed")
        except:
            x=0
    # xb.macro('SaveCores')(xb.sheets['趨勢'].range('R2').value)
    xb.save()
    for corefiles in opened:
        copyfile(os.path.join(currentPath,'TC.xlsb'), os.path.join(currentPath,corefiles))
    return avaliblecores


def openCores(args):
    workbookName =str(args[1])
    app = xw.apps[args[0]]
    app.books.open(os.path.join(currentPath,workbookName))
    for workbook in [i.name for i in app.books if i.name != workbookName]:
        app.books[workbook].close()


def reopenCores(toopen,avaliblecores):
    for i in range(len(toopen)-len(avaliblecores)):
        avaliblecores.append(newExcelInstance().pid)

    print('Opening '+str(toopen))
    print('PID '+str(avaliblecores))

    opening = []
    for index,book in enumerate(toopen):
        opening.append([avaliblecores[index],book])
    with Pool(len(toopen)) as p:
        p.map(openCores, opening)


def setCores(cores):
    global xbn
    xbn=[]
    for workbook in cores:
        workbookobj = xw.Book(os.path.join(currentPath,workbook))
        xbn.append(workbookobj)
        workbookobj.macro('SetTempSheet')()
    ToastNotifier().show_toast("Refresh",
                       "Done!",
                       icon_path=None,
                       duration=2,
                       threaded=True)

def runCores():
    opened =[j for j in cores if getBookByName(j) is not None]
    notopened =[i for i in cores if i not in opened]
    print('Opened: '+str(opened))
    print('Not Opened: '+str(notopened))
    reopenCores(cores,saveCores(cores))
    setCores(cores)


if __name__ == '__main__':
    xb.sheets['趨勢'].range('R2').value=2
    runCores()
    llServer(1)


