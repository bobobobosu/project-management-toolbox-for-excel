import itertools
import json
import os
import re
from datetime import datetime
from operator import itemgetter
import arrow
import googlemaps
import intervals as I
import numbers
import numpy as np
import pytz
import xlwings as xw
from ics import Calendar, Event
from scipy.optimize import lsq_linear
from itertools import product
import TimelinePlanner

BasePath = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\'
CalendarFile = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\Share\\calendar.ics'
JSONmaps = {}
JSONmaps['TransportMap'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\TransportMap.json'
JSONmaps['taskMap'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\ValueMap.json'
JSONmaps[
    'ValueExchangeMap'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\ValueExchangeMap.json'
JSONmaps['TransportMap'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\TransportMap.json'
JSONmaps[
    'ResourceModifyTimeline'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\ResourceModifyTimeline.json'
gmaps = googlemaps.Client(key='AIzaSyDxOrE24YnVFBkeKozEAwiU1u9fsMbQHG8')


@xw.func
def TimePlannedPercentArr(plannedStartArr, plannedEndArr, StartTime):
    planned = I.empty()
    plannedPercent = []
    for i in range(0, len(plannedStartArr)):
        planned = planned.union(I.closed(plannedStartArr[i].timestamp(), plannedEndArr[i].timestamp()))

    for j in range(0, len(StartTime) - 1):
        thisTime = I.closed(StartTime[j].timestamp(), StartTime[j + 1].timestamp())
        overlap = thisTime.intersection(planned)
        plannedPercent.append([(sum([x.upper - x.lower for x in overlap if (
                isinstance(x.upper, numbers.Number) and isinstance(x.lower, numbers.Number))])) / (
                                       thisTime.upper - thisTime.lower)])
    return plannedPercent


@xw.func
def TimePercentArr(taskArr, StartArr, EndArr, StartTime, EndTime):
    uniqueTask = list(set(taskArr))
    resultDict = {i: 0 for i in uniqueTask}
    for task in uniqueTask:
        resultDict[task] = TimePlannedPercent([val for idx, val in enumerate(StartArr) if taskArr[idx] == task],
                                                      [val for idx, val in enumerate(EndArr) if taskArr[idx] == task],
                                                      StartTime, EndTime)

    if '時序專案(288)' in resultDict:
        resultDict['時序專案(288)'] += 1 - sum([val for key, val in resultDict.items()])
    else:
        resultDict['時序專案(288)'] = 1 - sum([val for key, val in resultDict.items()])

    return [[key, val] for key, val in resultDict.items()]


@xw.func
def TimePlannedPercent(plannedStartArr, plannedEndArr, StartTime, EndTime):
    if not isinstance(plannedStartArr, list):
        return 0
    planned = I.empty()
    for i in range(0, len(plannedStartArr)):
        planned = planned.union(I.closed(plannedStartArr[i], plannedEndArr[i]))

    thisTime = I.closed(StartTime, EndTime)
    overlap = thisTime.intersection(planned)
    return (sum([x.upper - x.lower for x in overlap if
                 (isinstance(x.upper, numbers.Number) and isinstance(x.lower, numbers.Number))])) / (
                   thisTime.upper - thisTime.lower)


@xw.func(async_mode='threading')
def generateCalendar(startIndex, titleRange, dataRange):
    c = Calendar()

    for i in range(int(startIndex), len(dataRange)):
        if dataRange[i][titleRange.index('ID')] == 5073:
            hbyguh = 0
        e = Event(uid=str(dataRange[i][titleRange.index('ID')]))
        e.name = dataRange[i][titleRange.index('Subject')]
        e.duration = dataRange[i][titleRange.index('實際耗時')]
        e.location = dataRange[i][titleRange.index('Location')]
        e.begin = arrow.get(
            pytz.timezone(dataRange[i][titleRange.index('時區')]).localize(dataRange[i][titleRange.index('Start Date')]))
        e.description = '\n'.join([title + '\n\t' + str(dataRange[i][titleRange.index(title)]) for title in titleRange])
        # e.end = arrow.get(dataRange[i][titleRange.index('End Date')],tzinfo=dataRange[i][titleRange.index('時區')])
        c.events.add(e)
    with open(CalendarFile, 'w', encoding="utf-8") as f:
        f.writelines(c)
    return 'Calendar saved at ' + str(datetime.now())


@xw.func
@xw.ret(expand='table')
def returnoriginal(x):
    return [[i] for i in x]


def wrapper(func, args, res):
    res.append(func(*args))

@xw.func
def Range2Json(rng1):
    return json.dumps(rng1)
@xw.func
def Json2ArrPY(rng1):
    return json.loads(rng1)

@xw.func(async_mode='threading')
def Table2Json(name, dataRange, titleRange):
    tableDict = {}
    for i in range(0, len(titleRange)):
        tableDict[titleRange[i]] = [row[i] for row in dataRange]
    with open(os.path.join(BasePath, name + '.json'), 'w+') as outfile:
        json.dump(tableDict, outfile, indent=2)
    return 0


@xw.func(async_mode='threading')
def Timeline2Json(name, dataRange, titleRange, TitleOverride=None):
    tableDict = {}
    tableDict['titleRange'] = titleRange if TitleOverride is None else [
        (i if i not in TitleOverride[1] else TitleOverride[0][TitleOverride[1].index(i)]) for i in titleRange]
    tableDict['dataRange'] = {}
    for i in range(0, len(dataRange)):
        tableDict['dataRange'][dataRange[i][titleRange.index('ID')]] = dataRange[i]
    with open(os.path.join(BasePath, name + '.json'), 'w+') as outfile:
        json.dump(tableDict, outfile, indent=2, default=json_encode)
    return name + ' saved at ' + str(datetime.now())


@xw.func(async_mode='threading')
def Timeline2JsonSelected(name, idRange, dataRange, titleRange, TitleOverride=None):
    tableDict = {}
    tableDict['titleRange'] = titleRange if TitleOverride is None else [
        (i if i not in TitleOverride[1] else TitleOverride[0][TitleOverride[1].index(i)]) for i in titleRange]
    tableDict['dataRange'] = {}
    for i in range(0, len(idRange)):
        if idRange[i] is not None:
            tableDict['dataRange'][int(idRange[i])] = dataRange[i]
    with open(os.path.join(BasePath, name + '.json'), 'w+') as outfile:
        json.dump(tableDict, outfile, indent=2, default=json_encode)
    return name + ' saved at ' + str(datetime.now())


@xw.func
def getTimelineDataFile(name, id, title):
    with open(os.path.join(BasePath, name + '.json')) as json_data:
        tableDict = json.load(json_data)
    return getTimelineData(tableDict, id, title)


def getTimelineData(tableDict, id, title):
    return tableDict['dataRange'][str(int(id))][tableDict['titleRange'].index(title)]


@xw.func
def planAvalible(numArr, titleArr, startArr, endArr, movable, locationArr, currresourceJSON, accuresourceJSON, mode):
    ValueExchangeMap = UpdateLocValueExchangeMap(list(set(locationArr)))
    currresourceDict = []
    accuresourceDict = []
    for i in currresourceJSON:
        try:
            currresourceDict.append(json.loads(i))
        except:
            currresourceDict.append({})
    for i in accuresourceJSON:
        try:
            accuresourceDict.append(json.loads(i))
        except:
            accuresourceDict.append({})
    locationArr = [i.split('[', 1)[1].split(']')[0] for i in locationArr]

    mandatoryTimmedTasks = []
    for idx, mandTime in enumerate(startArr):
        startLoc = endLoc = locationArr[idx]
        # Location Overridebu[idx] in ValueExchangeMap:
        if titleArr[idx] in ValueExchangeMap:
            if 'moveto' in ValueExchangeMap[titleArr[idx]]:
                startLoc = ValueExchangeMap[titleArr[idx]]['location']
                endLoc = ValueExchangeMap[titleArr[idx]]['moveto']
        mandatoryTimmedTasks.append((
            int(startArr[idx].timestamp() / 60),
            int(endArr[idx].timestamp() / 60),
            startLoc,
            endLoc,
            movable[idx],
            ([(('-=', (key,), -int(val)) if val < 0 else ('+=', (key,), int(val))) for key, val in
              currresourceDict[idx].items()] if idx != 0 else [])
        ))

    initialState = [('=', (key,), (val if val > 0 else 0)) for key, val in accuresourceDict[0].items()] + \
                   [('=', ('t',), int(startArr[0].timestamp() / 60))] + \
                   [('at', locationArr[0])] + \
                   [('=', ('mand',), 0)]

    try:
        rawplan = TimelinePlanner.calculatePlan(initialState, mandatoryTimmedTasks, 1000000000, ValueExchangeMap, mode)
        result = []
        num = 0
        for plan in rawplan:
            if '#' in plan[0]:
                num = int(plan[0].replace('#', ''))
            else:
                thisresult = {}
                thisresult['編號'] = numArr[num] + 1
                thisresult['交易物件'] = plan[0]
                thisresult['預計耗時'] = plan[1] / (60 * 24)
                thisresult['Location'] = '[' + plan[2] + ']'
                thisresult['預計百分比'] = 1
                thisresult['起始百分比'] = 0
                result.append(thisresult)
        return json.dumps(result)
    except:
        raise
        # return 'No Plan'


@xw.func
def CurrentResourceArr(title, resourceCurrDeli):
    result = []
    resourceCurrList = [(([int(j) for j in i.split(',')]) if type(i) == str else []) for i in resourceCurrDeli]

    for resourceCurr in resourceCurrList:
        resource = {}
        if len(title) <= len(resourceCurr):
            for row in range(0, len(title)):
                resource[title[row]] = resourceCurr[row]
        result.append([json.dumps(resource)])
    return result


@xw.func
def supplyArr(taskArr, thisDelta, id, resourceNamesArr):
    id = [int(id[i] if id[i] is not None else 0) for i in range(0, len(id))]
    idindexDict = {int(val): idx for idx, val in enumerate(id) if val is not None}
    depenList = []
    supplyDict = {i: [] for i in id}
    thisDeltadict = []
    resourceDict = {i: [] for i in resourceNamesArr}
    resourceValDict = {i: 0 for i in resourceNamesArr}
    AccuResourceArr = []
    BuffArr = []
    Needed = []

    for i in range(0, len(thisDelta)):
        try:
            thisDeltadict.append(json.loads(thisDelta[i]))
        except:
            thisDeltadict.append({})

    for i in range(0, len(thisDeltadict)):
        if id[i] == 4857:
            de2=thisDeltadict[i]
            wefr = 0

        supplyIDList = []

        neededList = [key + ':' + str(val + resourceValDict[key]) for key, val in thisDeltadict[i].items() if
                      (val + resourceValDict[key]) < 0 and thisDeltadict[i][key] < 0]
        if len(neededList) > 0:
            Needed.append(str(neededList))
        else:
            Needed.append('')

        if taskArr[i] == '存取權歸零':
            for key, val in thisDeltadict[i].items():

                if val > 0:
                    resourceDict[key]=([[id[i], val]])
                    resourceValDict[key] = val
                elif val==0:
                    resourceDict[key] =[]
                    resourceValDict[key] =0
                else:
                    resourceValDict[key]=val


        else:
            if (id[i] == 5418):
                sss = 0
            for key, val in thisDeltadict[i].items():
                if val > 0:
                    resourceDict[key].extend([[id[i], val]])
                    resourceValDict[key] += val
            for key, val in thisDeltadict[i].items():
                if val < 0:
                    resourceValDict[key] += val
                    while len(resourceDict[key]) > 0 and -val > 0:
                        if resourceDict[key][0][1] > 0:
                            avalibleVal = resourceDict[key][0][1]
                            resourceDict[key][0][1] += val
                            val += avalibleVal
                            supplyIDList.append(resourceDict[key][0][0])
                            supplyDict[resourceDict[key][0][0]].append(id[i])
                        else:
                            resourceDict[key].pop(0)

        supplyIDList = list(set(supplyIDList))
        BuffArr.append(id[min([idindexDict[j] for j in supplyIDList])] if len(supplyIDList) > 0 else id[0])
        depenList.append(json.dumps(list(set(supplyIDList))))
        AccuResourceArr.append(json.dumps(resourceValDict))
    supplyList = [json.dumps(list(set(supplyDict[i])) if i is not None else json.dumps([])) for i in id]
    return [[supplyList[i], depenList[i], (AccuResourceArr[i]), BuffArr[i], Needed[i]] for i in
            range(0, len(id))]

@xw.func
def CalcCurrResourceArr(taskArr, completeArr, idArr):
    CurrResourceArr = []
    with open(JSONmaps['ValueExchangeMap']) as json_data:
        ValueExchangeMap = json.load(json_data)
    with open(JSONmaps['ResourceModifyTimeline']) as json_data:
        ResourceModifyTimeline = json.load(json_data)
    for row in range(0, len(taskArr)):
        if idArr[row] == 5293:
            wefr = 0

        resourceCurrDict = {}
        # apply modify timeline
        if idArr[row] is not None:
            if str(int(idArr[row])) in ResourceModifyTimeline['dataRange']:
                title = ResourceModifyTimeline['titleRange']
                data = ResourceModifyTimeline['dataRange']
                resourceCurrDict = {title[i]: data[str(int(idArr[row]))][i] for i in range(0, len(title)) if
                                    data[str(int(idArr[row]))][i] is not None}

        # apply plusminus
        if taskArr[row] in ValueExchangeMap and completeArr[row] == True:
            for ResourceName, ResourceVal in ValueExchangeMap[taskArr[row]]['in'][0].items():
                if ResourceName in resourceCurrDict:
                    resourceCurrDict[ResourceName] += ResourceVal
                else:
                    resourceCurrDict[ResourceName] = ResourceVal
            for ResourceName, ResourceVal in ValueExchangeMap[taskArr[row]]['out'][0].items():
                if ResourceName in resourceCurrDict:
                    resourceCurrDict[ResourceName] += ResourceVal
                else:
                    resourceCurrDict[ResourceName] = ResourceVal
        else:
            pass
        CurrResourceArr.append([json.dumps(resourceCurrDict)])
    return CurrResourceArr


@xw.func
def AccuResource2Table(AccuResource, title):
    rows = []
    for taskaccu in AccuResource:
        try:
            resourceCurr = json.loads(taskaccu)
            rows.append([(resourceCurr[key] if key in resourceCurr else 0) for key in title])
        except:
            rows.append([0 for i in range(0, len(title))])
    return rows


@xw.func
def AccuResource(thisDelta, prev):
    try:
        resourceCurr = json.loads(thisDelta)
        resourcePrev = json.loads(prev)

        for key, val in resourceCurr.items():
            if key in resourcePrev:
                resourceCurr[key] += resourcePrev[key]

        return json.dumps(resourceCurr)
    except:
        return thisDelta

@xw.func
def genTransportMap(location,force_update = False):
    g = [re.search(r"\[(\w+)\]", i) for i in location]
    try:
        with open(JSONmaps['TransportMap']) as json_data:
            TransportMap = json.load(json_data)
    except:
        TransportMap = {}


    location = [j for j in list(set([i.split('[', 1)[1].split(']')[0] for i in location])) if
                j not in ['Undefined', 'Moving']]

    permutations = list(itertools.permutations(location, 2))
    if force_update == False:
        permutations = [i for i in permutations if i not in [(val['location'],val['moveto']) for key,val in TransportMap.items()] ]

    for transport in permutations:
        TransportMap['MOVING:' + transport[0] + '->' + transport[1]] = {}
        TransportMap['MOVING:' + transport[0] + '->' + transport[1]]['in'] = [{}]
        TransportMap['MOVING:' + transport[0] + '->' + transport[1]]['out'] = [{}]
        TransportMap['MOVING:' + transport[0] + '->' + transport[1]]['time'] = int(
            getTransport(transport[0], transport[1]) * 1440)
        TransportMap['MOVING:' + transport[0] + '->' + transport[1]]['moveto'] = transport[1]
        TransportMap['MOVING:' + transport[0] + '->' + transport[1]]['location'] = transport[0]

    with open(JSONmaps['TransportMap'], 'w') as outfile:
        json.dump(TransportMap, outfile, indent=2)
    return TransportMap

@xw.func
def UpdateLocValueExchangeMap(locations):
    with open(JSONmaps['ValueExchangeMap']) as json_data:
        ValueExchangeMap = json.load(json_data)
    TransportMap = genTransportMap(locations)
    with open(JSONmaps['ValueExchangeMap'], 'w') as outfile:
        json.dump({**ValueExchangeMap, **TransportMap}, outfile, indent=2)
    return ValueExchangeMap

@xw.func()
def genValueExchangeMap(task, header, data, alltask, allestimatedTime, alllocation):
    allTaskData = {}
    for row in range(0, len(alltask)):
        allTaskData[alltask[row]] = {}
        allTaskData[alltask[row]]['time'] = allestimatedTime[row]
        allTaskData[alltask[row]]['location'] = alllocation[row]

    ValueExchangeMap = {}
    for row in range(0, len(task)):
        inList = {}
        outList = {}
        for idx, val in enumerate(data[row]):
            if float(val) < 0:
                inList[header[idx]] = val
            if float(val) > 0:
                outList[header[idx]] = val

        # inList = [{header[idx]: val} for idx, val in enumerate(data[row]) if float(val) < 0]
        # outList = [{header[idx]: val} for idx, val in enumerate(data[row]) if float(val) > 0]
        if not task[row] in ValueExchangeMap:
            ValueExchangeMap[task[row]] = {}
            ValueExchangeMap[task[row]]['in'] = [inList]
            ValueExchangeMap[task[row]]['out'] = [outList]  # + [{task[row]: 1}]
            ValueExchangeMap[task[row]]['time'] = int(allTaskData[task[row]]['time'] * 1440)
            ValueExchangeMap[task[row]]['location'] = allTaskData[task[row]]['location'].split('[', 1)[1].split(']')[0]
        else:
            ValueExchangeMap[task[row]]['in'] += [inList]
            ValueExchangeMap[task[row]]['out'] += [outList]
            ValueExchangeMap[task[row]]['location'] = allTaskData[task[row]]['location'].split('[', 1)[1].split(']')[0]

    TransportMap = genTransportMap(alllocation)
    with open(JSONmaps['ValueExchangeMap'], 'w') as outfile:
        json.dump({**ValueExchangeMap, **TransportMap}, outfile, indent=2)
    return 'ValueExchangeMap saved at ' + str(datetime.now())

@xw.func
def getTransport(addressFrom, addressTo):
    if addressFrom == addressTo or addressTo == '' or addressFrom == '':
        return 0

    timedistance = \
        gmaps.distance_matrix(addressFrom, addressTo, 'bicycling')['rows'][
            0]['elements'][0]
    time = timedistance['duration']['value'] / (60 * 60 * 24)
    if time > 30 * (1 / (24 * 60)):
        timedistance = \
            gmaps.distance_matrix(addressFrom, addressTo, 'driving')[
                'rows'][
                0]['elements'][0]
        time = timedistance['duration']['value'] / (60 * 60 * 24)
    return time


@xw.func
def genSUMINEstimate(fromarr, toarr, coeff, realsumin, tasks):
    A = np.zeros(shape=(len(fromarr), len(tasks)))
    b = np.zeros(len(fromarr))
    for row in range(0, len(fromarr)):
        if fromarr[row] in tasks and toarr[row] in tasks:
            A[row][tasks.index(fromarr[row])] = coeff[row]
            A[row][tasks.index(toarr[row])] = -1

    for row in range(0, len(tasks)):
        if realsumin[row] > 0:
            newA = np.zeros(shape=(1, len(tasks)))
            newA[0][row] = 1
            newb = [realsumin[row]]
            A = np.concatenate((A, newA))
            b = np.concatenate((b, newb))

    suminbounds = (np.zeros(len(tasks)), np.full(len(tasks), np.inf))
    suminbounds[0][tasks.index('SU-MIN')] = 1
    suminbounds[1][tasks.index('SU-MIN')] = 1.0000001
    res = lsq_linear(A, b, bounds=suminbounds, lsmr_tol='auto', verbose=1).x
    g = [[x] for x in res]
    return [[x] for x in res]


@xw.func
def add_one(data):
    return [1, 2, 3]


@xw.func
@xw.arg('x', np.array, ndim=2)
@xw.arg('y', np.array, ndim=2)
def matrix_mult(x, y):
    return x - y


@xw.func
def getParent(task, map, filter=''):
    with open(JSONmaps[map]) as json_data:
        dictmap = JSON2Dict(json.load(json_data))
    if not task in dictmap: return 'N/A'
    filteredAncestors = []
    for ancestors in dictmap[task]['ancestors']:
        filteredAncestors += [i for i in ancestors if i.find(filter) != -1 and i not in filteredAncestors]
    # return filteredAncestors[-1] if len(filteredAncestors) > 0 else 'N/A'
    return ";".join(reversed(filteredAncestors)) if len(filteredAncestors) > 0 else 'N/A'


# TaskMapHelpers
def dfsBuildDict(node, dictionary, ancestor):
    if not node['node'] in dictionary: dictionary[node['node']] = {}
    if not 'directchildren' in dictionary[node['node']]: dictionary[node['node']]['directchildren'] = []
    if not 'ancestors' in dictionary[node['node']]: dictionary[node['node']]['ancestors'] = []
    dictionary[node['node']]['directchildren'] = list(
        set(dictionary[node['node']]['directchildren'] + [i['node'] for i in node['children']]))
    dictionary[node['node']]['ancestors'] += [ancestor]
    for child in node['children']:
        dfsBuildDict(child, dictionary, ancestor + [node['node']])


def JSON2Dict(jsonData):
    dictionary = {}
    dfsBuildDict(jsonData, dictionary, [])
    return dictionary


def is_number(n):
    try:
        float(n)  # Type-casting the string to `float`.
        # If string is not a valid `float`,
        # it'll raise `ValueError` exception
    except ValueError:
        return False
    return True


def json_encode(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError("Type %s not serializable" % type(obj))


if __name__ == '__main__':
    xw.serve()
