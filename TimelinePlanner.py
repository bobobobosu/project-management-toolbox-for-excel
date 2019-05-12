from __future__ import print_function

import copy
import json
import time

from planner import planner
from pyddl import Domain, Problem, Action, neg

JSONmaps = {}
JSONmaps['taskMap'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\ValueMap.json'
JSONmaps[
    'ValueExchangeMap'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\ValueExchangeMap.json'
JSONmaps['TransportMap'] = 'C:\\_DATA\\_Storage\\_Sync\\Devices\\root\\Documents\\root\\Notebook\\TransportMap.json'


def getResourcesFromExchangeMap(ValueExchangeMap):
    thisResources = []
    for key, thisTask in ValueExchangeMap.items():
        thisResources += [j for i in thisTask['in'] for j in i] + [j for i in thisTask['out'] for j in i]
    return list(set(thisResources))


def problem(verbose, initialState, endState, ValueExchangeMap, mandatoryTimmedTasks, deadline, timeout, mode=1):
    timeoutStart = int(round(time.time() * 1000))
    tasks = tuple([key for key, val in ValueExchangeMap.items()])
    listOfLocations = list(set([i['location'] for key, i in ValueExchangeMap.items()]))
    taskactionsList = []
    mandatoryTasksList = []
    for idx, mandatoryTasks in enumerate(mandatoryTimmedTasks):
        startTime = mandatoryTasks[0]
        endTime = mandatoryTasks[1]
        startLoc = mandatoryTasks[2]
        endLoc = mandatoryTasks[3]
        movable = mandatoryTasks[4]
        effectResources = mandatoryTasks[5]
        parameters = [
            # ('location', 'l'),
        ]

        # ([('<=', ('t',), startTime)] if not movable else []) + \
        preconditions = [('>=', i[1], i[2]) for i in effectResources if i[0] == '-='] + \
                        ([('<=', ('t',), startTime)] if not movable else []) + \
                        [('=', ('mand',), idx)] + \
                        ([('at', startLoc)] if startLoc not in ['Undefined', 'Moving'] else [])
        effects = effectResources + \
                  ([('==', ('t',), endTime)] if not movable else [('+=', ('t',), endTime - startTime)]) + \
                  [('+=', ('mand',), 1)] + \
                  ([(neg(('at', loc)) if loc != endLoc else ('at', endLoc)) for loc in
                    listOfLocations] if endLoc not in ['Undefined', 'Moving'] else [])
        thisaction = Action(
            str('#' + str(idx)),
            parameters=tuple(parameters),
            preconditions=tuple(preconditions),
            effects=tuple(effects))
        taskactionsList.append(thisaction)
        mandatoryTasksList.append(effectResources)

    for task in tasks:
        for idxin, inMode in enumerate(ValueExchangeMap[task]['in']):
            for idxout, outMode in enumerate(ValueExchangeMap[task]['out']):
                location = ValueExchangeMap[task]['location']
                parameters = [
                    # ('location', 'l'),
                ]
                preconditions = [('<=', ('t',), int(deadline - int(ValueExchangeMap[task]['time'])))] + \
                                [('>=', (key,), -int(inMode[key])) for key, val in inMode.items()] + \
                                [('>=', ('t',), 0)] + \
                                ([('at', location)] if location not in ['Undefined', 'Moving'] else [])
                effects = [('+=', (key,), int(inMode[key])) for key, val in inMode.items()] + \
                          [('+=', (key,), int(outMode[key])) for key, val in outMode.items()] + \
                          [('+=', ('t',), ValueExchangeMap[task]['time'])] + \
                          ([(neg(('at', loc)) if loc != ValueExchangeMap[task]['moveto'] else (
                              'at', ValueExchangeMap[task]['moveto'])) for loc in listOfLocations if
                            loc not in ['Undefined', 'Moving']] if 'moveto' in ValueExchangeMap[task] else [])
                thisaction = Action(
                    (str(task) if (idxin + idxout == 0) else (str(task) + ' [' + str(idxin) + '_' + str(idxout) + ']')),
                    parameters=tuple(parameters),
                    preconditions=tuple(preconditions) if len(preconditions) > 0 else ((),),
                    effects=tuple(effects))
                taskactionsList.append(thisaction)

    domain = Domain(tuple(taskactionsList))
    problem = Problem(
        domain,
        {
            # 'location': tuple(listOfLocations)
        },
        init=initialState,
        goal=endState
    )

    def checkCompleteLevel(state, alllevels):
        statecopy = copy.deepcopy(state.f_dict)
        planLevel = state.f_dict[('mand',)]
        realLevel = planLevel
        progressAbovePlanLevel = 0

        if planLevel == len(mandatoryTasksList): return planLevel, 0
        for i in range(planLevel, (len(mandatoryTasksList) if alllevels else (planLevel + 1))):
            for req in mandatoryTasksList[i]:
                if req[0] == '-=' and req[1] in statecopy:
                    if statecopy[req[1]] > req[2]:
                        progressAbovePlanLevel += req[2]
                    else:
                        progressAbovePlanLevel += statecopy[req[1]]

            # apply mandatoryTasks and check
            for req in mandatoryTasksList[i]:
                if req[0] == '-=':
                    statecopy[req[1]] -= req[2]
                elif req[0] == '+=':
                    statecopy[req[1]] += req[2]
            if not all(
                    [(val >= 0) for key, val in statecopy.items() if key in [j[1] for j in mandatoryTasksList[i]]]):

                break
                # continue
            else:
                realLevel = i + 1

        return realLevel, progressAbovePlanLevel

    def checkCompletedProgress(planLevel):
        prevProgress = 0
        for i in range(0, planLevel):
            prevProgress += sum([(i[2] if i[2] > 0 else 0) for i in mandatoryTasksList[i]])
        return prevProgress

    def getPlanByState(state):
        tmpstate = state
        tmpplan = []
        while tmpstate.predecessor is not None:
            tmpplan.append(tmpstate.predecessor[1].name)
            tmpstate = tmpstate.predecessor[0]
        return list(reversed(tmpplan))

    def dependancy_heuristic(state):
        if int(round(time.time() * 1000)) - timeoutStart > timeout: raise

        planLevel = state.f_dict[('mand',)]
        if planLevel == 0: return 0

        # Current Plan
        planCurrent = getPlanByState(state)

        # Progress to next level
        realLevel, realProgressOnThisLevel = checkCompleteLevel(state, alllevels=False)

        # Progress to all following levels
        realLevel, realProgressOnAllLevel = checkCompleteLevel(state, alllevels=True)

        excessT = ''
        # Check if time of this level exceeds limit
        if planLevel > 0 and planLevel < len(mandatoryTimmedTasks):
            progressAbovePlanTime = sum(
                [ValueExchangeMap[key]['time'] for key in
                 planCurrent[planCurrent.index('#' + str(planLevel - 1)) + 1:]])
            excessT = planCurrent[:planCurrent.index('#' + str(planLevel - 1))]
            if progressAbovePlanTime > (
                    mandatoryTimmedTasks[planLevel][0] - mandatoryTimmedTasks[planLevel - 1][1]):
                realLevel -= 1
                planLevel -= 1
                realProgressOnThisLevel = 0
                realProgressOnAllLevel = 0

        # Sum of planned tasks
        planned = sum(
            [1 for key in
             planCurrent[:planCurrent.index('#' + str(planLevel - 1)) + 1] if
             key.startswith('MOVING')]) if planLevel > 0 else 0

        # Sum of resolved resource
        prevProgress = checkCompletedProgress(planLevel)

        if mode == 1:
            # Early the better
            dist = (planLevel + planned) + prevProgress + realProgressOnAllLevel
        elif mode == 2:
            # Early the better Aggressive
            dist = planLevel + (prevProgress + realProgressOnAllLevel) * 2
        elif mode == 3:
            # Faster the better
            dist = (planned + planLevel) * 100000 + (realProgressOnThisLevel)
        else:
            pass

        print(
            str(dist) + ' ' + str(planLevel) + ' ' + str(realLevel) + ' ' + str(prevProgress) + ' ' + str(
                realProgressOnThisLevel) + ' ' + str(realProgressOnAllLevel) + ' ' + str(planCurrent)
        )
        # + ' ' + str(getPlanByState(state))
        return -dist

    plan = planner(problem, heuristic=dependancy_heuristic, verbose=verbose)
    resultPlan = []
    loc = 'Undefined'
    for action in plan:
        duration = 0
        for eff in action.add_effects:
            if eff[0] == 'at': loc = eff[1]
        for numeff in action.num_effects:
            if numeff[0] == (('t'),): duration = numeff[1]
        resultPlan.append([action.name, duration, (loc if (not action.name.startswith('MOVING:')) else 'Moving')])

    return resultPlan


def calculatePlan(initialState, mandatoryTimmedTasks, timeout, ValueExchangeMap, mode=1000):
    endState = [('=', ('mand',), len(mandatoryTimmedTasks))]
    plan = problem(True, tuple(initialState), tuple(endState), ValueExchangeMap, mandatoryTimmedTasks,
                   max(i[0] for i in mandatoryTimmedTasks), timeout, mode)
    if plan is None:
        print('No Plan!')
    else:
        for action in plan:
            print(action)
    return plan


if __name__ == '__main__':
    with open(JSONmaps['ValueExchangeMap']) as json_data:
        ValueExchangeMap = json.load(json_data)

    # Parse arguments
    resources = getResourcesFromExchangeMap(ValueExchangeMap)

    mandatoryTimmedTasks = [(0, 500, '10411 Flora Vista Ave', '10411 Flora Vista Ave', False, []),
                            (4000, 4000, '10411 Flora Vista Ave', '10411 Flora Vista Ave', False,
                             []),
                            (6000, 6000, '10411 Flora Vista Ave', '10411 Flora Vista Ave', False,
                             []),
                            (7000, 7000, '10411 Flora Vista Ave', '10411 Flora Vista Ave', False,
                             []),
                            (10000, 10000, 'De Anza, Main Campus', '10411 Flora Vista Ave', False,
                             [('-=', ('r.PHYS Quiz Prep_M',), 4), ])]
    initialState = [('=', (key,), 0) for key in resources] + \
                   [('=', ('t',), 0)] + \
                   [('at', '10411 Flora Vista Ave')] + \
                   [('=', ('mand',), 0)]
    endState = [('=', ('mand',), len(mandatoryTimmedTasks))]
    plan = problem(True, tuple(initialState), tuple(endState), ValueExchangeMap, mandatoryTimmedTasks,
                   max(i[0] for i in mandatoryTimmedTasks), timeout=1000000000, mode=2)
    if plan is None:
        print('No Plan!')
    else:
        for action in plan:
            print(action)
    x = 0
