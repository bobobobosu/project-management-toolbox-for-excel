import ctypes
import json
import os
import re
import bisect
from urllib import pathname2url

from java.net import URL
from javax.imageio import ImageIO

SEPERATOR = "___"
SEPERATOR2 = "=>"

# Attributes
Attr_relate = "relate"
Attr_Active = "Active"
Attr_InActive = "InActive"
Attr_LastModified = "LastModified"
Attr_PathToRoot = "Attr_PathToRoot"

rootPath = "D:\\_Storage\\_Sync\\Documents\\root\\Notebook\\"
rootPath = node.map.file.parent
hfsPath = "192.168.150.100\\_Sync\\Documents\\root\\Notebook\\"

rootMode = "remotex"
rootxmindfilePath = os.path.join(rootPath, "root.xmind")
rootxmindfilePath = os.path.join(rootPath, "root.xmind")
rootxmindfolderPath = os.path.join(rootPath, "root\\")

staticFullList = []


class Node:
    def __init__(self, num):
        self.node = num


class NodeProxyEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Node):
            return {
                "node": obj.node.text,
                "id": obj.node.getId(),
                "children": [Node(i) for i in obj.node.children],
                'nodeIn': list(set([i.getSource().getId() for i in node.getConnectorsIn()]+[i.getId() for i in node.children])),
                'nodeOut': [i.getSource().getId() for i in node.getConnectorsOut()]
            }
        return super(NodeProxyEncoder, self).default(obj)


def deleteTopic(rootNode, Title):
    if rootNode.text == Title:
        rootNode.delete()
    else:
        if rootNode.children is not None:
            for child in rootNode.children:
                deleteTopic(child, Title)


def removeTopic(rootNode, Title):
    if rootNode.children is not None:
        for child in rootNode.children:
            if child.text == Title:
                if child.children is not None:
                    for childschild in child.children: childschild.moveTo(rootNode)
                    child.delete()

        for child in rootNode.children:
            removeTopic(child, Title)
    else:
        if rootNode.text == Title:
            rootNode.delete()


def deepMerge(fromParent, toParent):
    if fromParent.children is not None:
        for fromChild in fromParent.children:
            if toParent.children is not None:
                if fromChild.text in [i.text for i in toParent.children]:
                    targetNode = [i for i in toParent.children if
                                  i.text == fromChild.text][0]
                    deepMerge(fromChild, targetNode)
                else:
                    fromChild.moveTo(toParent)
            else:
                fromChild.moveTo(toParent)


def getNodeByID(c, node, id):
    return node.map.node(id)


def dfswalkThrough(nodeArray, tlist, writenotes, hiretachy="", force=False):
    if force is False and len(staticFullList) > 0 and nodeArray[0].getId() == nodeArray[
        0].getMap().getRootNode().getId():
        return staticFullList
    else:
        for node in nodeArray:
            for child in node.findAllDepthFirst():
                tlist += [child]
                flink = None
                if child.link.text is not None:
                    flink = child.link.text
                if child.text != SEPERATOR and writenotes == True:
                    # note = (child.getId()+"$$"+SEPERATOR2+SEPERATOR2.join([i.text for i in child.getPathToRoot()]))
                    child.attributes.set(Attr_PathToRoot, SEPERATOR2.join([i.text for i in child.getPathToRoot()]))
                    # child.noteText= note
                    child.setDetailsText("")
                    child.setHideDetails(True)
                    # if flink is not None:
                    #     ext = [".PNG",".png",".JPG",".jpg",".jpeg"]
                    #     if flink.endswith(tuple(ext)):
                    #         child.setDetailsText("<html><body><p><img src="+flink+"></p></body></html>")
            if node.getId() == node.getMap().getRootNode().getId():
                global staticFullList
                staticFullList = list(tlist)
    return tlist


def buildHTML(node):
    for child in node.findAllDepthFirst():
        flink = None
        if child.link.text is not None:
            flink = child.link.text
        if child.text != SEPERATOR:
            child.setDetailsText("")
            child.setHideDetails(True)
            if flink is not None:
                ext = [".PNG", ".png", ".JPG", ".jpg", ".jpeg"]
                if flink.endswith(tuple(ext)):
                    img = ImageIO.read(URL("file:" + flink))
                    mywidth = 400
                    wpercent = (mywidth / float(img.getWidth()))
                    hsize = int((float(img.getHeight()) * float(wpercent)))
                    child.setDetailsText(
                        "<html><body><p><img src=" + flink + " width=" + str(mywidth) + " height=" + str(
                            hsize) + "></p></body></html>")
                    child.parent.folded = True


def refreshSeperator(node):
    cleanSeperator(node)
    # addSeperator(node)


def cleanSeperator(node):
    if node.text == SEPERATOR:
        node.delete()
    else:
        if node.children is not None:
            for child in node.children:
                cleanSeperator(child)


def addSeperator(node):
    if node.children is not None:
        for child in node.children:
            addSeperator(node)
            child.createChild(SEPERATOR)


def getStructOfNode(nodes):
    return [i.text for i in nodes.getPathToRoot()]


def getCommonRoot(nodeN1, nodeN2):
    try:
        node1 = [i.text for i in nodeN1.getPathToRoot()]  # nodeN1.noteText.split("$$")[1].split(SEPERATOR2)
        node2 = [i.text for i in nodeN2.getPathToRoot()]  # nodeN2.noteText.split("$$")[1].split(SEPERATOR2)
        countstart = 0
        for i in range(0, min(len(node1), len(node2))):
            if node1[i] == node2[i]:
                countstart += 1
            else:
                break
        commonroot = node1[:countstart]
        node1 = node1[countstart:]
        node2 = node2[countstart:]
        return commonroot
    except:
        return []


def getDifferenceList(main, tocompare):
    count = 0
    for i in range(0, min(len(main), len(tocompare))):
        if main[i] == tocompare[i]:
            count += 1
        else:
            break
    return main[count:]


def getCommonRootList(node1, node2):
    try:
        countstart = 0
        for i in range(0, min(len(node1), len(node2))):
            if node1[i] == node2[i]:
                countstart += 1
            else:
                break
        commonroot = node1[:countstart]
        node1 = node1[countstart:]
        node2 = node2[countstart:]
        return commonroot
    except:
        return []


def is_hidden(filepath):
    name = os.path.basename(os.path.abspath(filepath))
    return name.startswith('_') or has_hidden_attribute(filepath)


def has_hidden_attribute(filepath):
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(unicode(filepath))
        assert attrs != -1
        result = bool(attrs & 2)
    except (AttributeError, AssertionError):
        result = False


def generateFolderNodes(node, pathtofolder):
    list_of_files = [os.path.join(pathtofolder, i) for i in os.listdir(pathtofolder)]
    for everyfile in list_of_files:
        try:
            if not ((os.path.basename(everyfile)).startswith('x_') or is_hidden(everyfile)):
                newchild = node.createChild(os.path.basename(everyfile))
                try:
                    newchild.link.text = alterPath(everyfile)
                except:
                    pass
                if os.path.isdir(everyfile) and (
                        os.path.basename(everyfile) not in [foldername(i) for i in everyfile.children]):
                    generateFolderNodes(newchild, everyfile)
        except:
            pass


def debugprint(strtest):
    try:
        node.getMap().getRootNode().createChild((strtest))
    except:
        pass


def getUniqueList(list):
    unique = {}
    for ele in (list):
        if ele.text not in unique:
            unique[ele.text] = [ele]
        else:
            if ele.getId() not in [i.getId() for i in unique[ele.text]]:
                unique[ele.text] = [ele] + unique[ele.text]

    return unique


def abbrev(string1, string2):
    hira1 = [i.text for i in string1.getPathToRoot()]  # .noteText.split("$$")[1].split(SEPERATOR2)
    hira2 = [i.text for i in string2.getPathToRoot()]  # .noteText.split("$$")[1].split(SEPERATOR2)
    difference = [i for i in hira2 if i not in hira1]
    returnstr = ""

    if len(difference) >= 3:
        returnstr = (difference[0]) + returnstr + "->"
        returnstr = returnstr + (difference[-2]) + "->"
        returnstr = returnstr + (difference[-1])
    elif len(difference) == 2:
        returnstr = (difference[0]) + "->" + (difference[-1])
    elif len(difference) == 1:
        c.statusInfo = "gggg"
        returnstr = (difference[-1])
    else:
        returnstr = ""

    return returnstr


def relateTopics2(node, fileDict, topicDict):
    # add search
    search = None
    if node.children is not None:
        for trash in node.children:
            if trash.text == SEPERATOR:
                search = trash

    if search is None:
        search = node.createChild(SEPERATOR)
    else:
        search.delete()
        search = node.createChild(SEPERATOR)

    # Add Search
    if node.text in topicDict:
        # search.createChild(str(len(topicDict[node.text])))
        for eachTopic in reversed(topicDict[node.text]):
            relateNode = node.map.node(eachTopic['node'])
            if not relateNode is None:
                if relateNode != node:
                    prefix = "(" + node.text + ") "
                    suffix = ""
                    v = search.createChild(abbrev(node, relateNode) + "->"+node.text)  # set its title
                    v.attributes.set(Attr_PathToRoot, SEPERATOR2.join([i.text for i in relateNode.getPathToRoot()]))
                    # v = search.createChild(relateNode.getId())
                    v.link.text = ("#" + relateNode.getId())

    # add url
    regex = re.compile(
        r'^(?:http|ftp)s?://'  # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|'  # domain...
        r'localhost|'  # localhost...
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'  # ...or ip
        r'(?::\d+)?'  # optional port
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)
    if re.match(regex, node.text) is not None:
        node.link.text = node.text

    # Add Folder & Files
    if foldername(node) in fileDict:
        pathtofolder = fileDict[foldername(node)]
        v = search.createChild("FILE")
        v.link.text = alterPath(pathtofolder)
        generateFolderNodes(v, pathtofolder)

    #Delete Uneccessary Search
    if len(search.children) ==0: search.delete()


def findFolderByID(id, fileList):
    result = []
    for paths in fileList:
        if id in os.path.basename(paths[0]):
            result = [paths[0]]
            return result
    return result


def linkPath(path):
    return "http://" + pathname2url((path.replace(rootPath, hfsPath)).encode("utf-8"))


def localPath(path):
    return pathname2url((path).encode("utf-8"))


def alterPath(path):
    if (rootMode == "remote"):
        return linkPath(path)
    else:
        return localPath(path)


def compareTree(node1, node2, same=True):
    if (node1.children is not None) and (node2.children is not None):
        node1Child = [i for i in node1.children]
        node2Child = [i for i in node2.children]
        if set([t.text for t in node1Child]) == set([k.text for k in node2Child]):
            same = same and True
            for node in node1Child:
                same = same and compareTree(node, [n for n in node2Child if n.text == node.text][0], same)
        else:
            same = False
    elif (node1.children is None) and (node2.children is None):
        same = same and True
    else:
        same = False

    return same


def autoClassificationByStruct(parent, struct):
    difference = getDifferenceList(struct[:-1], getStructOfNode(parent))
    difference = filter(lambda a: a != parent.text, difference)
    print(struct)
    print(getStructOfNode(parent))
    print(difference)
    currentNode = parent
    while len(difference) > 0:
        print(difference)
        if currentNode.children is not None:
            found = False
            for child in currentNode.children:
                if child.text in difference:
                    difference = filter(lambda a: a != child.text, difference)
                    currentNode = child
                    found = True
            if found == False: break
        else:
            break

    for residue in difference:
        currentNode = currentNode.createChild(residue)
    return currentNode


def autoClassificationByNode(fromNode, toParent):
    target = autoClassificationByStruct(toParent, getStructOfNode(fromNode))
    fromNode.moveTo(target)


def cleanDirt(listTopreserve, rootNode):
    endingNode = [i for i in dfswalkThrough([rootNode], [], False, force=True) if
                  (i.children is None or len(i.children) == 0)]
    print([i.text for i in endingNode])
    change = 1
    while change > 0:
        print([i.text for i in endingNode])
        change = 0
        for nodes in endingNode:
            if nodes.text not in listTopreserve:
                print("DELETED: " + nodes.text)
                nodes.delete()
                change += 1
        endingNode = [i for i in dfswalkThrough([rootNode], [], False) if
                      (i.children is None or len(i.children) == 0)]


def divideTreebyNode(c, node, rootID, dividerID):
    fullList = dfswalkThrough([node.getMap().getRootNode()], [], True, hiretachy="")
    # update
    rootNode = getNodeByID(c, node, rootID)
    dividerNode = getNodeByID(c, node, dividerID)
    endingNodesTitleList = [i.text for i in dfswalkThrough([rootNode], [], False) if
                            (i.children is None or len(i.children) == 0)]
    print("gggggggggggggg")
    print(endingNodesTitleList)

    # get samelayer\
    samelayer = [dividerNode.text]
    samelayerNode = [dividerNode]
    partialList = dfswalkThrough([rootNode], [], False)
    for nodes in partialList:
        if nodes.children is not None:
            if dividerNode.text in [k.text for k in nodes.children]:
                samelayer = list(set(samelayer + [k.text for k in nodes.children]))
                samelayerNode = list(set(samelayerNode + [k for k in nodes.children]))

    # draw boundry
    boundry = []
    for node in samelayerNode:
        print
        node.text
        boundry += [i.text for i in dfswalkThrough([node], [], False)]
        boundry += getStructOfNode(node)

    # others
    # move branch to tmp
    tmpBranch = node.getMap().getRootNode().createChild("tmpBranch")
    for child in rootNode.children: child.moveTo(tmpBranch)

    # add samlayer
    for layer in samelayer:
        rootNode.createChild(layer)

    # Others
    rootNode.createChild("Others")

    # copy back
    originalcopied = False
    for childNode in rootNode.children:
        if originalcopied == False:
            for child in tmpBranch.children: child.moveTo(childNode)
            originalcopied = True
        else:
            # for child in rootNode.children[0].children:
            copyTree(rootNode.children[0], childNode)

    dfswalkThrough([fullList[0]], fullList, True, "")

    ######Classified
    # delete nonmember
    for child in rootNode.children:
        todelete = [s for s in samelayer if s != child.text]
        print(todelete)
        for stringtodelete in todelete: deleteTopic(child, stringtodelete)

    # remove old class
    for child in rootNode.children:
        removeTopic(child, child.text)

    notothers = []
    for j in rootNode.children:
        if j.text != "Others":
            notothers += [j]

    # delete bad
    dfswalkThrough([fullList[0]], [], True, "")
    for sameNode in notothers:
        for newnodes in [i for i in dfswalkThrough([sameNode], [], True, "")]:
            if not set(getStructOfNode(newnodes)).issubset(boundry):
                newnodes.delete()

    # clean class residue
    print(endingNodesTitleList)
    cleanDirt(endingNodesTitleList, rootNode)

    # clean others
    for child in rootNode.children:
        removeTopic(child, "Others")


def copyTree(root, parent):
    if root.children is not None:
        for child in root.children:
            newchile = parent.createChild(child.text)
            copyTree(child, newchile)
    return parent
    # parent.appendAsCloneWithSubtree(root)


def foldername(node):
    return node.text + "_" + node.getId()


def generateFolder(node):
    fileList = [i for i in os.walk(rootPath) if os.path.isdir(i[0])]
    foldername = node.getId()
    # pathtofolder = [x[0] for x in os.walk(rootPath) if foldername in os.path.basename(x[0])]
    pathtofolder = findFolderByID(foldername, fileList)
    if not len(pathtofolder) > 0:
        os.makedirs(os.path.join(findNearestParentFolder(node,fileList), (node.text + "_" + node.getId())))


def findNearestParentFolder(currnode,allFolder):
    # allFolder = [i for i in os.walk(rootPath)]
    try:
        parentFolder = [i for i in allFolder if currnode.parent.getId() in os.path.basename(i[0])]
        if len(parentFolder) > 0:
            return parentFolder[0][0]
        else:
            return findNearestParentFolder(currnode.parent,allFolder)
    except:
        return rootPath


def toggleAllChild(node, toggle):
    node.folded = toggle
    if node.children is not None:
        for child in node.children:
            toggleAllChild(child, toggle)
    else:
        pass


def toggleSeperatorNode(node, toggle):
    if node.text == SEPERATOR:
        node.folded = toggle
    if node.children is not None:
        for child in node.children:
            toggleSeperatorNode(child, toggle)
    else:
        pass


def copyTreeOneLayer(root, parent):
    if root.children is not None:
        for child in root.children:
            newchile = parent.createChild(child.text)
    return parent
    # parent.appendAsCloneWithSubtree(root)


def collapseOneLayer(root):
    if root.children is not None:
        for child in root.children:
            if child.children is not None:
                for childschild in child.children:
                    childschild.moveTo(root)
            child.delete()


def deleteChild(node):
    if node.children is not None:
        for child in node.children:
            child.delete()
    else:
        pass


# dfswalkThrough([node], [], True)

# completeList = dfswalkThrough([node.getMap().getRootNode()], [], True, hiretachy="" )

# deleteChild(node)

# collapseOneLayer(node)

# copyTreeOneLayer(getNodeByID(c,node,"ID_665552210"),node)

# divideTreebyNode(c,node,"ID_1980548619", "ID_1445552548")

# autoClassificationByNode(getNodeByID(c,node,"ID_1490770708"),node)

# copyTree(completeList,getNodeByID(c,node,"ID_525529922"),node)

# generateFolder(node)

# relateTopicsDriver(node)

# removeTopic(node,"專案管理結構實現專案")


# toggleSeperatorNode(node.getMap().getRootNode(),True)


# node.createChild(node.getId()+"$$___"+"___".join([i.text for i in node.getPathToRoot()]))


def generateFolderDriver(node):
    generateFolder(node)


def partialRefresh2(node):
    cleanAll(node.parent)
    # completeList = dfswalkThrough([node.parent], [], True, hiretachy="" )
    relateTopicsDriver(node.parent, False)
    buildHTML(node.parent)
    toggleSeperatorNode(node.parent, True)
    setAttributes(node)


def partialRefresh(node):
    cleanAll(node)
    # completeList = dfswalkThrough([node], [], True, hiretachy="" )
    relateTopicsDriver(node, False)
    buildHTML(node)
    toggleSeperatorNode(node, True)
    setAttributes(node)


def cleanAll(node):
    deleteTopic(node, SEPERATOR)
    for child in node.findAllDepthFirst():
        child.attributes.clear()
        child.setHideDetails(True)
        child.link.text = ""
        child.link.text = None
        child.noteText = None
        child.setDetailsText(None)


def cleanNotes(node):
    for child in node.findAllDepthFirst():
        child.setHideDetails(True)
        # child.link.text = ""
        # child.link.text = None
        child.noteText = None
        # child.setDetailsText(None)


def setAttributes(node):
    for child in node.findAllDepthFirst():
        parents = getStructOfNode(child)
        if SEPERATOR in parents:
            if "FILE" not in parents:
                child.attributes.set(Attr_relate, 1)
        #     else:
        #         child.attributes.set(Attr_relate, 0)
        # else:
        #     child.attributes.set(Attr_relate, 0)
        if "Inactive" in parents:
            child.attributes.set(Attr_InActive, 1)
        
            # child.attributes.set(Attr_Active, 0)
        # else:
        #     child.attributes.set(Attr_Active, 1)

def globalLastModified(node):
    return max([i.lastModifiedAt for i in node.findAllDepthFirst()])

def genTopicDict(node,topicDict={}):
    for i in node.getMap().getRootNode().findAllDepthFirst():
        topic = {'node':i.getId(),'globalLastModified':globalLastModified(i)}
        if not i.text in topicDict:
            topicDict[i.text] = [topic] #[i.getId()]
        else:
            insertionIndex = bisect.bisect_left([j['globalLastModified'] for j in topicDict[i.text] ], topic['globalLastModified'])
            topicDict[i.text].insert(insertionIndex,topic)
    return topicDict

def relateTopicsDriver(node, dosearch=True):
    if dosearch == True:
        topicDict=genTopicDict(node.getMap().getRootNode())
    else:
        with open(os.path.join(rootPath, 'RelateTopics.json')) as json_data:
            topicDict =genTopicDict(node,json.load(json_data))
    fileDict = {}
    for root, dirs, files in os.walk(rootPath):
        fileDict[os.path.basename(root)] = root
    for i in node.findAllDepthFirst():
        relateTopics2(i, fileDict, topicDict)
    


def saveMaps(node):
    ValueNode = node.map.node('ID_1992023727')
    sortChildrenByGlobalLastModified(ValueNode)
    with open(os.path.join(rootPath, 'ValueMap.json'), 'w') as outfile:
        json.dump(Node(ValueNode), outfile, cls=NodeProxyEncoder, indent=2)
    with open(os.path.join(rootPath, 'FullMap.json'), 'w') as outfile:
        json.dump(Node(node.getMap().getRootNode()), outfile, cls=NodeProxyEncoder, indent=2)
    with open(os.path.join(rootPath, 'RelateTopics.json'), 'w') as outfile:
        json.dump(genTopicDict(node), outfile, indent=2, default=str)

def fullRefresh(node):
    cleanAll(node.getMap().getRootNode())
    relateTopicsDriver(node.getMap().getRootNode(), dosearch=True)
    buildHTML(node)
    toggleSeperatorNode(node.getMap().getRootNode(), True)
    setAttributes(node.getMap().getRootNode())


def NoSeatchRefresh(node):
    cleanAll(node.getMap().getRootNode())
    relateTopicsDriver(node.getMap().getRootNode(), dosearch=False)
    buildHTML(node)
    toggleSeperatorNode(node.getMap().getRootNode(), True)
    setAttributes(node.getMap().getRootNode())


def nodeRefresh(node):
    cleanAll(node)
    relateTopicsDriver(node, False)
    buildHTML(node)
    toggleSeperatorNode(node, True)
    setAttributes(node)

def sortChildrenByGlobalLastModified(node):
    childList = [(i,globalLastModified(i) ) for i in node.children]
    childList.sort(key=lambda tup: tup[1])  # sorts in place
    for idx, val in enumerate(reversed(childList)):
        val[0].moveTo(node, idx)
        sortChildrenByGlobalLastModified(val[0])



# cleanAll(node.getMap().getRootNode())

####Refresh Map
# fullRefresh(node)
# NoSeatchRefresh(node)
# nodeRefresh(node)


####Save Maps
saveMaps(node)


####Generate Folders
# generateFolder(node)    