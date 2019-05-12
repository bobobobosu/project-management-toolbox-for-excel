def globalLastModified(node):
    return max([i.lastModifiedAt for i in node.findAllDepthFirst()])
    
def sortChildrenByGlobalLastModified(node):
    childList = [(i,globalLastModified(i) ) for i in node.children]
    childList.sort(key=lambda tup: tup[1])  # sorts in place
    for idx, val in enumerate(reversed(childList)):
        val[0].moveTo(node, idx)
        sortChildrenByGlobalLastModified(val[0])

def all2str(node):
        stral = ""
        for allnode in node.getMap().getRootNode().findAllDepthFirst():
                stral+=str(allnode.getConnectorsIn())
        return stral

def getConnectingNodes(node):
        InOut = {}
        InOut['nodeIn'] = list(set([i.getSource().getId() for i in node.getConnectorsIn()]+[i.getId() for i in node.children]))
        InOut['nodeOut'] = [i.getSource().getId() for i in node.getConnectorsOut()]
        return InOut

def setConnectorShape(node):
        for allnodes in node.findAllDepthFirst():
                for i in node.getConnectorsIn():
                        i.setShape('LINE')
                for j in node.getConnectorsOut():
                        j.setShape('LINE')
# sortChildrenByGlobalLastModified(node)
setConnectorShape(node.getMap().getRootNode())
node.attributes.set('in',str(getConnectingNodes(node)['nodeIn']))
node.attributes.set('out',str(getConnectingNodes(node)['nodeOut']))