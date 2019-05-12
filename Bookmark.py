shortcut = [i for i in node.getMap().getRootNode().children if i.text =='Shortcut'][0].createChild(node.text)
shortcut.link.text ="#"+node.getId()