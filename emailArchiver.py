# encoding: utf-8

"""
Automatically archive emails (weekly report) into 
corresponding Outlook folders by their senders
"""

import os
import wx
import win32com.client

import time
import pandas as pd


# Retrieve Outlook.Inbox object
def initOutlookInbox():
    win32com.client.gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 4)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    return outlook.GetDefaultFolder(6)  # InBox


# Parse pathStr to real location (folder) in Outlook
def parseTarget(rootFolder, pathStr):
    fd = rootFolder
    try:
        for f in pathStr.split('/'):
            fd = fd.Folders[f]
        # print(fd.Name)
    except:
        print("ERROR: path not exist!")
        fd = None
    return fd


# Whether or not the email should be moved by keywords in its title
# e.g., ('42' || '四十二') && '周报' 
def isToMove(msg, weekWords):
    ifToMove = False
    for w in weekWords:
        ifToMove = ifToMove or (w in msg.Subject)
    return ('周报' in msg.Subject) and ifToMove


# load name-folder mapping from mapList.csv 
def loadNameMap():
    df = pd.read_csv("mapList.csv")
    nameMap = df.set_index('name').T.to_dict('list')
    # print(nameMap)
    return nameMap


# Move weekly report email to its achive location 
def sortEmails(weekWords):
    nameMap = loadNameMap()
    inbox = initOutlookInbox()  # InBox
    if inbox:
        for msg in inbox.Items:
            if isToMove(msg, weekWords):
                name = msg.SenderName if msg.SenderName in nameMap else 'Other'
                pathStr = nameMap[name][0].format(weekWords[0])  # get location by sender's name and week num

                print('>> {}'.format('='*20))
                print('>> Subject = \t' +  msg.Subject)
                print('>> SenderName = ' +  msg.SenderName)
                print('>> pathStr = \t' +  pathStr)

                target = parseTarget(inbox, pathStr)
                if target:
                    print('>> target = \t' +  target.Name)
                    msg.Move(target)
                    print('>> Moved')
                    time.sleep(1.0)  # allow outlook to finish the move


# Create name - location map from given folder (defined in pathStr)
def createMapList(pathStr = '2019/周报/WK42'):
    def goThru(curFolder, pathStr, mapList):
        for m in curFolder.Items:
            mapList.append({'name': m.SenderName, 'folder': pathStr})
            # print('>> {} \t@\t {}'.format(m.SenderName, pathStr))
        for f in curFolder.Folders:
            goThru(f, pathStr + '/' + f.Name, mapList)

    inbox = initOutlookInbox()  # InBox
    fd = parseTarget(inbox, pathStr)
    if fd:
        mapList = []
        goThru(fd, pathStr, mapList)

        print(">> Showing map ...")
        for m in mapList:
            print('>> Email from {} \t>>\t {}'.format(m['name'], m['folder']))
        
        df = pd.DataFrame(mapList)
        print(df)
        # df.to_csv("mapList.csv", index=False)


if __name__ == '__main__':
    # createMapList('2019/周报/WK42')
    sortEmails(['42', '四十二'])
