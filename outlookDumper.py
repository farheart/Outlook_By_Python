# encoding: utf-8

"""
1. Select a OS folder to save the backup .pst
2. Select a Outlook folder to start backup
3. Go thru all sub-folders under the start folder and do msg backup
"""


import os
import win32com.client
import wx


def listFolders(root, pathStr):
    pathStr = os.path.join(pathStr, root.Name)
    print(pathStr)
    os.makedirs(pathStr)
    
    for m in root.Items:
        fn = getFileName(pathStr, m)
        print('>> ' + fn)
        m.SaveAs(fn)
    
    for f in root.Folders:
        listFolders(f, pathStr)
       
       
def safeName(name):
    return "".join([' ' if c in '\/:*?"<>|' else c for c in name])


def getFileName(pathStr, msg):
    baseName = safeName(msg.Subject)
    fn = os.path.join(pathStr, baseName) + '.msg'
    if os.path.exists(fn):
        i = 1
        while os.path.exists(fn):
            newName = '{0} - ({1})'.format(baseName, i)
            fn = os.path.join(pathStr, newName) + '.msg'
            i += 1
    return fn


def main():
    # Select a target OS folder
    dumpPath = get_path()
    
    # Select a Outlook folder to start.
    win32com.client.gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 4)
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    source_folder = outlook.PickFolder()
    if source_folder:
        listFolders(source_folder, dumpPath)    
        
        
def get_path():
    app = wx.App(None)
    dlg = wx.DirDialog(None, "Choose a folder to dump:")
    if dlg.ShowModal() == wx.ID_OK:
        path = dlg.GetPath()
    dlg.Destroy()
    return path    


if __name__ == '__main__':
    main()
