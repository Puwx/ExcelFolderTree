"""
This script creates an Excel table workbook that contains the folder structure
for the folder that is select through the tkinter askdirectory file dialog
when the script is run.
"""
import os
import re
import xlwt
import tkinter
from tkinter.filedialog import askdirectory


def generateFileTree(base_dir):
    os.chdir(base_dir)
    
    baseDepth = os.getcwd().count(os.sep)
    ignore = ['.git','hooks','info','logs','objects','.vscode','__pycache__',
              'refs','lfs','heads','remotes','tmp','origin','pack','tags']
    
    foldList = []
    rowCount = 0
    for full,__,files in os.walk(os.getcwd()):
        if any([os.path.basename(full).endswith(ignoreExt) for ignoreExt in  ignore]):
            continue
        #Certain invisible folder will be named 2 characters and just contain non-extensioned, hexadecimal files.
        #The code below identifies these folders and skips them since they aren't really part of the folder tree.
        elif len(os.path.basename(full)) == 2:
            if all([re.fullmatch(r"[a-f0-9]{38,}",file) for file in files]):
                continue
        foldLevel = full.count(os.sep)-baseDepth
        foldList.append([full,foldLevel,rowCount])
        rowCount +=1
        for file in files:
            foldList.append([os.path.join(full,file),foldLevel+1,rowCount])
            rowCount +=1
        rowCount+=1
        
    return foldList       

def createFileTree(folder_map):
    fullBaseDir = folder_map[0][0]
    wb = xlwt.Workbook()
    ws = wb.add_sheet('FILE TREE')
    foldStyle = xlwt.Style.easyxf('font: bold on, color red')
    fileStyle = xlwt.Style.easyxf('font: bold on, color blue')
    
    for fileObj,colDepth,row in folder_map:
        style = foldStyle if os.path.isdir(fileObj) else fileStyle
        ws.write(row,colDepth,os.path.basename(fileObj),style)
    wb.save(f'{os.path.basename(fullBaseDir)}_Structure.xls')



if __name__ == "__main__":
    root = tkinter.Tk()
    foldOI = askdirectory(parent=root)
    root.destroy()
    
    output = generateFileTree(foldOI)
    createFileTree(output)
