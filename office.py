# -*- coding: utf-8 -*-
'''
Created on Jul 12, 2019

@author: wwang14
'''
from win32com.client import Dispatch
import os
from time import sleep

class Word():
    def __init__(self):
        self.word = 0
        self.word_connect()
        self.word_file = None
        
    def word_connect(self):
        self.word = Dispatch("Word.Application")
        
    def open(self, path_cfg_file, isVisible=True):
        if os.path.isfile(path_cfg_file):
            self.word.Visible = isVisible
            try:
                self.word_file = self.word.Documents.Open(path_cfg_file,Encoding='gbk') 
                #self.word_file = self.word.Documents.Add()
                return True
            except:
                print("Word already open, keep the existing process")
                return False
            sleep(5)
    
    def insertBookmark(self,name):# insert Bookmark with name
#         print self.word_file.Bookmarks.Count
        self.word_file.Bookmarks.Add(name)        
        
    def insertTable(self,bookmark_name,row,col):#insert Table after bookmark
        bk = self.word_file.Bookmarks(bookmark_name)
        rg = bk.Range
        self.word_file.Tables.Add(rg, row, col, 1, 0)
            
    def getTableCount(self):#get table count
        return self.word_file.Tables.Count
    
    def getRowCountOfTable(self,tb_id):#get row count of a table
        return self.word_file.Tables[tb_id].Rows.Count
    
    def getColCountOfTable(self,tb_id):#get column count of table
        return self.word_file.Tables[tb_id].Columns.Count
    
    def deleteRowOfTable(self,tb_id,row):#delete a row of a table
        self.word_file.Tables[tb_id].Rows[row].Delete()
    
    def deleteColOfTable(self,tb_id,col):#delete a column of a table
        self.word_file.Tables[tb_id].Columns[col].Delete()
        return True
    
    def insertRowOfTable(self,tb_id,r_id):#insert a row of a table
        row = self.word_file.Tables[tb_id].Rows(r_id)
        self.word_file.Tables[tb_id].Rows.Add(row)
        return True
    
    def insertColOfTable(self,tb_id,c_id):#insert a column of a table
        row = self.word_file.Tables[tb_id].Columns(c_id)
        self.word_file.Tables[tb_id].Columns.Add(row)
        return True
    
    def setOutlineStyleofTable(self,tb_id, Outlinestyle):#set outline style of a table
        #设置table的外边框  2为虚线  1实线  0是没有线
        self.word_file.Tables[tb_id].Borders.OutsideLineStyle = Outlinestyle
        return True
    
    def setRowHeightofTable(self,tb_id, row, height):#set row height of a table
        self.word_file.Tables[tb_id].Rows(row).Height = height
        return True
    
    def setColWidthofTable(self,tb_id, col, width):
        self.word_file.Tables[tb_id].Columns(col).Width = width
        return True
    
    def selectTable(self,tb_id):#select on a table
        self.word_file.Tables[tb_id].Select()
        return True
    
    def getCellOfTable(self,tb_id,row,col): # get value of a cell in a table
        #self.word_file.Tables[tb_id].Cell(row,col).Select()
        #print self.word_file.Tables[tb_id].Cell.Row
        return self.word_file.Tables[tb_id].Cell(row,col).Range.Text
    
    def setCellOfTable(self,tb_id,row,col,value):#set value of a cell in a table
        #self.word_file.Tables[tb_id].Cell(row,col).Select()
        self.word_file.Tables[tb_id].Cell(row,col).Range.Text = value
        return True   
    
    def setFontSizeofCell(self,tb_id, row, col, fontsize):#set front size of a cell
        self.word_file.Tables[tb_id].Cell(row,col).Range.Font.Size = fontsize
        return True
    
    def setFontStyleofCell(self,tb_id, row, col, fontname):#set front size of a cell
        self.word_file.Tables[tb_id].Cell(row,col).Range.Font.Name = fontname
        return True
    
    def setUnderlineofCell(self,tb_id, row, col, flag=True):#set underline of a cell
        self.word_file.Tables[tb_id].Cell(row,col).Range.Font.Underline = flag
        return True      
                
    def save(self, newfilename=None):  #save file
        if newfilename:  
            self.filename = newfilename  
            self.word_file.SaveAs(newfilename)  
        else:  
            self.word_file.Save()   
        return True
            
    def close(self,SaveChanges=1):  #close and save file
        self.word_file.Close(SaveChanges)  
        return True
    
    def quit(self): # quit Word Application
        self.word.Quit()
        del self.word
        return True
    
class Excel():
    def __init__(self):
        self.excel = 0
        self.excel_connect()
        self.excel.Visible = True
        self.excel_file = None
        #VTSUsedRowCnt = VTSSheet.UsedRange.Rows.Count
    
    def excel_connect(self): # connect excel application
        self.excel = Dispatch("Excel.Application")
    
    def createNew(self,name):
        self.excel_file = self.excel.Workbooks.Add() 
        self.excel_file.SaveAs(name)
        return True
    
    def open(self, path_cfg_file, isVisible=True):#open file by path
        if os.path.isfile(path_cfg_file):
            self.excel.Visible = isVisible
            try:
                self.excel_file = self.excel.Workbooks.Open(path_cfg_file) 
            except:
                print("Excel already open, keep the existing process")
            sleep(5)
    
    def getSheetNames(self):#Get a list of all sheet name
        shtName = list()
        for sheet in self.excel_file.Worksheets:
            shtName.append(sheet.Name)
        return shtName
    
    def deleteSheet(self,name):#delete a sheet by name
        self.excel_file.Worksheets(name).Delete()
        return True
    
    def activeSheet(self,name):
        self.excel_file.Worksheets(name).Active()
        return True
    
    def selectSheet(self,name):#select postion
        self.excel_file.Worksheets(name).Select()
        return True
            
    def addSheet(self,name):#add a sheet by name
        if name in self.getSheetNames():
            print("sheet %s alread exists" % name)
            return True
        sht = self.excel_file.Worksheets.Add()
        sht.Name = name
        return True
    
#     def copySheet(self,name):
#         self.excel_file.Worksheets(name).Copy()
    def selectCell(self,sheet,row,col):
        sht = self.excel_file.Worksheets(sheet)
        sht.Cells(row,col).Select()
        return True
    
    def getCell(self, sheet, row, col):#get value of a cell by sheet name,row and colume 
        sht = self.excel_file.Worksheets(sheet)  
        return sht.Cells(row, col).Value
    
    def setCell(self, sheet, row, col, value):  #set value of a cell by sheet name,row and colume
        try:
            sht = self.excel_file.Worksheets(sheet)  
            sht.Cells(row, col).Value = value
        except:
            self.setCell(sheet, row, col, value)
            sleep(1)
        return True  
    
    def getRange(self, sheet, row1, col1, row2, col2):  #get values of range by sheet name,row and colume;return a 2d tupples
        #"return a 2d array (i.e. tuple of tuples)"  
        sht = self.excel_file.Worksheets(sheet)  
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value  
    
    def save(self, newfilename=None):  #save file
        if newfilename:
            self.filename = newfilename  
            self.excel_file.SaveAs(newfilename)  
        else:  
            self.excel_file.Save()   
        return True
            
    def close(self,SaveChanges=1):  #close and save file
        self.excel_file.Close(SaveChanges)  
#         self.excel.Quit()
#         del self.excel
        return True
    
    def quit(self):
        self.excel.Quit()
        del self.excel
        return True
    
if __name__=='__main__':
    excel = Excel()
    file_path = os.getcwd()
    excel.createNew(file_path+r'\aaabbb.xlsx')
    excel.addSheet("aaabbb")
    excel.save()
    excel.quit()