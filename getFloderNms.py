import os
import xlsxwriter
from office import Excel

filePath1 = 'C:\Project\PY_TOOLS\TestGroup'



def test():
    clIndex = 0
    homepath = os.getcwd()
    needHandleList = os.listdir(filePath1)
    print(type(needHandleList))
    workbook = xlsxwriter.Workbook('TemplateXlsxWriter.xlsx')
    worksheet = workbook.add_worksheet('template')
    for folderNm in needHandleList:
        clIndex += 1
        if folderNm.split('_')[-1] == 'RLFS':
            worksheet.write(clIndex, 4, folderNm)
        elif folderNm.split('_')[-1] == 'RLS' or folderNm.split('_')[-1] == 'LFS' or folderNm.split('_')[-1] == 'LS':
            clIndex -= 1
        else:
            worksheet.write(clIndex, 4, folderNm)
        print(clIndex)
    workbook.close()
    Excel().open(homepath+r'\TemplateXlsxWriter.xlsx')
    
if __name__=='__main__':
    test()    