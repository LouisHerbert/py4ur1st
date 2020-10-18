# -*- coding: UTF-8 -*-
'''
Created on 2020年10月15日

@author: swei
'''
from word2tsxml import txt
# import word2tsxml
import os
import shutil
# import numpy as np
# from _ast import Return
# from django.views.decorators.http import condition
from numpy import double
communiframe={
    "Time":0.000,
    "Chn":"Li",
    "ID":0x00,
    "FrameName":"",
    "Dir":"Rx",
    "DLC":8,
    "Data":[0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00],
    "checksum":0x00,
    "headertime":0.000,
    "fulltime":0.000,
    "SOF":0.000,
    "BR":0x00,
    "break":0x00,
    "EOH":0.00,
    "EOB":[0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00],
    "EOF":0.00,
    "RBR":19200,
    "HBR":19200.000,
    "HSO":0,
    "ESO":0,
    "CSM":"enhanced"
    }
SVFrame={
    "Time":0.000,
    "Id":"SV:",
    "FrameName":"", #::IO::VN1600_1::DIN0 ::IO::VN1600_1::AIN
    "Data":""
    }
RevErrorFrame={
    "Time":0.000,
    "Chn":"Li",
    "EventType":"RcvError: unexpected byte during bus idle phase ",
    "char":0x00,  # or""
    "StateReason":0,
    "ShortError":0,
    "DlcTimeout":0,
    "HasDatabytes":0,
    "SOF":0.000,
    "BR":0x00,
    "break":0x00,
    "RBR":19200,
    "RSO":0,
    "HBR":0.000,
    "HSO":0,
    "CSM":"unknown"    
    }
SpecialFrame={  #WakeupFrame
    "Time":0.000,
    "Chn":"Li",
    "EventType":"SleepModeEvent",  #SleepModeEvent/Dominant signal detected/Dominant signal continuing/Dominant signal finished
    "Dir":"",  #Rx
    "Data":"",  #entering sleep mode due to bus idle timeout
    "Time":"",  #7000602 microseconds
    "SOF":0.000,
    "BR":0.000,
    "lengthCode":0
    }

class ASCFile(object):
    def __init__(self):
        self.filepath=None
        self.filename=None
        self.txt=txt()
        self.comframe=dict()
        self.userframe=dict()
        pass
    def Asc2Txt(self,filepath,filename):
        self.filepath=filepath
        self.filename=filename[:-4]+".txt"
        files = os.listdir(filepath) 
        if(filename in files):
            ascpath=filepath+"\\"+filename
            new_txt_name=filename[:-4]+".txt"
            shutil.copy(ascpath, self.filepath+"\\"+new_txt_name)
        else:
            print(filename+" is not exist in the path of "+self.filepath)
    def ReplaceSpace(self,str_frame):
        List=str_frame.split(" ")
        List_new=list()
        for item in List:
            if(item!=""):
                List_new.append(item)
        str_new=""
        for item in List_new:
            if(item!=List_new[-1]):
                item=item+" "
            str_new=str_new+item
        return str_new
    def GetFrameInfo(self):
        self.txt.open(self.filepath, self.filename,"r")
        #print(self.filename)
        lineList=(self.txt.read()).split("\n")
        RxList=[]
        TxList=[]
        RevErrorList=[]
        SVList=[]
        SpecList=[]
        infoList=[]
        startrecord=False
        for line in lineList:
            line=self.ReplaceSpace(line)
            if(line.startswith("Begin TriggerBlock")>0):
                startrecord=True
            elif(line==lineList[-1]):
                startrecord=False
            if(startrecord==True):
                if(("Rx" in line) and ("EOB" in line) and ('Li' in line)):
                    RxList.append(line)
                elif("Tx" in line):
                    TxList.append(line)
                elif("RcvError" in line):
                    RevErrorList.append(line)
                elif("SV" in line):
                    SVList.append(line)
                else:
                    SpecList.append(line)
                    #print(line) 
        infoList.append(RxList)
        infoList.append(TxList)
        infoList.append(RevErrorList)
        infoList.append(SVList)
        #print((RxList))
        #print((TxList))
        #print((RevErrorList))
        #print((SVList))
        #print((SpecList))
        self.comframe=communiframe.copy()
              
        return infoList
    def GetCommFrameInfo(self,infoList,flag="Rx"):
        comfra=dict()
        for index in range(0,len(infoList)):
            comfra[str(index)]=self.CombCommunFrameInfo(infoList[index])
#         print(comfra)
        self.comframe=comfra
    def CheckValue(self,List,targetList,startindex=0):
        for index in range(0,len(targetList)):
            if(List[startindex+index] == targetList[index]):
                pass
            else:
                print(List[index]+" is not equal to "+targetList[index])
                return False
        return True
    def FindPosition(self,List,targetList,startindexList):
        if(len(targetList)==len(startindexList)):
            for target in targetList:
                for index in range(0,len(target)):
                    if(self.CheckValue(List, target, startindexList[index])):
                        pass
                    else:
                        print(str(target)+" Can not find in List["+str(startindexList[index])+"]")
                        return False
        else:
            print("length of targetList is not equal to startindexList")
            return False
        return True
    def GetValue(self,List,startindex,length,flag="hex"):
        list_re=[]
        for i in range(0,length):
            if(flag == "hex"):
                list_re.append(hex(bytes.fromhex(List[startindex+i].zfill(2))[0]))
            elif(flag == "int"):
                list_re.append(int(List[startindex+i]))
            elif(flag == "double"):
                list_re.append(double(List[startindex+i]))
            elif(flag == "string"):
                list_re.append(List[startindex+i])
            else:
                print("Flag error,only can select hex/int/double/string")
                return False
        if(len(list_re)==1):
            list_re=list_re[0]
        return list_re
    def CombCommunFrameInfo(self,str_info):
        List=str_info.split(" ")
        comframe=communiframe.copy()
        #print(len(List))
        comframe["Time"]=(List[0])
        comframe["Chn"]=List[1]
        comframe["ID"]=hex(bytes.fromhex(List[2].zfill(2))[0])
        if((List[3]=="Rx")or (List[3]=="Tx")):
            comframe["FrameName"]=""
            start_index=3
        else:
            comframe["FrameName"]=List[3]
            start_index=4
        comframe["Dir"]=List[start_index]
        comframe["DLC"]=bytes().fromhex(List[start_index+1].zfill(2))[0]
        targetinfo=dict()
        datalsit=[]
        for index in range(0,int(List[start_index+1])):
            #print(hex(bytes.fromhex(List[start_index+2+index])[0]))
            datalsit.append(hex(bytes.fromhex(List[start_index+2+index])[0]))
        comframe["Data"]=datalsit
#         startindex=start_index+2+int(List[start_index+1])
#         targetList=[["checksum","="],["header","time","="],["full","time","="],["SOF","="]
#                    ,["BR","="],["break","="],["EOH","="],["EOB","="],["sim","="]
#                    ,["EOF","="],["RBR","="],["HBR","="],["HSO","="],["RSO","="],["CSM","="]] 
#         datanumList=[1,1,1,1,1,2,1,8,1,1,1,1,1,1,1]
#         formatList=["hex","int","int","double","int","int","double","double","int","double","int","double","int","int","string"]
#         startindexList=[]
#         startindex_1=startindex
#         for i in range(0,len(targetList)):
#             key=""
#             for subitem in targetList[i]:
#                 if(subitem != "="):
#                     key=key+subitem
#             
#             comframe[key]=self.GetValue(List, startindex_1, datanumList[i], formatList[i])    
#             startindex_1=startindex_1+len(targetList[i])+datanumList[i]
#             #startindexList.append(startindex_1)
#             
                        

            
        
#         print(comframe)
        return comframe
    def GetVariInfo(self,infoList):
        pass
    def GetSpecInfo(self,infoList):
        pass
        return True
    def GetUserInfoFrame(self):
        UserFrame={
            "Time":"",
            "Chn":"Li",
            "ID":"0x27",
            "FrameName":"", #::IO::VN1600_1::DIN0 ::IO::VN1600_1::AIN
            "Data":""
        }
        userframe=dict()            
        for index in range(0,len(self.comframe)):
            subuserframe=dict()
            for item in (self.comframe[str(index)]):
                for key in UserFrame:
                    if(key==item):
                        subuserframe[item]=self.comframe[str(index)][item]     
            
            userframe[str(index)]=subuserframe
#             print(userframe[str(0)])
            
        for index in range(0,len(userframe)):
            print(str(index)+" : ")
            print(userframe[str(index)])
#         print(userframe)
if __name__ =="__main__":
    filepath=r"C:\Project\RLT_GEELY_20R2\01_PlanSpec_RLFS\Activationstrategy Lightsensor"
    filename="CANoeTrace.asc"
    obj=ASCFile()
    obj.Asc2Txt(filepath, filename)
    List=obj.GetFrameInfo()
    List_1=List[0]
    obj.GetCommFrameInfo(List_1)
    obj.GetUserInfoFrame()