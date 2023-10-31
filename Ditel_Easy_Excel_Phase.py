import sys
import os
import openpyxl
import pandas

#定数
VERSION:str = "v0.0.1"  #バージョン

class _terminal:
    def __init__(self):
        self._printStatus = None
    
    def print(self, _status:bool, _message:str):
        if(_status):
            self._printStatus = "  OK  "
        else:
            self._printStatus = "FAILED"
        
        print("[{}]  {}".format(self._printStatus, _message))
        
class _dataBase:
    def __init__(self, _dataBaseFilePath:str):
        self._excelFilePath:str = _dataBaseFilePath
        self.workBook = openpyxl.load_workbook(filename=self._excelFilePath, read_only=True)
        
        self.sheet = self.workBook["DataBase"]
    
    def readCellData(self, _column:int, _row:int):  #_colum:A,B,C… row:1,2,3…
        return self.sheet.cell(column=_column, row=_row).value
    
    def end(self):
        self.workBook.close
        
class _OscilloscopeData:
    def __init__(self, _dataDirectoryPath, _OscilloscopeDataName):
        self.baseCSVFile = pandas.read_csv("{}\\{}".format(_dataDirectoryPath, ))

terminal = _terminal()



print("*** Start Ditel Easy Excel Phase Contrast Program ***");
terminal.print(True, "version : {}".format(VERSION))

try:
    DATA_DIRECTORY_PATH:str = sys.argv[1]
    READ_DATA1:str = sys.argv[2]
    READ_DATA2:str = sys.argv[3]
    
    if((READ_DATA1 != "CH1") and (READ_DATA1 != "CH2") and (READ_DATA1 != "MTH")):
        raise ValueError(None)
    
    if((READ_DATA2 != "CH1") and (READ_DATA2 != "CH2") and (READ_DATA2 != "MTH")):
        raise ValueError(None)
    
    terminal.print(True, "Read Data Information")
    terminal.print(True, "Read Data Directory Path = {}".format(DATA_DIRECTORY_PATH))
    terminal.print(True, "Read Data1 Type = {}".format(READ_DATA1))
    terminal.print(True, "Read Data2 Type = {}".format(READ_DATA2))
except:
    terminal.print(False, "Read Data Information")
    print("Please input Data Directory Path")
    print('Example : python3 ./Ditel_Easy_Excel_Phase "Data Directory Path" "Data1 Type" "Data2 Type"')
    print('Only "CH1", "CH2" or "MTH" can be entered for "DATA Type" and "DATA2 Type"')
    exit()
    
DATA_BASE_FILE_PAHT:str = "{}\\{}".format(os.getcwd(), "readDataBase.xlsx")

try:
    dataBase = _dataBase(DATA_BASE_FILE_PAHT)
    
    terminal.print(True, "Read Data Base File")
except:
    terminal.print(False, "Read Data Base File")
    
    exit()
    
print(dataBase.readCellData(1, 1))

dataBase.end()