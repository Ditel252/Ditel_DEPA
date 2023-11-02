import sys
import os
import openpyxl
import pandas
import xlwings

#定数
VERSION:str = "v0.0.1"  #バージョン
READ_DATA_START_ROW = 1251

class _terminal:    #ターミナル出力関係
    def __init__(self):
        self._printStatus = None
    
    def print(self, _status:bool, _message:str):
        if(_status):
            self._printStatus = "  OK  "
        else:
            self._printStatus = "FAILED"
        
        print("[{}]  {}".format(self._printStatus, _message))
        
class _cell:
    def __init__(self, _readFilePath, _readSheetName):
        self._excelFile = xlwings.Book(_readFilePath)
        self._workSheet = self._excelFile.sheets[_readSheetName]
        
    def getValue(self, _readCellAddreess):
        try:
            return float(self._workSheet.range(_readCellAddreess).value)
        except:
            return self._workSheet.range(_readCellAddreess).value
    
    def end(self):
        self._excelFile.close()
        
class _dataBase:    #データベースの読み取り
    def __init__(self, _dataBaseFilePath:str):
        self._excelFilePath:str = _dataBaseFilePath
        self._workBook = openpyxl.load_workbook(filename=self._excelFilePath, read_only=True)
        
        self.sheet = self._workBook["DataBase"]
    
    def readCellData(self, _column:int, _row:int):  #_colum:A,B,C… row:1,2,3…
        return self.sheet.cell(column=_column, row=_row).value
    
    def end(self):
        self._workBook.close
        
class _oscilloscopeData:    #csvからxlsxに変換する
    def __init__(self, _dataDirectoryPath, _oscilloscopeDataName):
        self._inputOscilloscopeData1Name:str = "{}\\{}\\F{}{}.CSV".format(_dataDirectoryPath, _oscilloscopeDataName, _oscilloscopeDataName[3:], READ_DATA1)
        self._inputOscilloscopeData2Name:str = "{}\\{}\\F{}{}.CSV".format(_dataDirectoryPath, _oscilloscopeDataName, _oscilloscopeDataName[3:], READ_DATA2)
        
        self._outputOscilloscopeData1Name:str = "{}\\{}\\F{}{}.xlsx".format(_dataDirectoryPath, _oscilloscopeDataName, _oscilloscopeDataName[3:], READ_DATA1)
        self._outputOscilloscopeData2Name:str = "{}\\{}\\F{}{}.xlsx".format(_dataDirectoryPath, _oscilloscopeDataName, _oscilloscopeDataName[3:], READ_DATA2)
        
        self.baseCSVFile1 = pandas.read_csv(self._inputOscilloscopeData1Name)
        self.baseCSVFile2 = pandas.read_csv(self._inputOscilloscopeData2Name)
        
    def convert(self):
        self.baseCSVFile1.to_excel(self._outputOscilloscopeData1Name)
        self.baseCSVFile2.to_excel(self._outputOscilloscopeData2Name)

terminal = _terminal()

class _driveApproximateFomula:  #近似式を導出する
    def __init__(self, _readFilePath:str, _frequency:int):
        self._readFilePeriod:float = 1.0 / float(_frequency)
        
        self.excelFilePath = _readFilePath
        self._workBook = openpyxl.load_workbook(_readFilePath)
        self.mainSheet = self._workBook["Sheet1"]
        
        self.readDataEndRow:int = None
        
    def _findRangeOf1Cycle(self):
        _nowColumn = READ_DATA_START_ROW
        
        while(True):
            if(float(self.mainSheet.cell(column=5, row=_nowColumn).value) > self._readFilePeriod):
                _nowColumn -= 1
                break
            else:
                _nowColumn += 1
        
        terminal.print(True, "Find Range Of 1 Cycle")
        
        self.readDataEndRow = _nowColumn
        self.copyDataEndRow = None
        
        self._workBook.close()
    
    def _extractRelevantValue(self):
        self._workBook.create_sheet(title="forCalculation")
        
        self.calculationSheet = self._workBook["forCalculation"]
        
        originalColumn:int = 5
        originalRow:int = READ_DATA_START_ROW
        toColumn:int = 1
        toRow:int = 1
        
        while (True):
            if(originalRow <= self.readDataEndRow):
                self.calculationSheet.cell(column=toColumn, row=toRow).value = self.mainSheet.cell(column=originalColumn, row=originalRow).value
                
                originalRow += 1
                toRow += 1
            else:
                break
        
        originalRow = READ_DATA_START_ROW
        toRow = 1
        originalColumn += 1
        toColumn += 1
            
        while (True):
            if(originalRow <= self.readDataEndRow):
                self.calculationSheet.cell(column=toColumn, row=toRow).value = self.mainSheet.cell(column=originalColumn, row=originalRow).value
                
                originalRow += 1
                toRow += 1
            else:
                break
        
        if((toRow - 2) == (self.readDataEndRow - READ_DATA_START_ROW)):
            terminal.print(True, "Copy Extract For The Relevant Value")
            self.copyDataEndRow = toRow - 1
        else:
            terminal.print(False, "Copy Extract For The Relevant Value")
            exit()
    
    def _findApproximateFomula(self):
        self.calculationSheet["E2"] = "近似式"
        self.calculationSheet["E3"] = "指数"
        self.calculationSheet["F3"] = "近似式の係数"
        
        for _i in range(10, -1, -1):
            self.calculationSheet.cell(column=5, row=(14 -_i)).value = int(_i)
            
            if(_i != 0):
                self.calculationSheet.cell(column=6, row=(14 -_i)).value = "=INDEX(LINEST(B1:B{:d},A1:A{:d}^{{10,9,8,7,6,5,4,3,2,1}}),1,E{:d})".format(self.copyDataEndRow, self.copyDataEndRow, 14 -_i)
            else:
                self.calculationSheet.cell(column=6, row=(14 -_i)).value = "=INDEX(LINEST(B1:B{:d},A1:A{:d}^{{10,9,8,7,6,5,4,3,2,1}}),1,11)".format(self.copyDataEndRow, self.copyDataEndRow)
                
        for _i in range(1, self.copyDataEndRow + 1, 1):
            self.calculationSheet.cell(column=3, row=_i).value = "=F4*(A{}^E4)+F5*(A{}^E5)+F6*(A{}^E6)+F7*(A{}^E7)+F8*(A{}^E8)+F9*(A{}^E9)+F10*(A{}^E10)+F11*(A{}^E11)+F12*(A{}^E12)+F13*(A{}^E13)+F14".format(_i, _i, _i, _i, _i, _i, _i, _i, _i, _i)
    
    def _findMaximumTime(self):
        self.calculationSheet["E16"] = "yの最大値"
        
        self.calculationSheet["E17"] = "=MAX(C1:C{:d})".format(self.copyDataEndRow)
    
    def end(self):
        self._workBook.save("{}_TemporaryData.xlsx".format(self.excelFilePath))
        self._workBook.close()
        
        return "{}_TemporaryData.xlsx".format(self.excelFilePath)

class _findPhasePeakValue:
    def __init__(self, _readFilePath:str):
        self._excelFilePath:str = _readFilePath
        self._workBook = openpyxl.load_workbook(_readFilePath)
        self._workSheet = self._workBook["forCalculation"]
        
        self.cell = _cell(_readFilePath, "forCalculation")
        
        self.yMaxValue:float = self.cell.getValue("E17")
        
        self.phasePeakValue:float = None
    
    def _findPhase(self):
        _readRow:int = 1
        
        while(True):
            if(self.cell.getValue("C{:d}".format(_readRow)) == None):
                terminal.print(False, "Find Peak Value")
                self._workBook.close()
                exit()
            elif(self.cell.getValue("C{:d}".format(_readRow)) == self.yMaxValue):
                terminal.print(True, "Find Peak Value")
                self.phasePeakValue = self._workSheet.cell(column=1, row=_readRow).value
                print(_readRow)
                break
            
            _readRow += 1
        
    
    def end(self):
        self._workBook.close()
        self.cell.end()
        return self.phasePeakValue
        
        
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
    
oscilloscopeData = _oscilloscopeData(DATA_DIRECTORY_PATH, dataBase.readCellData(2, 2))

oscilloscopeData.convert()

approximateFomula = _driveApproximateFomula(oscilloscopeData._outputOscilloscopeData1Name, dataBase.readCellData(1, 2))

approximateFomula._findRangeOf1Cycle()

approximateFomula._extractRelevantValue()

approximateFomula._findApproximateFomula()

approximateFomula._findMaximumTime()

phasePeakValue = _findPhasePeakValue(approximateFomula.end())

phasePeakValue._findPhase()

print("{:.5g}".format(float(phasePeakValue.end())))

dataBase.end()