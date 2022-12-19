# pip install pywin32

from win32com.client import Dispatch, GetObject

class Excel:
    def __init__(self, newIntance=True, disableProperties=False):
        if newIntance:
            self.xl = Dispatch('Excel.Application')
        else:
            self.xl = GetObject(Class='Excel.Application')
        self.disableProperties = disableProperties
        if self.disableProperties: self.systemProperties(False)
    
    def __del__(self):
        if self.disableProperties: self.systemProperties(True)

    def getExcelApp(self):
        return self.xl

    def vbaExecute(self, textoMacro):
        vbComp = self.xl.VBE.ActiveVBProject.VBComponents.Add(1)
        vbComp.CodeModule.AddFromString('Sub Temp()' + chr(10) + textoMacro + chr(10) + 'End Sub')
        self.xl.Run(vbComp.Name + '.Temp')
        self.xl.VBE.ActiveVBProject.VBComponents.Remove(vbComp)

    def getWorkbook(self, wbNameOrIndex):
        try:
            return self.xl.Workbooks(wbNameOrIndex)
        except:
            raise Exception(f'Workbook {wbNameOrIndex} not founded')

    def openExcelFile(self, xlFile, wsNameOrIndex=None, IsVisible=False):
        self.xl.Visible = IsVisible
        wb = self.xl.Workbooks.Open(xlFile)
        try:
            wb.AutoSaveOn = False
        except:
            pass
        if not wsNameOrIndex is None:
            ws = wb.Sheets(wsNameOrIndex)
        else:
            ws = None

        return wb, ws

    def getSheet(self, wb, wsNameOrIndex):
        return wb.Sheets(wsNameOrIndex)

    def closeExcelFile(self, wb, IsToSave=True):
        wb.Close(IsToSave)
        if self.xl.Workbooks.Count == 0: self.xl.Quit()

    def loadDataToListObject(self, ws, loNameOrIndex, aData, isIncremental=False, redimColumns=False, header=None):
        lo = ws.ListObjects(loNameOrIndex)
        lo.AutoFilter.ShowAllData()
        
        if isIncremental:
            if not lo.DataBodyRange is None: 
                aData.extend(lo.DataBodyRange.Value2)
                lo.DataBodyRange.Delete()
        else:
            if not lo.DataBodyRange is None: 
                lo.DataBodyRange.Delete()

        numRows = len(aData) + lo.HeaderRowRange.Row #lo.HeaderRowRange.Row = 2 linhas, uma em branco + linha de cabeçalho se tabela inicia na linha 2
        if redimColumns:
            numCols = len(aData[0]) + (lo.HeaderRowRange.Column - 1) #(lo.HeaderRowRange.Column - 1) = -1 coluna em branco se começando da coluna 'B'
        else:
            numCols = lo.ListColumns.Count + lo.HeaderRowRange.Column - 1

        lo.Resize(ws.Range(ws.Cells(lo.HeaderRowRange.Row, lo.HeaderRowRange.Column), ws.Cells(numRows, numCols)))
        lo.DataBodyRange.Value2 = aData
        if not header is None: lo.HeaderRowRange.Value2 = header

        return lo

    def getArrayFromListObject(self, ws, loNameOrIndex, IsToTransposeMatrix=False):
        matrix = ws.ListObjects(loNameOrIndex).DataBodyRange.Value2
        if IsToTransposeMatrix:
            return self.tranpose2DMatrix(matrix)
        else:
            return matrix

    def getArrayFromRange(self, ws, startRow=1, startCol=1, colKey=0, hasHeader=False):
        if colKey == 0: colKey = startCol
        lastRow = ws.Cells(ws.Rows.Count, colKey).End(-4162).Row
        lastCol = ws.Cells(startRow, ws.Columns.Count).End(-4159).Column
        if hasHeader: startRow += 1
        return ws.Range(ws.Cells(startRow, startCol), ws.Cells(lastRow, lastCol)).Value2

    def save(self, wb, xlFileName=None, xlFormat=51):
        """
        Esta rotina irá salvar a pasta de trabalho informada, com o nome informado, utilizando o formato padrão xlsx.\n
        Considerar wb a pasta de trabalha já aberta que vc irá salvar como e,\n
        xlFileName o diretório e nome do arquivo.\n
        Ex: r'C\:Users\MarcioRibeiro\Desktop\CursoPython.xlsx'\n
        Para salvar em outros formatos, consulte a lista de formatos no link abaixo:\n
        https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat
        """
        if xlFileName:
            wb.SaveAs(xlFileName, xlFormat)
        else:
            wb.Save()

    def getValueFromName(self, wb, name):
        return wb.Names(name).RefersToRange.Text

    def setValueToName(self, wb, name, value):
        wb.Names(name).RefersToRange.Value = value
    
    def filterListObject(self, ws, loName, field, value):
        lo = ws.ListObjects(loName)
        lo.AutoFilter.ShowAllData()
        lo.DataBodyRange.AutoFilter(Field=field, Criteria1=f'={value}')
    
    def copyColumnDataFromListObject(self, ws, loNameOrIndex, columnIndexOrName, justVisibleRows=False):
        if justVisibleRows:
            ws.ListObjects(loNameOrIndex).ListColumns(columnIndexOrName).DataBodyRange.SpecialCells(12).Copy() # 12 = xlCellTypeVisible
        else:
            ws.ListObjects(loNameOrIndex).ListColumns(columnIndexOrName).DataBodyRange.Copy()

    def clearAutoFilter(self, ws, loNameOrIndex=None):
        try:
            ws.ShowAllData()
        except:
            if not loNameOrIndex: loNameOrIndex = 1
            ws.ListObjects(loNameOrIndex).AutoFilter.ShowAllData()
    
    def copyVisibleRowsFromListObject(self, ws, loName, columnIndexOrName):
        lo = ws.ListObjects(loName)
        lo.ListColumns(columnIndexOrName).DataBodyRange.SpecialCells(12).Copy()

    def delVisibleRowsFromsListObject(self, ws, loName, columnIndexOrName):
        lo = ws.ListObjects(loName)
        lo.ListColumns(columnIndexOrName).DataBodyRange.SpecialCells(12).EntireRow.Delete()

    def putFormulaInListColumnOfListObject(self, ws, loName, columnIndexOrName, formulaText):
        lo = ws.ListObjects(loName)
        lo.ListColumns(columnIndexOrName).DataBodyRange.Formula = formulaText

    def tranpose2DMatrix(self, matrix):
        return [[matrix[r][c] for r in range(len(matrix))] for c in range(len(matrix[0]))]

    def systemProperties(self, isToEnable):
        self.xl.ScreenUpdating = isToEnable
        self.xl.DisplayAlerts = isToEnable
        self.xl.EnableEvents = isToEnable
