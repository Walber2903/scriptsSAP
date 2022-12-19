# pip install pandas

import pandas as pd
import numpy as np

class Txt:
    def __init__(self, file, header, colList=[], colListIsToAdd=True, delimiter='|', delFirstRows=None, delLastRows=None, numericColList=[], encoding='iso-8859-1', thousands=None, decimal='.') -> None:
        self.file = file
        self.header = header
        self.colList = colList
        self.colListIsToAdd = colListIsToAdd
        self.delimiter = delimiter
        self.delFirstRows = delFirstRows
        self.delLastRows = delLastRows
        self.numericColList = numericColList
        self.encoding = encoding
        self.thousands = thousands
        self.decimal = decimal
        self.df = self._loadDataFrame_()

    def _loadDataFrame_(self):
        if self.colListIsToAdd and len(self.colList):
            df = pd.read_csv(self.file, sep=self.delimiter, header=self.header, usecols=self.colList, encoding=self.encoding, thousands=self.thousands, decimal=self.decimal)
        else:
            df = pd.read_csv(self.file, sep=self.delimiter, header=self.header, encoding=self.encoding, thousands=self.thousands, decimal=self.decimal)

        df.rename(columns=lambda x: x.strip(), inplace=True)
        strColumns = df.select_dtypes(['object'])
        df[strColumns.columns] = strColumns.apply(lambda x: x.str.strip())
        
        if not self.colListIsToAdd and len(self.colList):
            if type(self.colList[0]) == str:
                df.drop(self.colList, axis=1, inplace=True)
            else:
                df.drop(df.columns[self.colList], axis=1, inplace=True)

        df[df.columns[0]].replace(df.columns[0], np.nan, inplace=True)
        df.dropna(subset=[df.columns[0]], inplace=True)
        if self.delFirstRows: df = df[self.delFirstRows:]
        if self.delLastRows: df = df[:len(df)-self.delLastRows]
        
        for col in self.numericColList:
            mask = df[col].str.endswith('-')
            df.loc[mask, col] = '-' + df.loc[mask, col].str.replace('-', '')

        return df

    def getArray(self):
        return np.array(self.df)
    
    def getDataFrame(self):
        return self.df

    def getColumnsArray(self):
        return np.array(self.df.columns)
