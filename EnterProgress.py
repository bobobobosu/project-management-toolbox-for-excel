import os
import xlwings as xw
if __name__ == '__main__':
    currentPath = os.path.dirname(os.path.abspath(__file__))
    xb = xw.Book(os.path.join(currentPath,'TC.xlsb'))
    xb.macro('EnterProgressMinimize')()
