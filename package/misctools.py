import openpyxl
from openpyxl import load_workbook
import itertools
import string
import shutil


class MiscTools:
       
    def list_value_duplicator(obj, dupVal): #Repeats all values in a list
        obj = list(itertools.chain.from_iterable(itertools.repeat(obj, dupVal) for obj in obj))
        return obj;
    
    def col_to_num(col):
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num
            