#!usr/bin/env python
# -*- coding:utf-8 -*-

import os
import sys
from xlrd import open_workbook, Book, biffh
from xlutils.copy import copy
from __builtin__ import False

class WTXLS(Book):
    def __init__(self, srcxmlfile, cell_overwrite_ok=True, formatting_info=True):
        
        self.sheetObj = None
        self.srcxmlfile = srcxmlfile
#        self.destxmlfile = destxmlfile
#         self.sheetname = sheetname
#         self.value_row = value_row
#         self.value_col = value_col
        self.cell_overwrite_ok = True
        self.formatting_info = True
        self.op_book = open_workbook(self.srcxmlfile, formatting_info=self.formatting_info)
        self.wb = copy(self.op_book)
    
    def __del__(self):
        try:
            bsnm = os.path.basename(self.srcxmlfile)
            print "Saving my edition of the file [%s]. " % bsnm
            #self.wb.save(r'bbbbbbbbbbbbbbb.xls')
            self.wb.save(self.srcxmlfile)
        except IOError as e:
            print "File [%s] is opened or used by others. close and try again." % bsnm
            
    def is_valid_xls_file(self):
        if os.path.isfile(self.srcxmlfile):
            if self.srcxmlfile.endswith("xls"):
                try:
                    Book.encoding = "utf8" #if sheet name contains Chinese chracters, will be occure exception.
                    #self.op_book = open_workbook(self.srcxmlfile, formatting_info=self.formatting_info)
                    if self.op_book.nsheets <=0:
                        return False
                except Exception as e:
                    print "occure exception when checking file validation: ", str(e)
                    return False
                with open(self.srcxmlfile, "rb") as f:
                    f.seek(0, 2) # EOF
                    size = f.tell()
                    f.seek(0, 0) # BOF
                    if size == 0:
                        raise biffh.XLRDError("File size is 0 byte")     
            else:
                    print "File not end up with '.xls' perfix."
                    return False
        return True
    
    def get_sheet_lst(self):
        #return self.op_book.sheets()
        return self.op_book.sheet_names()
    
    def get_sht_obj(self, shtnm):
        if not self.op_book:
            self.op_book = open_workbook(self.srcxmlfile, formatting_info=self.formatting_info)
        if not shtnm in self.get_sheet_lst():
            raise IOError("The input sheet name is not in current sheet list.")
        self.sheetObj = self.op_book.sheet_by_name(shtnm)
        return self.sheetObj
    
    def get_position_value(self, shtnm, row_pos, col_pos):
        vs = sys.version.split()[0]
        if not int(vs.split('.')[0]) == 2 or int(vs.split('.')[1]) < 3:
            print "basestring type starts from version2.3 but removed from version3.0"
            return None
        
        if not isinstance(shtnm, basestring) or not isinstance(row_pos, int) or not isinstance(col_pos, int) or \
           not shtnm in self.get_sheet_lst():
            print "The input param(s) is/are not valid."
            return None
        
        if row_pos > self.get_row_len(shtnm) or col_pos > self.get_col_len(shtnm):
            print "Your request position (%s, %s) is out of line." % (row_pos, col_pos)
            return None
        
        return self.get_sht_obj(shtnm).cell(row_pos, col_pos).value
    

    def wt_xls(self, sheetname=None, **kwargs):
        '''row_x = x, col_y = y, value = [["a","b",]]'''

        if sheetname == None:
            sheetname = "Copy"
            print "Found no valid sheet names, use %s instead." % sheetname
        if sheetname not in self.get_sheet_lst():
            self.wb.add_sheet(sheetname)
            self.get_sheet_lst().append(sheetname)
        sheet_index = self.get_sheet_lst().index(sheetname)
        
        ws = self.wb.get_sheet(sheet_index)
        value = kwargs.get("value",[[]])
        rowx = kwargs.get("row_x", "NULL")
        coly = kwargs.get("col_y", "NULL")
        
        if rowx == "NULL" and coly == "NULL":
            '''write sheet from (0, 0) to (rowx, coly)'''
            for i in xrange(len(value)):
                for j in xrange(len(i)):
                    ws.write(i, j, value[i][j])
        else:
            if rowx != "NULL" and coly == "NULL":
                '''write sheet from (rowx, 0) to (rowx, n)'''
                for i in value:
                    try:
                        coly = 0
                        for j in i:
                            ws.write(rowx, coly, j)
                            coly += 1
                    except Exception as e:
                        print str(e)
                  
            else:      
                if rowx == "NULL" and coly != "NULL":
                    '''write sheet from (0, coly) to (n, coly)'''
                    for i in value:
                        try:
                            row_x = 0
                            for j in i:
                                ws.write(row_x, coly, j)
                                row_x += 1 
                        except Exception as e:
                            print str(e)
                    
                else:       
                    if rowx != "NULL" and coly != "NULL":
                        '''write sheet to (rowx, coly)'''
                        for i in value:
                            if len(i) == 1:
                                for j in i:
                                    ws.write(rowx, coly, j)
                            else:
                                ws.write(rowx, coly, i)

    def get_row_len(self, sheetname):
        '''
        return the total row lenth of a sheet named "sheetname".
        input: the sheet name, type:xxx
        output:the row lenth, type:int
        '''
        return self.get_sht_obj(sheetname).nrows
    
    def get_col_len(self, sheetname):
        '''
        return the total col lenth of a sheet named "sheetname".
        input: the sheet name, type:xxx
        output:the colume lenth, type:int
        '''
        return self.get_sht_obj(sheetname).ncols
    
    def get_row_values(self, sheetname, row_pos):
        return self.get_sht_obj(sheetname).row_values(row_pos)
    
    def get_col_values(self, sheetname, col_pos):
        return self.get_sht_obj(sheetname).row_values(col_pos)

		
def demo_WTXLS():
    VALID_HEAD_COLUMN_NAME = ('Copy2',
                              'Asset2',
                              'Put-Post-Get2',
                              '',
                              '/Copy/CopyConfigCap2.xml',
                              'Description2', 
                              'Expected Behavior2', 
                              'Output File2', 
                              'Input File2')
    line_num = 20
    srcfl = r'E:\test\tries\Farad_LEDM_Test_tmplate.xls'
    wtxml = WTXLS(srcfl)
    
    if wtxml.is_valid_xls_file():
        wtxml.wt_xls(sheetname="Copy", row_x = line_num, value=[list(VALID_HEAD_COLUMN_NAME)])
        
    for fl in getfilelist(r"E:\test\tickets",[]):
        fl = os.path.basename(fl)
        if wtxml.get_position_value("Copy", line_num+1, 4).upper()=="GET":
            wtxml.wt_xls("Copy",row_x=line_num+1, col_y=7,value=[[fl]])
        else:
            wtxml.wt_xls("Copy",row_x=line_num+1, col_y=8,value=[[fl]])
            
        line_num += 1
  
if __name__=='__main__':
    #demo_WTXLS()
	pass
