# -*- coding: utf-8 -*-
"""

https://xlsxwriter.readthedocs.io/working_with_charts.html

@author: Alex
"""

# import xlsxwriter module 
import xlsxwriter
import time
import datetime
import os
from xlsxwriter.utility import xl_rowcol_to_cell

class ExcelUtil:
    
    def __init__(self):
    
       
        # Start from the first cell. 
        # Rows and columns are zero indexed. 
        self.row = 0
        self.column = 0
    
    # Workbook() takes one, non-optional, argument   
    # which is the filename that we want to create. 
    def createFile(self, filename):
        self.filename=filename
        #workbook = xlsxwriter.Workbook('c:/pyApps/junk/chart_Line3.xlsx') 
        self.workbook = xlsxwriter.Workbook(filename)
        # The workbook object is then used to add new   
        # worksheet via the add_worksheet() method.  
        self.worksheet = self.workbook.add_worksheet() 
        print("created file=",filename)
        
    # write data to file
    def writeData(self, row,col,data):
        
        self.row=row
        self.col=col
        self.worksheet.write(row, col, data)     
         
    # Finally, close the Excel file  
    # via the close() method.  
    def closeFile(self):
        self.workbook.close() 
        self.timeStamp()
        
    def timeStamp(self):
        
        ts=time.time()
        st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d-%H_%M')
        #print 'time=',st
        L=len(self.filename)
        #take .xlsx out from file name
        FileName_1=self.filename[0:L-5]
        print ('fileName_1',FileName_1)
        print('filename',self.filename)
        #add serial Number, test tempearature and time stamp
        self.filename_stamped=FileName_1+'_'+st+'.xlsx'
        #rename the existing file
        os.rename(self.filename, self.filename_stamped )
        print ('frequency sweep results file:',self.filename_stamped)
        
    # =============== Excel Chart with xlsxwriter https://xlsxwriter.readthedocs.io/working_with_charts.html
    
    
    def writeTableDataLabels(self):
        # Write some simple text.
        self.worksheet.write('A2', 'Freq')

        # Text with formatting.
        self.worksheet.write('A3', 'Power')
        
    def createExcelChart(self):
        # here we create a line chart object . 
        #chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        #self.chart1 = self.workbook.add_chart({'type': 'line'}) 
        #draw smooth line chart
        self.chart1 = self.workbook.add_chart({'type': 'scatter', 'subtype':'smooth'}) 
    
        # freq power horizontal table
      
        
        #convert cell's row and col to format like a$26 in order to specify the last cell of the table
        EndOfRow2 = xl_rowcol_to_cell(1,self.col, row_abs=True) 
        EndOfRow3 = xl_rowcol_to_cell(2,self.col, row_abs=True) 
       
        cat_range='=Sheet1!$B$2:$'+EndOfRow2
        val_range='=Sheet1!$B$3:$'+EndOfRow3
        
        self.chart1.add_series({
             #'name': '= Sheet1!$A$1',
             'categories': cat_range, 
             'values': val_range,
            #  'trendline': {
            #  # 'type': 'moving_average',
            #  # 'period':2,
            #  'type': 'polynomial',
            #  'name': 'trend',
            #  'order': 2,
            #  'display_equation': True,
            #  'line': {
            # 'color': 'red',
            # 'width': 1,
            
            # },
             #}
        })       
        
        # Add a chart title  
        self.chart1.set_title ({'name': 'Power vs. Frequency'}) 
        # Add x-axis label 
        self.chart1.set_x_axis({'name': 'MHz',
                                'name_layout': {
                                'x': 0.34,
                                'y': 0.85,
                                }
                            }) 
            
        # Add y-axis label 
        self.chart1.set_y_axis({'name': 'dBm'})
        
        self.chart1.set_x_axis({
            #'num_font': {'italic': True},
            'major_gridlines': {
                'visible': True,
                'line': {'width': 1}
            },
            'minor_gridlines': {
                'visible': True,
                'line': {'width': 1, 'dash_type': 'dash'}
            },
        })

        # a chart is anchored to cell D2 .  
        self.worksheet.insert_chart('B4', self.chart1, {'x_offset': 25, 'y_offset': 10})
        
        
