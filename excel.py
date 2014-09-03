#!/usr/bin/env python3

from __future__ import print_function
from time import strftime, localtime
import datetime
import xlsxwriter


def create_excel_report(dest_filename,MeterType, TxId,sDate,eDate,hDetails, dDetails, hProfile, dProfile, 
                       hEvents, dEvents, OnePhase):    
    """ Creates the excel file
    """
    workbook = xlsxwriter.Workbook(dest_filename)
    
    #-----------------
    #Excel Formats
    #-----------------
    sDateObj = datetime.datetime.strptime(sDate, '%Y-%m-%d')
    eDateObj = datetime.datetime.strptime(eDate, '%Y-%m-%d')
    cSiteName = ' for Tx.'+TxId
    date_format = workbook.add_format({'num_format': 'dd-mm-yy hh:mm'})
    fBU = workbook.add_format()
    fBU.set_bold()
    fBU.set_underline()

    #-----------------
    #Details
    #-----------------
    ws0 = workbook.add_worksheet('Details')
    ws0.write('A2','SITE POWER QUALITY REPORT',fBU)
    gTime = strftime("%Y-%m-%d %H:%M:%S", localtime())
    ws0.write('A3','Report generated at '+gTime)
    ws0 = dump_vertical_data(ws0,hDetails,dDetails,3)
    ws0.set_column('A:A',25)
    
    #-----------------
    #Profile
    #-----------------
    numProfileRows = str(len(dProfile)+1)
    ws1 = workbook.add_worksheet('Profile')
    ws1 = dump_date_data(ws1,hProfile,dProfile,date_format)
    ws1.set_column('A:A',15)

    #-----------------
    #Events
    #-----------------
    ws2 = workbook.add_worksheet('Events')
    ws2 = dump_data(ws2,hEvents,dEvents)
    ws2.set_column('A:B',20)   
    
    ##################################################################################
    #-----------------
    #V
    #-----------------
    ws3 = workbook.add_worksheet('V')
    ws3 = create_histogram_V(ws3,numProfileRows,OnePhase,fBU)

    #-----------------
    #THD
    #-----------------
    if MeterType == 'EDMI':
        ws4 = workbook.add_worksheet('THD')
        ws4 = create_histogram_THD(ws4,numProfileRows,OnePhase,fBU)
    
    #-----------------
    #U
    #-----------------
    if OnePhase == False:
        ws5 = workbook.add_worksheet('U')
        ws5 = create_histogram_U(ws5,numProfileRows,fBU)
    
    
    ##################################################################################
    cVars = numProfileRows,cSiteName, sDateObj, eDateObj
    
    #-----------------
    #ProfileVoltsG
    #-----------------
    wsc1 = workbook.add_chartsheet('ProfileVoltsG')
    c1 = create_chart_scatter(workbook,'ProfileVoltsG',cVars,OnePhase)
    wsc1.set_chart(c1) #Place Chart
    
    #-----------------
    #ProfileTHDG
    #-----------------
    if MeterType == 'EDMI':
        wsc2 = workbook.add_chartsheet('ProfileTHDG')
        c2 = create_chart_scatter(workbook,'ProfileTHDG',cVars,OnePhase)
        wsc2.set_chart(c2) #Place Chart

    #-----------------
    #ProfileUG
    #-----------------
    if OnePhase == False:
        wsc3 = workbook.add_chartsheet('ProfileUG')
        c3 = create_chart_scatter(workbook,'ProfileUG',cVars,OnePhase)
        wsc3.set_chart(c3) #Place Chart

    ##################################################################################
    #-----------------
    #VG
    #-----------------
    wsc4 = workbook.add_chartsheet('VG')
    c4 = create_chart_histogram(workbook,'V','74','Voltage Frequency Distribution'+cSiteName,'Voltage')
    wsc4.set_chart(c4) #Place Chart

    #-----------------
    #THDG
    #-----------------
    if MeterType == 'EDMI':
        wsc5 = workbook.add_chartsheet('THDG')
        c5 = create_chart_histogram(workbook,'THD','84','THD Frequency Distribution'+cSiteName,'THD (%)')
        wsc5.set_chart(c5) #Place Chart

    #-----------------
    #UG
    #-----------------
    if OnePhase == False:
        wsc6 = workbook.add_chartsheet('UG')
        c6 = create_chart_histogram(workbook,'U','84','Unbalance Frequency Distribution'+cSiteName,'Unbalance')
        wsc6.set_chart(c6) #Place Chart

    #-----------------
    #Save File
    #-----------------
    
    workbook.close()
    
    return dest_filename


def create_chart_scatter(workbook,cname,cVars,OnePhase):
    """Create scatter plot
    """
    
    numProfileRows,cSiteName, sDateObj, eDateObj = cVars
    c = workbook.add_chart({'type': 'scatter',
                            'subtype': 'straight'})
                                 
    if cname == 'ProfileUG':
        c.add_series({
            'name':       '=Profile!$H$1',
            'categories': '=Profile!$A$2:$A$'+numProfileRows,
            'values':     '=Profile!$H$2:$H$'+numProfileRows,
        })
        
        c.set_y_axis({'name': 'Unbalance (%)'})
        c.set_title ({'name': 'Unbalance Profile'+cSiteName})
    elif cname == 'ProfileTHDG':
        c.add_series({
        'name':       '=Profile!$E$1',
        'categories': '=Profile!$A$2:$A$'+numProfileRows,
        'values':     '=Profile!$E$2:$E$'+numProfileRows,
        })
        if OnePhase == False:
            c.add_series({
                'name':       '=Profile!$F$1',
                'categories': '=Profile!$A$2:$A$'+numProfileRows,
                'values':     '=Profile!$F$2:$F$'+numProfileRows,
            })
            c.add_series({
                'name':       '=Profile!$G$1',
                'categories': '=Profile!$A$2:$A$'+numProfileRows,
                'values':     '=Profile!$G$2:$G$'+numProfileRows,
            })
        c.set_title ({'name': 'THD Profile'+cSiteName})
        c.set_y_axis({'name': 'THD (%)'})
        
    elif cname == 'ProfileVoltsG':
        c.add_series({
        'name':       '=Profile!$B$1',
        'categories': '=Profile!$A$2:$A$'+numProfileRows,
        'values':     '=Profile!$B$2:$B$'+numProfileRows,
        })
        if OnePhase == False:
            c.add_series({
                'name':       '=Profile!$C$1',
                'categories': '=Profile!$A$2:$A$'+numProfileRows,
                'values':     '=Profile!$C$2:$C$'+numProfileRows,
            })
            c.add_series({
                'name':       '=Profile!$D$1',
                'categories': '=Profile!$A$2:$A$'+numProfileRows,
                'values':     '=Profile!$D$2:$D$'+numProfileRows,
            })    
        c.set_title ({'name': 'Voltage Profile'+cSiteName})
        c.set_y_axis({'name': 'Voltage'})

    #Generic settings
    c.set_x_axis({'name': 'Date',
                   'num_font': {'rotation': -45}, 
                   'date_axis': True,
                   'min': sDateObj,
                   'max': eDateObj,
                   'num_format': 'dd/mm/yyyy',
    })
    c.set_legend({'position': 'top'})
    c.set_style(2)
    c.set_size({'width': 900, 'height': 500})
        
    return c

    
def create_chart_histogram(workbook,sheetName,lastRow,title,axisTitle):
    """ Create histogram chart object
    """
    c = workbook.add_chart({'type': 'column',
                                 'subtype': 'stacked'})
    c.set_style(2)
    c.add_series({
        'name':       '='+sheetName+'!$F$3',
        'categories': '='+sheetName+'!$A$4:$A$'+lastRow,
        'values':     '='+sheetName+'!$F$4:$F$'+lastRow,
        'fill':   {'color': 'green'},
        'border': {'color': 'black'},
        'gap':        0,        
    })

    c.add_series({
        'name':       '='+sheetName+'!$G$3',
        'categories': '='+sheetName+'!$A$4:$A$'+lastRow,
        'values':     '='+sheetName+'!$G$4:$G$'+lastRow,
        'fill':   {'color': 'red'},
        'border': {'color': 'black'},
        'gap':        0, 
    })
    
    c.set_title ({'name': title})
    c.set_x_axis({'name': axisTitle,
                  'num_font': {'rotation': -45},
    })
    c.set_y_axis({'name': 'Freq. of Occurance (%)',
                  'num_format': '0%',
                  })
    c.set_plotarea({
        'border': {'color': 'black', 'width': 1},
        'fill':   {'color': '#FFFFC2'}
    })  
    c.set_legend({'position': 'top'})

    return c
        

def dump_data(ws,headings,data):
    """ Iterate over the data and write it out row by row.
    """
    for i, colVal in enumerate(headings):
        ws.write(0,i,colVal)
    for i, row in enumerate(data):
        for j, colVal in enumerate(row):
            ws.write(i+1,j,colVal)
    return ws


def dump_vertical_data(ws,headings,data,sRow):
    """ Iterate over the data and write it out row by row.
    """
    for i, colVal in enumerate(headings):
        ws.write(i+sRow,0,colVal)
    for i, row in enumerate(data):
        for j, colVal in enumerate(row):
            ws.write(j+sRow,i+1,colVal)
    return ws


def dump_date_data(ws,headings,data,date_format):
    """ Iterate over the data and write it out row by row.
        First column should be treated as a date
    """
    for i, colVal in enumerate(headings):
        ws.write(0,i,colVal)
    for i, row in enumerate(data):
        for j, colVal in enumerate(row):
            if j == 0:
                try:
                    date_time = datetime.datetime.strptime(colVal, '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    date_time = datetime.datetime.strptime(colVal, '%Y-%m-%d')
                ws.write_datetime('A'+str(i+2), date_time, date_format)
            else:
                ws.write(i+1,j,colVal)
    return ws


def create_histogram_V(ws,numProfileRows,OnePhase,fBU):
    """ Iterate over the data and write it out row by row.
    """

    #Headings
    ws.write(1,0,'Voltage Distribution Data',fBU)
    cLR = 73 #Last row to apply formulas to

    if OnePhase == True:
        ws.write_formula('D2', '=COUNT(Profile!B2:B'+numProfileRows+')')
        arFormula = '{=FREQUENCY(Profile!B2:B'+numProfileRows +',V!A4:A'+str(cLR)+')}'
    else:
        ws.write_formula('D2', '=3*COUNT(Profile!B2:B'+numProfileRows+')')
        arFormula = '{=FREQUENCY(Profile!B2:D'+numProfileRows +',V!A4:A'+str(cLR)+')}'

    ws.write(1,5,'Limit')
    ws.write(1,6,0.94)
    ws.write(1,7,1.06)
    icFormula = '=IF(AND(Arow>=(240*$G$2),Arow<=(240*$H$2)),Drow,0)'
    ocFormula = '=IF(AND(Arow>=(240*$G$2),Arow<=(240*$H$2)),0,Drow)'
    
    #A - Bins
    for i, colVal in enumerate(range(200,271)):
        ws.write(i+3,0,colVal)

    #Rows B to G
    ws = create_histogram_ext(ws,cLR,arFormula,icFormula,ocFormula)
    return ws

def create_histogram_THD(ws,numProfileRows,OnePhase,fBU):
    """ Iterate over the data and write it out row by row.
    """

    #Headings
    
    ws.write(1,0,'THD Distribution Data',fBU)
    cLR = 83 #Last row to apply formulas to

    if OnePhase == True:
        ws.write_formula('D2', '=COUNT(Profile!B2:B'+numProfileRows+')')
        arFormula = '{=FREQUENCY(Profile!E2:E'+numProfileRows +',THD!A4:A'+str(cLR)+')}'
    else:
        ws.write_formula('D2', '=3*COUNT(Profile!B2:B'+numProfileRows+')')
        arFormula = '{=FREQUENCY(Profile!E2:G'+numProfileRows +',THD!A4:A'+str(cLR)+')}'

    ws.write(1,5,'Limit')
    ws.write(1,6,8)
    icFormula = '=IF(Arow<$G$2,Drow,0)'
    ocFormula = '=IF(Arow<$G$2,0,Drow)'
    
    #A - Bins
    bin_list = []
    i = 0
    while i <= 10:
        bin_list.append(i)
        i += 0.125
    for i, colVal in enumerate(bin_list):
        ws.write(i+3,0,colVal)

    #Rows B to G
    ws = create_histogram_ext(ws,cLR,arFormula,icFormula,ocFormula)
    
    return ws

def create_histogram_U(ws,numProfileRows,fBU):
    """ Iterate over the data and write it out row by row.
    """

    #Headings
    ws.write(1,0,'Unbalance Distribution Data',fBU)
    cLR = 83 #Last row to apply formulas to

    ws.write_formula('D2', '=COUNT(Profile!B2:B'+numProfileRows+')')
    arFormula = '{=FREQUENCY(Profile!H2:H'+numProfileRows +',U!A4:A'+str(cLR)+')}'

    ws.write(1,5,'Limit')
    ws.write(1,6,2.55)
    icFormula = '=IF(Arow<$G$2,Drow,0)'
    ocFormula = '=IF(Arow<$G$2,0,Drow)'
    
    #A - Bins
    bin_list = []
    i = 0
    while i <= 4:
        bin_list.append(i)
        i += 0.05
    for i, colVal in enumerate(bin_list):
        ws.write(i+3,0,colVal)

    #Rows B to G
    ws = create_histogram_ext(ws,cLR,arFormula,icFormula,ocFormula)
    
    return ws

def create_histogram_ext(ws,cLR,arFormula,icFormula,ocFormula):
    """ Extra Columns other than Bins (which are sheet specific)
    """
    
    #headings
    headings = ['bins','occurances in bin','cumulative total','occurances in bin (%)','cumulative total (%)','in compliance','out of compliance']
    for i, colVal in enumerate(headings):
        ws.write(2,i,colVal)
        
    #B - Occurrence Array
    ws.write_array_formula('B4:B'+str(cLR)+'',arFormula)
    ws.write_formula('B'+str(cLR+1)+'','=IF(C'+str(cLR)+'<D2,D2-C'+str(cLR)+',0)')

    #C - Cumulative
    ws.write_formula('C4', '=B4')
    for i in range(5,cLR+2):
        cFormula = '=B'+str(i)+'+C'+str(i-1)
        ws.write_formula('C'+str(i),cFormula)

    for i in range(4,cLR+2):
        #D - occurances in bin (%)
        cFormula = '=B'+str(i)+'/D2'
        ws.write_formula('D'+str(i),cFormula)
        
        #E - Cumulative Perc
        cFormula = '=C'+str(i)+'/D2'
        ws.write_formula('E'+str(i),cFormula)  
        
        #F - In Compliance
        cFormula = icFormula.replace('row',str(i))
        ws.write_formula('F'+str(i),cFormula)
        
        #G - Out of Compliance
        cFormula = ocFormula.replace('row',str(i))
        ws.write_formula('G'+str(i),cFormula)
        
    return ws


