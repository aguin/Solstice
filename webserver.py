from flask import Flask, render_template, request, make_response, jsonify
import sqlite3
import datetime
import io
import os

import database
import excel
import getchartdata

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html',
                           title='Solstice')


@app.route('/about/')
def about():
    return render_template('about.html')


def chart_meter_line(meterId):
    h, d = get_meter_readings(meterId)
    x = []; y1 = []; y2 = []; y3 = []
    for row in d:
        x.append(datetime.datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S'))
        y1.append(row[1])
        y2.append(row[2])
        y3.append(row[3])
 
    fig=Figure()
    ax=fig.add_subplot(111)
    ax.plot_date(x, y1, '-')
    ax.plot_date(x, y2, '-')
    ax.plot_date(x, y3, '-')                
    fig.autofmt_xdate()

    return fig           
            
@app.route('/meters/')
@app.route('/meters/<meterId>/')
def meters(meterId=None):
    """ Generates the required data for the meters page
    """
    if meterId is None:
        h, d = get_meter_list()
        tableMeterList = {'headings':h, 'data':d}
        return render_template('meters.html', tableMeterList=tableMeterList)
    else:

        lat, lon = get_meter_coords(meterId)
        location = {'lat':lat, 'lon':lon}
        
        settings_filename = os.path.abspath('settings/dbExample.json')
        query_filename = os.path.abspath('sql/MonthlyReports.sql')
        params_dict = {'METERID': meterId}
        h, d = database.run_query(settings_filename, query_filename, params_dict)
        reports = {'headings':h, 'data':d}
        
        settings_filename = os.path.abspath('settings/dbExample.json')
        query_filename = os.path.abspath('sql/Last10Events.sql')
        params_dict = {'METERID': meterId}
        h, d = database.run_query(settings_filename, query_filename, params_dict)
        tableLast10Events = {'headings':h, 'data':d}        
        
        return render_template('meter.html', meterId=meterId,
                             tableLast10Events=tableLast10Events, 
                             location=location, reports=reports)


@app.route('/chartdata/')
@app.route('/chartdata/<meterId>.json')
def chartdata(meterId=None):
    if meterId is None:
        return 'json chart api'
    else:
        if request.method == 'GET':
            params = request.args.to_dict()
        elif request.method == 'POST':
            params = request.form.to_dict()
        else:
            params = {}
        
        cType = params['cType']
        if cType == 'volt':
            flotData = getchartdata.get_voltage_chart_data(meterId)
        elif cType == 'thd':
            flotData = getchartdata.get_thd_chart_data(meterId)
        elif cType == 'unbal':
            flotData = getchartdata.get_unbal_chart_data(meterId)            
        elif cType == 'events':
            flotData = getchartdata.get_events_chart_data(meterId)        
        return jsonify(flotData)
    

@app.route('/reports/')
@app.route('/reports/<meterId>', methods=['POST', 'GET'])
def reports(meterId=None):
    """ Creates excel .xlsx reports for a given date range
    """
    if meterId is None:
        return 'Report API'
    else:
        if request.method == 'GET':
            params = request.args.to_dict()
        elif request.method == 'POST':
            params = request.form.to_dict()
        else:
            params = {}
        try:
            sDate = params['sDate']
            eDate = params['eDate']    
        except KeyError:
            return 'ERROR: URL must be in form meterNo?sDate=2014-06-01&eDate=2014-06-02'

        settings_filename = os.path.abspath('settings/dbExample.json')
        query_filename = os.path.abspath('sql/MeterReadings.sql')
        params_dict = {'METERID': meterId, 'SDATE':sDate, 'EDATE':eDate}
        hProfile, dProfile = database.run_query(settings_filename, query_filename, params_dict)

        OnePhase = False
        
        settings_filename = os.path.abspath('settings/dbExample.json')
        query_filename = os.path.abspath('sql/MeterEvents.sql')
        params_dict = {'METERID': meterId, 'SDATE':sDate, 'EDATE':eDate}
        hEvents, dEvents = database.run_query(settings_filename, query_filename, params_dict)
        filePath = excel.create_excel_report(meterId, sDate, eDate, OnePhase, hProfile, dProfile, 
                       hEvents, dEvents)
                       
        return render_template('report_download.html', filePath=filePath)
                       
                       
                       
def get_meter_list():
    settings_filename = os.path.abspath('settings/dbExample.json')
    query_filename = os.path.abspath('sql/ListMeters.sql')
    params_dict = {}
    headings, rows = database.run_query(settings_filename, query_filename, params_dict)    

    return headings, rows


def get_meter_coords(meterId):
    dbPath = 'data/meters.db'
    conn = sqlite3.connect(dbPath)
    curs = conn.cursor()

    settings_filename = os.path.abspath('settings/dbExample.json')
    query_filename = os.path.abspath('sql/MeterCoords.sql')
    params_dict = {'METERID': meterId}
    headings, rows = database.run_query(settings_filename, query_filename, params_dict)
        

    lat = rows[0][0]
    lon = rows[0][1]

    return lat, lon

