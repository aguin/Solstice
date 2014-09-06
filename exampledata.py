#!/usr/bin/python3

#System Modules
import random
import time
import os
import csv
import os
import sqlite3

def create_example_db():
    """ Create a dummy database with metering data
    """

    if not os.path.exists('data/'):
        os.makedirs('data/')
    dbPath = 'data/meters.db'
    try:
        os.remove(dbPath)
    except OSError:
        pass

    conn = sqlite3.connect(dbPath)
    curs = conn.cursor()

    # Create the meter list table
    query = 'CREATE TABLE METER_DETAILS (meter_id TEXT, meter_desc TEXT);'
    curs.execute(query) 
    query = 'INSERT INTO METER_DETAILS (meter_id, meter_desc) VALUES (?, ?);'
    for x in range(1, 5):
        meterId = x
        meterName = 'Example Meter ' + str(meterId)
        row = [meterId, meterName]
        curs.execute(query, row)

    # Create the meter locations table
    query = 'CREATE TABLE METER_LOCATIONS (meter_id TEXT, lat REAL, lon REAL);'
    curs.execute(query)
    query = 'INSERT INTO METER_LOCATIONS (meter_id, lat, lon) VALUES (?, ?, ?);'
    for x in range(1, 5):
        lat, lon = create_random_coords()
        row = [x, lat, lon]
        curs.execute(query, row)

    # Create the meter readings table
    query = 'CREATE TABLE READINGS (meter_id TEXT, measdate DATETIME, v1 REAL, v2 REAL, v3 REAL, thd1 REAL, thd2 REAL, thd3 REAL, unbal REAL, PRIMARY KEY (meter_id, measdate));'
    curs.execute(query)
    query = 'INSERT INTO READINGS (meter_id, measdate, v1, v2, v3, thd1, thd2, thd3, unbal) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);'
    startDate = "2012-01-01 00:00:00"
    endDate = "2014-07-01 00:00:00"
    for x in range(1, 5):
        meterId = x
        for y in range(1,10000):
            dateFormat = '%Y-%m-%d %H:%M:%S'
            measdate = strTimeProp(startDate, endDate, dateFormat, random.random())
            v1 = create_random_voltage()
            v2 = create_random_voltage()
            v3 = create_random_voltage()
            thd1 = create_random_thd()
            thd2 = create_random_thd()
            thd3 = create_random_thd()                        
            unbal = calc_unbalance([v1,v2,v3])
            row = [meterId, measdate, v1, v2, v3, thd1, thd2, thd3, unbal]
            try:
                curs.execute(query, row)
            except sqlite3.Error as e:
                pass # Random date wasn't random enough


    # Create the meter events
    query = 'CREATE TABLE EVENTS (meter_id TEXT, event_start DATETIME, event_end DATETIME, event_type REAL, amplitude REAL, duration REAL, phases TEXT, PRIMARY KEY (meter_id, event_start, event_type));'
    curs.execute(query)
    query = 'INSERT INTO EVENTS (meter_id, event_start, event_end, event_type, amplitude, duration, phases) VALUES (?, ?, ?, ?, ?, ?, ?);'
    startDate = "2012-01-01 00:00:00"
    endDate = "2014-07-01 00:00:00"
    for x in range(1, 5):
        meterId = x
        # TODO Make the event values actually line up with readings
        for y in range(1,250):
            dateFormat = '%Y-%m-%d %H:%M:%S'
            event_start = strTimeProp(startDate, endDate, dateFormat, random.random())
            event_end = strTimeProp(event_start, endDate, dateFormat, random.random())
            event_type = 'SAG'
            amplitude = 210.0
            duration = 100.0
            phases = 'ABC'
            row = [meterId, event_start, event_end, event_type, amplitude, duration, phases]
            try:
                curs.execute(query, row)
            except sqlite3.Error as e:
                pass # Random date wasn't random enough
        for y in range(1,250):
            dateFormat = '%Y-%m-%d %H:%M:%S'
            event_start = strTimeProp(startDate, endDate, dateFormat, random.random())
            event_end = strTimeProp(event_start, endDate, dateFormat, random.random())
            event_type = 'SWL'
            amplitude = 260.0
            duration = 50.0
            phases = 'ABC'
            row = [meterId, event_start, event_end, event_type, amplitude, duration, phases]
            try:
                curs.execute(query, row)
            except sqlite3.Error as e:
                pass # Random date wasn't random enough                


    conn.commit()
    conn.close()

    return True



def calc_unbalance(a):
    avg = sum(a, 0.0) / len(a)
    max_dev = max(abs(el - avg) for el in a)
    return (max_dev/avg)

def create_random_coords():
    """ Creates a random lat lon combo
    """
    lat = round(-27.0 - random.random()*3,5)
    lon = round(153 - random.random()*3,5)

    return lat, lon

def create_random_voltage():
    """ Creates a random voltage between 230 and 250
    """
    return 230.0 + ( 20 * random.random() )

    
def create_random_thd():
    """ Creates a random voltage between 230 and 250
    """
    return ( 20 * random.random() )


def strTimeProp(start, end, format, prop):
    """Get a time at a proportion of a range of two formatted times.
    start and end should be strings specifying times formated in the
    given format (strftime-style), giving an interval [start, end].
    prop specifies how a proportion of the interval to be taken after
    start.  The returned time will be in the specified format.
    """

    stime = time.mktime(time.strptime(start, format))
    etime = time.mktime(time.strptime(end, format))

    ptime = stime + prop * (etime - stime)

    return time.strftime(format, time.localtime(ptime))
