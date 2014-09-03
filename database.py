import json
import os
import re
import io
import csv
import sys

# set path names and db connection settings
SQL_PATH = "sql"

def construct_dict(cursor):
    """ transforms the db cursor rows from table format to a
        list of dictionary objects
    """
    rows = cursor.fetchall()
    return [dict((cursor.description[i][0], value) for i, value in enumerate(row))
            for row in rows]


def construct_list(cursor):
    """ transforms the db cursor into a list of records,
        where the first item is the header
    """
    header = [h[0] for h in cursor.description]
    data = cursor.fetchall()
    return header, data


def construct_csv(cursor):
    """ transforms the db cursor rows into a csv file string
    """
    
    header, data = construct_list(cursor)
    # python 2 and 3 handle writing files differently
    if sys.version_info[0] <= 2:
        output = io.BytesIO()
    else:
        output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow(header)
    for row in data:
        writer.writerow(row)

    return output.getvalue()


def load_query(query_filename):
    with open(query_filename) as f:
        return f.read()


def get_driver(driver_name):
    if driver_name == 'sqlite3':
        import sqlite3 as db_driver
    elif driver_name == 'cx_Oracle':
        import cx_Oracle as db_driver
    elif driver_name == 'pyodbc':
        import pyodbc as db_driver
    elif driver_name == 'psycopg2':
        import psycopg2 as db_driver
    elif driver_name == 'PyMySql':
        import PyMySql as db_driver
    else:
        # TODO: pick a better exception type and message
        raise ImportError
    return db_driver


def run_query(settings_filename, query_filename, params_dict, data_format='list'):
    # set up a db connection from the settings    
    with open(settings_filename) as f:
        SETTINGS = json.load(f)
    db_driver = get_driver(SETTINGS['db_driver'])
    conn = db_driver.connect(SETTINGS['db_connection_string'])
    cursor = conn.cursor()

    # run the query
    cursor.execute(load_query(query_filename), params_dict)
    if data_format == 'list':
        # format into a table with header
        query_results = construct_list(cursor)
    elif data_format == 'dict':
        # format into dictionary
        query_results = construct_dict(cursor)
    elif data_format == 'csv':
        # format into dictionary
        query_results = construct_csv(cursor)
    conn.close()

    return query_results


