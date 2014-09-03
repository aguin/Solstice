#!/usr/bin/python3
import webserver
import exampledata

if __name__ == '__main__':
    exampledata.create_example_db()
    webserver.app.run(debug=True)
    
    
