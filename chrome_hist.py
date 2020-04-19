# Python Project 3
# Zack Rouse
# CSI 5740
# A python script to obtain data from Chrome web history.

# native
import argparse
import os
import sqlite3
import datetime
import time

# 3rd party
import xlsxwriter

def db_connect(db):
    # returns a sql lite database cursor to a database file
    try:
        print("Connecting to database file...\n")
        conn = sqlite3.connect(db)
        c = conn.cursor()
        return c
    except Exception as e:
        print(f"Failed to connect the database for: {db}.  Exception: {e}")
        exit(1)

def get_url(c, url_id):
    # returns a url for a url_id from the Chrome History visits table.
    query = c.execute(f'SELECT url FROM urls WHERE id == {url_id}')
    result = query.fetchone()
    return result[0]

def get_segment_url(c, segment_id):
    # returns a segment url for a segment_id from the Chrome History visits table.
    query = c.execute(f'SELECT name FROM segments WHERE id == {segment_id}')
    result = query.fetchone()
    if result == None:
        return "N/A"
    else:
        return result[0]

def get_visits(c, limit):
    try:
        # returns a list of objects for the most recent 25 urls within Chrome History visits table.
        print(f"Pulling the last {limit} visits from database...\n")
        visits = []
        query = c.execute(f"SELECT url, datetime((visit_time/1000000)-11644473600, 'unixepoch', 'localtime') AS time, segment_id FROM visits ORDER BY visit_time DESC LIMIT {limit}")
        results = query.fetchall()
        print("Displaying visit list of visit objects: \n")
        for i, row in enumerate(results):
            visit = {}
            url = get_url(c, row[0])
            segment_url = get_segment_url(c, row[2])
            visit_time = row[1]
            visit = {
                "url": url,
                "visit_time": visit_time,
                "segment_url": segment_url
            }
            print(f"  {i+1}) {visit}")
            visits.append(visit)
        #print(visits)
        return(visits)
    except Exception as e:
        print(f"There was an error retrieving the visits from the database file: {e}")
        exit(2)

def export_summary(visits_list, out_path):
    try:
        # exports an excel file of the visits list passed. if the directory does not exist, it will create it.
        # we use expanduser on os path so that the script can respect the user environment home directory
        print("\nExporting the visits data to an excel summary file...")
        if not os.path.exists(os.path.expanduser(out_path)):
            os.makedirs(os.path.expanduser(out_path))
        workbook = xlsxwriter.Workbook(f'{os.path.expanduser(out_path)}/summary.xlsx')
        worksheet = workbook.add_worksheet()
        if len(visits_list) != 0:
            row = 1
            col = 0
            worksheet.write(0, 0, 'URL')
            worksheet.write(0, 1, 'Date and Time of Visit')
            worksheet.write(0, 2, 'Segment URL')
            for visit in visits_list:
                worksheet.write(row, col, visit['url'])
                worksheet.write(row, col + 1, visit['visit_time'])
                worksheet.write(row, col + 2, visit['segment_url'])
                row += 1
        workbook.close()
    except Exception as e:
        print(f"Error exporting excel summary file: {e}")
        exit(1)

def main():
    # arg parser
    parser = argparse.ArgumentParser()
    # make argument flags
    parser.add_argument('-i', '--src', help='Source db file to scan', required=True, type=str)
    parser.add_argument('-o', '--dest', help='Output folder', required=True, type=str)
    # parse the arguments
    args = parser.parse_args()
    # make db cursor from connection
    db_cursor = db_connect(args.src)
    # grab custom visits list of visit objects from database
    visits = get_visits(db_cursor, 25)
    # export the visit results to an excel file located at the output folder
    export_summary(visits, args.dest)
    print("Done.")
    exit(0)

if __name__ == "__main__":
    main()