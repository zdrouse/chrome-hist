# Python Project 3
# Zack Rouse
# CSI 5740
# A python script to obtain data from Chrome web history.

# native
import argparse
import os
import sqlite3
import datetime

# 3rd party
import xlsxwriter

def db_connect(db):
    # returns a sql lite database cursor to a database file
    try:
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
    return result[0]

def get_visits(c, limit):
    # returns a list of objects for the most recent 25 urls within Chrome History visits table.
    visits = []
    query = c.execute(f'SELECT url, visit_time, segment_id FROM visits ORDER BY visit_time DESC LIMIT {limit}')
    results = query.fetchall()
    for i, row in enumerate(results):
        visit = {}
        url = get_url(c, row[0])
        segment_url = get_segment_url(c, row[2])
        visit = {
            "url": url,
            "visit_time": row[1],
            "segment_url": row[2]
        }
        visits.append(visit)
    return(visits)

def main():
    # arg parser
    parser = argparse.ArgumentParser()
    # make argument flags
    parser.add_argument('-i', '--src', help='Source db file to scan', required=True, type=str)
    # parser.add_argument('-o', '--dest', help='Output folder', required=True, type=str)
    # parse the arguments
    args = parser.parse_args()

    db_cursor = db_connect(args.src)
    visits = get_visits(db_cursor, 25)

if __name__ == "__main__":
    main()