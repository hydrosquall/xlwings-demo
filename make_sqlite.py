# -*- coding: utf-8 -*-
# @Author: Cameron Yick
# @Date:   2017-08-03 10:23:43
# make_sqlite.py: Create a demo sqlite database for usage with
# requests-cache
import sqlite3
from sqlite3 import Error
import os


def create_connection(db_file):
    """ create a database connection to a SQLite database """
    try:
        conn = sqlite3.connect(db_file)
        print(sqlite3.version)
    except Error as e:
        print(e)
    finally:
        conn.close()


if __name__ == '__main__':
    outfile = os.path.join('hellopandas', 'example.db')
    create_connection(outfile)
