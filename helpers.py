import xlrd
import csv
import sqlite3
import sys
import subprocess
import pathlib
import pandas as pd


def excel_convert(path):
    excel_path = path
    csv_path = path
    pathlib.Path(path + '/outputFiles/').mkdir(parents=True, exist_ok=True)

    with xlrd.open_workbook(excel_path) as wb:
        sh = wb.sheet_by_index(0)
    with open(csv_path, 'w', newline='') as f:
        c = csv.writer(f)
        for r in range(sh.nrows):
            c.writerow(sh.row_values(r))
    return csv_path


def create_database(path):
    # Use pandas to create a dataframe (with no index # column) from csv
    # str.strip() allows it to remove whitespace so it doesn't break
    db_name = path + '/outputFiles/database.db'
    sqlite3.connect(db_name)
    return db_name


def connect_database(mode, path, csv, database):
    df = pd.read_csv(csv, index_col=0)
    df.columns = df.columns.str.strip()

    # Connect to the database given in create_database()
    conn = sqlite3.connect(database)
    try:
        df.to_sql("ResultsTable", conn)
    except ValueError:
        df.to_sql("ResultsTable", conn, if_exists='replace')

    query_path = path + '/query_list.csv'
    output_path = path + '/outputFiles/results.csv'
    query_list = pd.read_csv(query_path, delimiter=',')

    for i in query_list:
        try:
            with open(output_path, 'a') as file:
                file.write(i)
                file.write('\n')
            table = pd.read_sql(i, conn)
            table.to_csv(output_path, mode='a', index=False, header=False)
        except (sqlite3.Error, pd.io.sql.DatabaseError):
            with open(output_path, 'a') as file:
                file.write('\n**Error: Query Failed**')
                file.write('\n\n')
            print('\nERROR: The query "' + i + '" failed!\n')
            pass

    # Closes the connection to the database
    conn.close()

    return output_path
