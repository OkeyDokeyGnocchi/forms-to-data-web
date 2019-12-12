import xlrd
import csv
import sqlite3
import pandas as pd
import string
import random


def connect_database(user_query, converted_csv, database, user_filename):
    df = pd.read_csv(converted_csv, index_col=0)
    df.columns = df.columns.str.strip()

    # Connect to the database given in create_database()
    conn = sqlite3.connect(database)
    try:
        df.to_sql("ResultsTable", conn)
    except ValueError:
        df.to_sql("ResultsTable", conn, if_exists='replace')

    output_file = 'results_' + user_filename + '.csv'
    query_list = pd.read_csv(user_query, delimiter=',')

    for i in query_list:
        try:
            with open(output_file, 'a') as file:
                file.write(i)
                file.write('\n')
            table = pd.read_sql(i, conn)
            table.to_csv(output_file, mode='a', index=False, header=False)
        except (sqlite3.Error, pd.io.sql.DatabaseError):
            with open(output_file, 'a') as file:
                file.write('\n**Error: Query Failed**')
                file.write('\n\n')
            print('\nERROR: The query "' + i + '" failed!\n')
            pass

    # Closes the connection to the database
    conn.close()

    return output_file


def create_database(user_filename):
    # Use pandas to create a dataframe (with no index # column) from csv
    # str.strip() allows it to remove whitespace so it doesn't break
    db_name = user_filename + '_database.db'
    sqlite3.connect(db_name)
    return db_name


def excel_convert(xlsx_file, user_filename):
    excel_path = xlsx_file
    csv_path = user_filename + '.csv'
    # pathlib.Path(path + '/outputFiles/').mkdir(parents=True, exist_ok=True)

    with xlrd.open_workbook(excel_path) as wb:
        sh = wb.sheet_by_index(0)
    with open(csv_path, 'w', newline='') as f:
        c = csv.writer(f)
        for r in range(sh.nrows):
            c.writerow(sh.row_values(r))
    return csv_path


def generate_filename(size=6, chars=string.ascii_uppercase + string.digits):
    return ''.join(random.choice(chars) for _ in range(size))
