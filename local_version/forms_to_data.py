# Turn Excel File to CSV, CSV to SQLITE3 DB+Table, Query new table from csv

import xlrd
import csv
import sqlite3
import os
import sys
import subprocess
import pathlib
import pandas as pd


def excel_convert(mode, path):
    if mode == 'adv':
        # Ask if the user needs to convert Excel to CSV
        print('Do you need to convert an xlsx file to csv? [y/n]')
        convert_needed = input()
        convert_needed = convert_needed.lower()

        # If yes, have user give the Excel and CSV files
        # Append the proper extensions if they aren't given
        if convert_needed == 'y' or convert_needed == 'yes':
            print('Enter Excel file input path:')
            excel_path = input()
            if '.xlsx' not in excel_path:
                excel_path = excel_path + '.xlsx'
            print('Enter CSV file output path:')
            csv_path = input()
            if '.csv' not in csv_path:
                csv_path = csv_path + '.csv'
        # If conversion isn't needed proceed with getting the csv path
        else:
            print('\nOK, moving forward without converting.\n')
            print('Enter CSV file output path:')
            csv_path = input()
            if '.csv' not in csv_path:
                csv_path = csv_path + '.csv'

    # Set easy mode filepaths for Excel conversion
    elif mode == 'easy':
        excel_path = path + '/easy_input.xlsx'
        csv_path = path + '/outputFiles/input.csv'
        # Create the new outputFiles directory in the working directory
        pathlib.Path(path + '/outputFiles/').mkdir(parents=True, exist_ok=True)

    with xlrd.open_workbook(excel_path) as wb:
        sh = wb.sheet_by_index(0)
    with open(csv_path, 'w', newline='') as f:
        c = csv.writer(f)
        for r in range(sh.nrows):
            c.writerow(sh.row_values(r))
    return csv_path


def create_database(mode, path):
    if mode == 'adv':
        # Ask user to give the name for the new DB, append .db if needed
        print('Please give the path of your new database:')
        print('If you just type the desired name I\'ll create it in the '
              'directory this is running from')
        db_name = input()
        if '.db' not in db_name:
            db_name = db_name + '.db'

    elif mode == 'easy':
        db_name = path + '/outputFiles/database.db'

    # Call SQLITE3 and connect to the user-given db_name. Confirm with print
    sqlite3.connect(db_name)
    print('I have created the database for you as ' + db_name + '\n')
    return db_name


def connect_database(mode, path, csv, database):
    # Use pandas to create a dataframe (with no index # column) from csv
    # str.strip() allows it to remove whitespace so it doesn't break
    df = pd.read_csv(csv, index_col=0)
    df.columns = df.columns.str.strip()

    # Connect to the database given in create_database()
    conn = sqlite3.connect(database)
    try:
        df.to_sql("ResultsTable", conn)
    except ValueError:
        print('Error: the table already exists from a previous run.')
        print('If you continue, the data in your database will be replaced.')
        print('\nDo you wish to delete the old data and add new data? [y/n]')
        table_replace = input()
        table_replace = table_replace.lower()
        if table_replace == 'y' or table_replace == 'yes':
            df.to_sql("ResultsTable", conn, if_exists='replace')
        else:
            print('The old data has not been deleted.')
            print('Please re-run the program in a different directory.')
            print('\nClosing the program.')
            sys.exit()

    if mode == 'adv':
        print('Please provide the name of your csv file with queries: ')
        query_path = input()
        if '.csv' not in query_path:
            query_path = query_path + '.csv'
        print('\nPlease type the name you would like for your output file: ')
        output_path = input()
        if '.csv' not in output_path:
            output_path = output_path + '.csv'
    elif mode == 'easy':
        query_path = path + '/query_list.csv'
        output_path = path + '/outputFiles/results.csv'

    # Use pandas to read csv with queries and write results to csv
    print('Retrieving queries from ' + query_path + '.\n')
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

    print('Writing query results to CSV file.')

    # Closes the connection to the database
    conn.close()

    return output_path


def open_results(mode, path, output_file):
    # Open results.csv
    if sys.platform == 'win32':
        os.startfile(output_file)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, output_file])


def run_queries(mode, path, query_path, db_name, output_path):
    # Connect to the database given in create_database()
    conn = sqlite3.connect(db_name)
    # Use pandas to read csv with queries and write results to csv
    print('Retrieving queries.\n')
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

    print('Writing query results to CSV file.')

    # Closes the connection to the database
    conn.close()


def query_mode_paths(mode, path):
    if mode == 'easy':
        db_name = path + '/outputFiles/database.db'
        query_path = path + 'query_list.csv'
        output_path = path + '/outputFiles/results.csv'
    else:
        print('Please give the full path of your database:')
        db_name = input()
        if '.db' not in db_name:
            db_name = db_name + '.db'
        print('Please give the full path to your csv file with queries: ')
        query_path = input()
        if '.csv' not in query_path:
            query_path = query_path + '.csv'
        print('Please give the full path for your output file: ')
        output_path = input()
        if '.csv' not in output_path:
            output_path = output_path + '.csv'

    return db_name, query_path, output_path


def forms_to_data(mode, path):
    csv_path = excel_convert(mode, path)
    db_name = create_database(mode, path)
    output_path = connect_database(mode, path, csv_path, db_name)
    open_results(mode, path, output_path)


def query_mode():
    print('\nWelcome to query mode.\n')
    print('Did you create your database with '
          '(1)easy or (2)advanced mode? [1/2]')
    mode_input = input()
    if mode_input == '1':
        mode = 'easy'
    elif mode_input == '2':
        mode = 'adv'

    db_name, query_path, output_path = query_mode_paths(mode)
    run_queries(mode, query_path, db_name, output_path)
    open_results(mode, output_path)


def main():
    while True:
        path = os.path.dirname(sys.argv[0])
        print('\nWelcome to Forms-to-Data!\n')
        print('Would you like (1)easy, (2)advanced, or '
              '(3)query mode? [1/2/3]')
        user_mode = input()
        if user_mode == '1':
            forms_to_data('easy', path)
            break
        elif user_mode == '2':
            forms_to_data('adv', path)
            break
        elif user_mode == '3':
            query_mode()
            break
        else:
            print('Please select either (1)easy, (2)advanced, or '
                  '(3)query mode.\n')


if __name__ == '__main__':
    main()
