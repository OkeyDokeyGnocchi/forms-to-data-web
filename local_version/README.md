# Forms-to-Data

## What is it?
Forms-to-Data is a project designed around taking the results of an online form (e.g. Microsoft Forms) downloaded as an Excel .xlsx file, creating a database and table based on it, then running a query list through the database table.  The results of these queries are sent to a generated .csv file for the sake of analysis.  This project was a direct result of needing data on the students' perceived "campus climate" within a university department.

## Installation
Installation is intended to be as simple as possible.  At the time of this writing after downloading the forms-to-data.py file there are a few requirements to get the program to run:
- Python 3
  - pandas
  - xlrd

The long-term goal is to find a way to package everything needed into a single file so that it can be run without these types of requirements so that users can quickly and easily use the program to query results data with as little overhead as possible.

## Usage
Forms-to-Data is designed to be easy to use for users regardless of technical knowledge.  When run it operates in one of two primary modes: easy and advanced (selected by the user by typing either "1" or "2").
### Easy Mode

Easy mode is used by selecting option 1 at the beginning of the program.  Easy mode requires two files with specific names within the same directory (folder) that the program is run from: easy_input.xlsx and query_list.csv.  The files are as follows:
- easy_input.xlsx is your data spreadsheet (single sheet).  The top row of your file should give the desired column names for your database table (e.g. age, favoritePizza, favoriteDrink, etc.) with each row beneath it representing a single subset of that data (such as a single person's responses in a survey).  It is highly recommmended that you make these column headers either a single word or words connected with an underscore "\_" or a hyphen/dash "-" as this will make running SQL queries against your data much easier.
- query_list.csv is where you will write the SQL queries that you would like the program to run against the automatically-created database and table.  This file can be filled with a query in each column in the top row (NOTE: it will only run queries in the top row!).  The Forms-to-Data requires that this file be comma-delimited (this is the standard .csv format for most spreadsheet products).  More information on SQL queries for beginners can be found later in the README.</br>

If these files exist in the directory (folder) that the program is run in, the program will automatically create a folder named "outputFiles".  The program will then convert your easy_input.xlsx to a copy named input.csv, create the new database "database.db", create a table within database.db named "resultsTable", run your queries in query_list.csv through it, and create the file results.csv that contains your queries and the results row after row (query1, result1, query2, result2, etc.).

As a quick rundown the files involved are:
- _./easy_input.xlsx_ = your data in an xlsx file. The first row consists of your column headers (what data is in this column all the way down).
- _./query_list.csv_ = your SQL queries in the top row of a csv file.
- _./outputFiles/input.csv_ = your data in a csv file. Made by Forms-to-Data from your easy_input.xlsx file.
- _./outputFiles/database.db_ = the database file created by Forms-to-Data from the input.csv file.
- _./outputFiles/results.csv_ = the results from running your queries in query_list.csv through the created database.db. Contains your queries as they were run and the result of each query beneath it.

### Advanced Mode

Advanced mode is used by selecting option 2 at the beginning of the program. The primary change in Advanced mode is that the user selects the names/paths of their input and output files. Note: you do not have to type the various file extensions, if you omit them the program will fill in .xlsx/.csv/.db as appropriate.
Running through the program takes the following steps:
- Select 2 for advanced mode
- Select whether you will need to convert a .xlsx file to .csv or not (you can skip the conversion step by having your data in .csv form already!)
  - If you do need to convert from .xlsx to .csv either type the name of the .xlsx file (if it is in the same directory as the program) or the full path to the file you wish to use (e.g. "C:/Users/Me/Downloads/data.xlsx"). You will be prompted to give the name/path you would like to give your new .csv file.
- Give the desired name/path of the new database (.db) file. Forms-to-Data will print your selected name/location where it created the .db file back to you for your convenience.
- Give the name/path of your .csv file with your list of SQL queries that you would like to run through the database.
- Give the name/path for the output .csv file.
- Forms-to-Data will open the new .csv file with your query results (query1, result1, query2, result2, etc.) using your default program for .csv files.

