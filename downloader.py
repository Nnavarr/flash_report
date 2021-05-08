import pandas as pd
from getpass import getuser, getpass
import pyodbc
import urllib
import sqlalchemy
import datetime as dt
import os

# Establish columns in SQL tables, SQL table names, and report file names
COLNAMES = ["AccountingWeek", "Category", "start_date", "end_date", "amount"]
TABLENAMES = ["FlashReport", "FlashReportUHI"]
FILENAMES = ["Flashrpt.csv", "Flashuhi.csv"]

def build_sql_engine(database, user=None, pw=None):
    """str -> SQLAlchemy engine
    Returns a sqlalchemy engine for the provided database using the
    provided credentials. """
    if not user:
        user = getuser()
    if not pw:
        pw = os.environ.get("sql_pwd")
    base_con = (
        "Driver={{SQL Server}};"
        "Server=<server name>;"
        "Database={};"
        "UID={};"
        "PWD={};"
    ).format(database, user, pw)

    # SQLAlchemy extension of standard connection
    params = urllib.parse.quote_plus(base_con)
    engine_str = "mssql+pyodbc:///?odbc_connect=%s" % params
    return sqlalchemy.create_engine(engine_str)

def main():
    # engine = build_sql_engine("PNL_Reports")
    if len(getuser()) < 7:
        user = '1217543'
    else:
        user = getuser()

    pw = os.getenv('sql_pwd')
    engine = build_sql_engine("PNL_Reports", user, pw)

    # Build the sql query, get the data, and write to .csv on the desktop
    for table, file in zip(TABLENAMES, FILENAMES):
        query = "SELECT {0} FROM {1}".format(", ".join(COLNAMES), table)
        data = pd.read_sql(query, engine)
        save_loc = "C:/Users/{0}/Desktop/{1}".format(getuser(), file)
        data.to_csv(save_loc, index=False)



if __name__ == '__main__':
    main()
