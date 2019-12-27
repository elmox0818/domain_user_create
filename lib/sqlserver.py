import pyodbc
import pandas as pd

# DSN=db_koa;UID=koa;PWD=koamgr;


def ExecuteQueryBySQLServer(sql):
    con = pyodbc.connect(
        r'DRIVER={SQL Server};SERVER=violet;DATABASE=KOA;UID=koa;PWD=koamgr;')
    cur = con.cursor()
    cur.execute(sql)
    con.commit()
    con.close()


def ReadQueryBySQLServer(sql):
    con = pyodbc.connect(
        r'DRIVER={SQL Server};SERVER=violet;DATABASE=KOA;UID=koa;PWD=koamgr;')
    df = pd.io.sql.read_sql(sql, con)
    con.close()
    return(df)
