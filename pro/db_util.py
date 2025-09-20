import pymysql

conn = pymysql.connect(
    host="localhost",
    user="root",
    password="12345",
    database="user",
    port=3306,
    charset="utf8mb4"
)


def get_conn():
    return conn
