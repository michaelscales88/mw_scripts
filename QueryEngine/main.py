# sys.path.insert(0, 'python{0}/'.format(sys.version_info[0]))
import pypyodbc
import datetime
import sys
import psycopg2
# sys.path.append("C:\\Users\\mscales\\AppData\\Local\\Temp\\pip-build-6tkd7eee\\psycopg2")
# from psycopg2 import psycopg1


today = datetime.datetime.today().strftime('%d%m%y')

psycopg2conn_string = ("dbname='chronicall' user='Chronicall' host='10.1.3.17' password='ChR0n1c@ll1337'")
connection_string = "Driver={SQL Server};Server=localhost;Database=mw_calling;Trusted_Connection=yes"
connection_string2 = ("DRIVER={PostgreSQL Unicode};"
                      "Server=10.1.3.17;"
                      "PORT=9086;"
                      "Database=chronicall;"
                      "UID=Chronicall;"
                      "PWD=ChR0n1c@ll1337;")
connection_string3 = ("Driver={PostgreSQL ANSI};"
                      "Server=10.1.3.17;"
                      "Port=5432;"
                      "Database=chronicall;"
                      "Uid=Chronicall;"
                      "Pwd=ChR0n1c@ll1337;"
                      "sslmode=require;")
SQLCommand = ("INSERT INTO testTable(clientname, stuff) VALUES (7517, 'stuff_stuff')")

try:
    # cnx = pypyodbc.connect(connection_string3)
    conn = psycopg2.connect(psycopg2conn_string)
    cur = conn.cursor()
    # cur.execute(SQLCommand)
    # cur.commit()
    cur.close()
except Exception as e:
    print("failed connection")
    print(e)
else:
    conn.close()
