import pymysql
import pandas as pd
import datetime
from sqlalchemy import create_engine
from MySQL_Authorization import MySQL_Auth

def pull_hy_match_set():

    access_token = MySQL_Auth()
    conn = pymysql.connect(host='localhost', port=3306, user='tlack', passwd=access_token, db='bens_desk')
    engine = create_engine('mysql+pymysql://tlack:%s@localhost/bens_desk' % (access_token))

    db_max_date = pd.DataFrame(pd.read_sql("SELECT As_Of_Date FROM bens_desk.hyhg_index \
                                WHERE As_Of_Date IN (SELECT max(As_Of_Date) FROM bens_desk.hyhg_index)", conn))
    db_max_date = str(db_max_date.iloc[0,0])
    db_max_date = db_max_date[0:10]
    db_max_date = db_max_date[0:4] + db_max_date[5:7] + db_max_date[8:10]
    return db_max_date