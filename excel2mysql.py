#!/usr/bin/env python

import os
import xlrd
import mysql.connector as mdb

def excel2mysql():
    try:
        # connect mysql
        conn = mdb.connect(user='root', password='', database='test', use_unicode=True)
        cursor = conn.cursor()
        cursor.execute('create table if not exists excel (filename varchar(20), rownum varchar(10), value varchar(5000))')
#       cursor.execute('insert into user ()')

        # read files
        folder_path = os.path.join(os.path.dirname(__file__), 'excel')
        files = filter(lambda x: x.split('.')[-1] in ['xls', 'xlsx'], os.listdir(folder_path))
        for filename in files:
            try:
                data = xlrd.open_workbook(os.path.join(folder_path, filename))
                print "open %s" % filename
            except StandardError, e:
                print "failed open %s" % filename
                print e
            table = data.sheet_by_index(0)
            for i in range(table.nrows):
                print table.row_values(i)
                cursor.execute('insert into excel (filename, rownum, value) values (%s, %s, %s)', [filename, str(i), ','.join(str(n) for n in table.row_values(i))])
            conn.commit()
        cursor.close()
        conn.close()
    except mdb.Error, e:
        conn.rollback()
        print "Error %d: %s" % (e.args[0], e.args[1])
        sys.exit(1)

if __name__ == '__main__':
    excel2mysql()
