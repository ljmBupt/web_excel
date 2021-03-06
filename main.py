#!/usr/bin/env python

import tornado.httpserver
import tornado.ioloop
import tornado.options
import tornado.web

import os
# import xlrd
import mysql.connector as mdb

from tornado.options import define, options

# define Tornado defaults
define("port", default=8000, help="run on the given port", type=int)

# application configuration
class Application(tornado.web.Application):
    def __init__(self):
        handlers = [
            (r"/", FileHandler),
            (r"/show", ExcelHandler),
        ]
        settings = dict(
            template_path = os.path.join(os.path.dirname(__file__), "templates"),
            static_path = os.path.join(os.path.dirname(__file__), "static"),
            debug = True,
        )
        tornado.web.Application.__init__(self, handlers, **settings)

class FileHandler(tornado.web.RequestHandler):
    def gen_filelist(self):
        excel_path = os.path.join(os.path.dirname(__file__), "excel")
        excel_files = filter(lambda x: x.split('.')[-1] in ['xls', 'xlsx'], os.listdir(excel_path))
        return excel_files

    def get(self):
        files = self.gen_filelist()
        self.render(
            "files.html",
            files=files
        )

class ExcelHandler(tornado.web.RequestHandler):
    def read_excel(self, filename='html.xls'):
        excel_path = os.path.join(os.path.dirname(__file__), "excel")
        file_path = os.path.join(excel_path, filename)
        try:
            conn = mdb.connect(user='root', password='', database='test', use_unicode=True)
            cursor = conn.cursor()
            cursor.execute('select count(*) from excel where filename = %s', [filename])
            values = cursor.fetchall()
            nrows = values[0][0]
            text = []
            for i in range(nrows):
                cursor.execute('select * from excel where filename = %s and rownum = %s', [filename, str(i)])
                values = cursor.fetchall()
                text.append(values[0][2].split(','))
            cursor.close()
            conn.close()
        except StandardError, e:
            print e
        return text

    def post(self):
        filename = self.get_argument('filename')
        text= self.read_excel(filename)
        self.render(
            "excel.html",
            filename=filename,
            text=text
        )


# Start it up
def main():
    tornado.options.parse_command_line()
    http_server = tornado.httpserver.HTTPServer(Application())
    http_server.listen(options.port)
    tornado.ioloop.IOLoop.instance().start()

if __name__ == "__main__":
    main()
