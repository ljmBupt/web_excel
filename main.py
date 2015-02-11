#!/usr/bin/env python

import os
import tornado.httpserver
import tornado.ioloop
import tornado.options
import tornado.web

from tornado.options import define, options

# define Tornado defaults
define("port", default=8000, help="run on the given port", type=int)

# application configuration
class Application(tornado.web.Application):
    def __init__(self):
        handlers = [
            (r"/", SimpleHandler),
        ]
        settings = dict(
            template_path = os.path.join(os.path.dirname(__file__), "templates"),
            static_path = os.path.join(os.path.dirname(__file__), "static"),
            debug = True,
        )
        tornado.web.Application.__init__(self, handlers, **settings)

class SimpleHandler(tornado.web.RequestHandler):
    def read_excel(self, filename="html.xlsx"):
        excel_path = os.path.join(os.path.dirname(__file__), "html.xlsx")
        data = xlrd.open_workbook(excel_path)

    def get(self):
        self.render(
            "excel.html",
            title="Excel Page"
            header="Excel"
            intro=""
        )


# Start it up
def main():
    tornado.options.parse_command_line()
    http_server = tornado.httpserver.HTTPServer(Application())
    http_server.listen(options.port)
    tornado.ioloop.IOLoop.instance().start()

if __name__ == "__main__"
    main()
