# encoding=utf-8

import logging
import time
import TsListSpider


logging.basicConfig()

if __name__ == "__main__":

    app = TsListSpider.TsListSpider()
    # app.TestGetProxy()
    #
    # app.TestConn()
    # crawler.TestExist()
    # app.TestQueryAll()
    # app.RunGetTsIndexFromSql('2015-9-1','2015-10-9')
    # app.RunGetTsIndexFromSql()
    # app.TestHttp()
    # time.sleep(24*60*60)
    app.RunSpider()

    while 1:
        app.Running()
        time.sleep(1)
