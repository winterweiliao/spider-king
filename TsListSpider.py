# encoding=utf-8

import os
import re
import shutil
import time
from datetime import *
import requests
import win32com.client
import pymssql


reload(sys)
#sys.setdefaultencoding('gbk')
sys.setdefaultencoding('utf8')


class TsListSpider:
    """
    参考网站：http://www.runoob.com/
    """
    def __init__(self):
        """
        数据初始化
        """
        print u'TS自动化实时采集v2015.09\n'
        print u'已启动...\n'
        
        self.session = requests.Session()
        self.proxies = {}
        #self.proxies = {"http": "http://203.195.160.14:80"}
        #self.proxies = {"http": "http://58.214.5.229:80"}
        #self.proxies = {"http": "http://223.19.196.232:80"}
        self.headers = {}

        self.host = "192.168.0.69"
        self.user = "sa"
        self.pwd = "sa_123456"
        self.db = "vip_test"
        
        self.taskfile = ".\\__TSDXCode.mdb"
        self.logging = 1
        self.__baselist = self.__QueryTaskList2()
        self.__errlist = list()
        self.__updatecount = 0
        self.__proxylist = set()
        self.__reuselist = set()
        self.__proxy = "203.195.160.14:80"
        start = str(int(time.time()))
        self.__outputdir = ".\\output\\"+start+"\\"        
        self.__template = ".\\__TaskList_T.mdb"

    def Running(self):
        today = date.today().year
        self.Run(2014,today)

    def Run(self,fromyear=2015,toyear=None):
        """
        入口
        """
        if(toyear == None):
            toyear = fromyear
            
        print u'起始年:',fromyear
        print u'结束年:',toyear
        tasklist = self.__GetTaskList(fromyear,toyear)        
        self.__errlist = list()
        self.__WalkJobList(tasklist)

        while len(self.__errlist)>0:
            tasklist = self.__errlist
            self.__errlist = list()            
            self.__WalkJobList(tasklist)

            
        print(u'\n\n已完成。\n')

    def RunGetTsIndexFromSql(self,fromtime=None,totime=None):
        """
        入口
        """
            
        print u'起始时间:',fromtime
        print u'结束时间:',totime
        outputmdb = self.__outputdir+"tasklist.mdb"
        os.makedirs(self.__outputdir)
        shutil.copyfile(self.__template,outputmdb)
        
        while 1:
            resList = self.__QueryTsIndexList(fromtime,totime)
            print u'saving...'
            for item in resList:
                self.__SaveTsIndexToMdb(outputmdb,item)
            break


        print(u'\n\n已完成。\n')

    def RunSpider(self,fromtime=None,totime=None):
        """
        入口
        """
            
        print u'起始时间:',fromtime
        print u'结束时间:',totime
        outputmdb = self.__outputdir+"tasklist.mdb"
        os.makedirs(self.__outputdir)
        shutil.copyfile(self.__template,outputmdb)
        
        
        
        while 1:
            for item in self.__baselist:
                print item.TaskUrl
                self.__DoJob2(item)
            break

        while len(self.__errlist)>0:
            tasklist = self.__errlist
            self.__errlist = list()            
            for item in tasklist:
                #print item.TaskUrl
                self.__DoJob2(item)


        print(u'\n\n已完成。\n')

    def __GetTaskList(self,fromyear,toyear):
        reslist = list()

        while fromyear<toyear or fromyear==toyear:
            templist = self.__InitTaskList(self.__baselist,fromyear)
            reslist.extend(templist)
            fromyear = fromyear+1

        return reslist

    def __InitTaskList(self,list0,year):
        resList = list()
        for item in list0:
            task = Task(item.sw,year)
            resList.append(task)
        
        return resList

    def __WalkJobList(self,list0):
        for item in list0:
            self.__DoJob(item)

    def __DoJob(self,task):        
        try:
            i=0
            reslist = list()
            
            while i<200:                
                if i%20==0:
                    self.__Login()
                print 'Keyword:',task.sw
                text = self.__HttpGet(task.TaskUrl)
                if text.decode('utf8').find(u'提示页面')>-1:
                    raise(NameError, u'验证码')
                    
                templist = self.__ParseTsList(task,text)
                if len(templist)==0:
                    break
                self.__SaveList(templist)
                self.__reuselist.add(self.__proxy)
                print u'丢失计数:',len(self.__errlist),u'\t更新计数:',self.__updatecount
                task.Next()
                i = i+1
            
        except Exception,ex:
            self.__errlist.append(task)
            print Exception,":",ex
            print u'丢失计数:',len(self.__errlist),u'\t更新计数:',self.__updatecount

    def __DoJob2(self,task):        
        try:
            i=0
            reslist = list()
            
            while 1:                
                print 'i:',str(i)
                print 'page:',str(task.page)
                text = self.__HttpGet(task.TaskUrl)
                if text.decode('utf8').find(u'提示页面')>-1:
                    raise(NameError, u'验证码')
                    
                templist = self.__ParseBaiduList(task,text)
                self.__SaveList2(templist)
                #print u'丢失计数:',len(self.__errlist),u'\t更新计数:',self.__updatecount
                task.Next()
                i = i+1
            
        except Exception,ex:
            self.__errlist.append(task)
            print Exception,":",ex
            #print u'丢失计数:',len(self.__errlist),u'\t更新计数:',self.__updatecount
        
    def __Login(self):
        if len(self.__proxylist)==0:
            print 'getting proxy...'
            self.__proxylist = self.__GetProxyFromDaili666(10)
            for item in self.__reuselist:
                self.__proxylist.add(item)
            print 'proxy count:',len(self.__proxylist)
            if len(self.__proxylist)==0:
                return
                
        if self.logging:
            print u'login...'
        url = "http://www.duxiu.com/loginhl.jsp?send=true&UserName=shr&PassWord=shr"
        self.__proxy=self.__proxylist.pop()
        self.proxies={'http': 'http://' + self.__proxy}
        print self.proxies

        try:
            content = self.__HttpPost(url,"")
            #print content.decode("utf8")
        except Exception,ex:
            self.__Login()
            
    def __HttpGet(self,url):
        if self.logging:
            i = url.find('/search')
            msg = url[i:]
            print 'download...'
            #print 'url:',url
        sReturn = ""
        self.session = requests.Session()
        r = self.session.get(url=url,headers=self.headers,proxies=self.proxies,timeout=20)
        data = r.content


        sReturn = data.decode('utf-8')

        i = data.find('</span>')
        msg = data[i:3]
        #print 'sssss:',i


        #print sReturn.encode('gb18030')

        return sReturn

    def __HttpPost(self,url,data):
        sReturn = ""
        r = self.session.post(url=url,data=data,headers=self.headers,proxies=self.proxies,timeout=10)
        sReturn = r.content

        return sReturn

    def __GetProxyFromDaili666(self,num):
        headers = {
            'Accept': '*/*',
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET4.0C; .NET4.0E)',
        }
        url = 'http://tiqu.daili666.com/ip/?'
        url += 'tid=' + '556923006054759'
        url += '&num=' + str(num)
        #url += '&filter=on'
        url += '&foreign=all'
        #url += '&delay=5'       #延迟时间
        
        #print('daili666 url:' + url)
        r = None	#声明
        exMsg = None
        try:
            #print('before url:' + url)
            r = requests.get(url=url, headers=headers, timeout=50)	
        except Exception,ex:
            exMsg = repr(sys.exc_info()[0]) + ';' + repr(sys.exc_info()[1])
            print('exMsg:' + exMsg)
        
        if exMsg or r.status_code != 200:
            return set()
            
        ProxyPoolTotal = set()
            
        #print('daili666:' + repr(r.content)) 
        lst = r.content.split('\n')
        for line in lst:
            line = line.strip()
            if len(line)>10:# and self.__ValidProxy(line):
                print line
                ProxyPoolTotal.add(line)
        
        return ProxyPoolTotal

    def __ValidProxy(self,proxy,url="http://www.duxiu.com",feature=u'系统登录'):
        bReturn = 0
        
        proxies = {'http': 'http://' + proxy}
        try:
            r = requests.get(url=url, proxies=proxies, timeout=10)	
            if r.status_code == 200 and r.content.decode("utf8").find(feature):
                bReturn = 1
                
        except Exception,ex:
            exMsg = repr(sys.exc_info()[0]) + ';' + repr(sys.exc_info()[1])
            print('exMsg:' + exMsg)

        return bReturn
    
    def __ParseBaiduList(self,task,txtHtml):
        """
        解析图书列表信息      
        """
        if self.logging:
            print(u'matching...')
        retlist = list()
        p = re.compile(r'<span class=\"g\">(?P<site>.+?)/&nbsp;</span>')
        for m in p.finditer(txtHtml):
            item = Content()
            retlist.append(item)
            item.site = m.group('site').replace('<b>','').replace('</b>','')

        return retlist

    def __ParseTsList(self,task,txtHtml):
        
        """
        解析图书列表信息      
        """
        if self.logging:
            print(u'matching...')
        retlist = list()
        p = re.compile(r'http://book.duxiu.com/bookDetail.jsp\?dxNumber=(?P<dxNumber>\d+)&d=(?P<d>\w+)&fenlei=(?P<fenlei>\d+)&sw=')
        for m in p.finditer(txtHtml):
            item = TsIndex()
            retlist.append(item)
            item.Keys = m.group('dxNumber')
            item.FromYear = task.year
            item.FullTextAddr = m.group(0)

        return retlist

    def __SaveList(self,list0):
        print u'saving...'
        for item in list0:
            self.__SaveToSqlServer(item)

    def __SaveList2(self,list0):
        #print u'saving...'
        outputmdb = self.__outputdir+"tasklist.mdb"
        for item in list0:
            self.__SaveContentToMdb(outputmdb,item)

    def __SaveToSqlServer(self,item):
        item.UpdateTime = date.today()
        if self.logging:
            print item.Keys,item.FromYear,item.UpdateTime
        if self.__Exist(item):
            print u'已存在。'
            return
            
        sql = ""
        sql += "insert into tasklist([Keys],[FromYear],[UpdateTime],[FullTextAddr])"
        sql += "values("
        sql += "'"+item.Keys+ "','"+str(item.FromYear)+ "','"+str(item.UpdateTime)+ "','"+item.FullTextAddr+ "'"
        sql += ")"
        #print sql

        try:
            self.__ExecNonQuery(sql)
            self.__updatecount += 1
            if self.logging:
                print(u'写入成功。')
        except:
            print(u'写入失败。')

    def __SaveTsIndexToMdb(self,mdb,item):
        if self.logging:
            print item.Keys,item.FromYear,item.UpdateTime
            
        sql = ""
        sql += "insert into tasklist([Keys],[FromYear],[UpdateTime],[FullTextAddr])"
        sql += "values("
        sql += "'"+item.Keys+ "','"+str(item.FromYear)+ "','"+str(item.UpdateTime)+ "','"+item.FullTextAddr+ "'"
        sql += ")"
        #print sql

        try:
            conn = win32com.client.Dispatch(r'ADODB.Connection')
            DSN = 'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE='+mdb+';'
            conn.Open(DSN)
            conn.Execute(sql)
            conn.Close()
            #if self.logging:
                #print(u'写入成功。')
        except:
            print(u'写入失败。')

    def __SaveContentToMdb(self,mdb,item):
        if self.logging:
            print item.site

        if self.__ExistInMdb(mdb,item):
            print u'已存在。'
            return
            
        sql = ""
        sql += "insert into content([Site])"
        sql += "values("
        sql += "'"+item.site+"'"
        sql += ")"
        #print sql

        try:
            conn = win32com.client.Dispatch(r'ADODB.Connection')
            DSN = 'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE='+mdb+';'
            conn.Open(DSN)
            conn.Execute(sql)
            conn.Close()
            #if self.logging:
                #print(u'写入成功。')
        except:
            print(u'写入失败。')

    def __ExistInMdb(self,mdb,item):
        bReturn = 0
        sql = ""
        sql += "SELECT [Site] FROM content where 1=1 "
        sql += " and [Site]='"+item.site + "'"

        conn = win32com.client.Dispatch(r'ADODB.Connection')
        DSN = 'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE='+mdb+';'
        conn.Open(DSN)
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open(sql,conn,1,3)
        #print rs.RecordCount
        bReturn = rs.RecordCount


        return bReturn
    
    def __Exist(self,item):
        bReturn = 0
        sql = ""
        sql += "SELECT [Keys] FROM tasklist where 1=1 "
        sql += " and [Keys]='"+item.Keys + "'"
        
        reslist = self.__ExecQuery(sql)
        bReturn = len(reslist)

        return bReturn

    def __QueryTsIndexList(self,start=None,end=None):
        if self.logging:
            print(u'loading basedata...')
        reslist = list()
        sql = ""
        sql += "SELECT [Keys],[FromYear],[UpdateTime],[FullTextAddr] FROM tasklist where 1=1 "
        if start!=None:
            sql += " and [UpdateTime]>='"+ start + "'"
        if end!=None:
            sql += " and [UpdateTime]<='"+ end + "'"
        sql += " order by UpdateTime "
        
        queryList = self.__ExecQuery(sql)
        for (Keys,FromYear,UpdateTime,FullTextAddr) in queryList:
            item = TsIndex(Keys,FromYear,FullTextAddr)
            item.UpdateTime = UpdateTime
            reslist.append(item)

        return reslist

    def __QueryTaskList(self):
        """
        查询任务列表      
        """
        if self.logging:
            print(u'loading basedata...')
        resList = list()
        conn = win32com.client.Dispatch(r'ADODB.Connection')
        DSN = 'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE='+self.taskfile+';'
        conn.Open(DSN)
        rs = win32com.client.Dispatch(r'ADODB.Recordset')
        rs.Open('select Keywords from Search order by Keywords',conn,1,3)
        #print rs.RecordCount
        rs.MoveFirst()
        while not rs.EOF:
            Keywords = rs.Fields.Item('Keywords').Value
            task=Task(Keywords)
            resList.append(task)
            rs.MoveNext()
        conn.Close()
        return resList

    def __QueryTaskList2(self):
        """
        查询任务列表      
        """
        if self.logging:
            print(u'loading basedata...')
        resList = list()

        task=Task('','')
        resList.append(task)

        return resList
        
    def __GetSqlConnection(self):
        """
        获得数据库连接信息      
        """
        self.conn = pymssql.connect(host=self.host,user=self.user,password=self.pwd,database=self.db,charset="utf8")
        cur = self.conn.cursor()
        if not cur:
            raise(NameError, u'数据库连接失败。')
        else:
            return cur

    def __ExecQuery(self,sql):
        """
        执行查询语句
        返回的是一个包含tuple的list，list的元素是记录行，tuple的元素是每行记录的字段        
        """
        cur = self.__GetSqlConnection()
        cur.execute(sql)
        resList = cur.fetchall()

        #查询完毕后必须关闭连接
        self.conn.close()
        return resList

    def __ExecNonQuery(self,sql):
        """
        执行非查询语句
        """
        cur = self.__GetSqlConnection()
        cur.execute(sql)
        self.conn.commit()
        self.conn.close()

    
#************ 测试 ************
    def TestConn(self):
        if self.__GetSqlConnection():
            print u'连接成功。'
        else:
            print u'连接失败。'
        
    def TestQueryAll(self):
        resList = self.__ExecQuery("SELECT count([Keys]) FROM tasklist")
        total = resList
                   
        resList = self.__ExecQuery("SELECT [Keys],[FromYear],[UpdateTime] FROM tasklist order by UpdateTime")
        i=1
        for (bid,num,years) in resList:
            print bid,num,years,total,i
            i = i+1
            #break

    def TestClear(self):
        try:
            self.__ExecNonQuery("delete FROM tasklist")
            print u'清理成功。'
        except:
            print u'清理失败。'

    def TestGetProxy(self):
        self.__GetProxyFromDaili666(100)

    def TestQuery(self,year=None,bid=None):
        sql = ""
        sql += "SELECT [bid],[num],[years] FROM modify_bid_num_info where 1=1 "
        if year!=None:
            sql += " and [years]='"+ str(year) + "'"
        if bid!=None:
            sql += " and [bid]='"+ bid + "'"
        
        resList = self.__ExecQuery(sql)
        i=1
        for (bid,num,years) in resList:
            print bid,num,years,i
            i = i+1
        
    def TestHttp(self):
        url = "http://book2.duxiu.com/search?channel=search&sw=%CE%EF%C0%ED&Field=1&sectyear=2015&Pages=1"
        content = self.__HttpGet(url)
        print content.decode("utf8")

    def TestHttp2(self):
        url = "http://www.duxiu.com/loginhl.jsp?send=true&UserName=shr&PassWord=shr"
        content = self.__HttpPost(url,"")
        print content.decode("utf8")

    def TestParse(self):
        url = "http://acad.cnki.net/Kns55/oldnavi/n_issue.aspx?NaviID=1&Field=year&BaseID=DLKJ&Value=2014"
        content = self.__HttpGet(url)
        resList = self.__ParseNumList("DLKJ",content)
        for item in resList:
            print item.Bid,item.Years,item.Num

    def TestSaveToSqlServer(self):
        sql = ""
        sql += "insert into modify_bid_num_info([bid],[num],[years])"
        sql += "values("
        sql += "'AAAK','03','2014'"
        sql += ")"
        print sql.encode("utf8")
        try:
            self.__ExecNonQuery(sql)
            print(u'插入成功。')
        except:
            print(u'插入失败。')

    def TestExist(self):
        item = QkNum('ZZZZ','2000','01')
        item = QkNum('ZZZZ','2000','99')
        if self.__Exist(item):
            print u'已存在',repr(item)
        else:
            print u'不存在',repr(item)

    def TestDoJob(self,year=None,bid=None):
        if bid!=None and year!=None:
            item = Task(bid,year)
            self.__DoJob(item)
        if bid!=None and year==None:
            today = date.today().year
            fromyear = 1985
            while fromyear<=today:
                item = Task(bid,fromyear)
                self.__DoJob(item)
                fromyear = fromyear+1
            
        


#************ 实体类：TsIndex ************
class TsIndex:
    def __init__(self,Keys=None,FromYear=None,FullTextAddr=None):
        self.__InitField()
        if Keys is not None:
            self.Keys = Keys
        if FromYear is not None:
            self.FromYear=FromYear
	if FullTextAddr is not None:
            self.FullTextAddr=FullTextAddr

    def __InitField(self):
        self.Keys = ""
	self.FromYear=""
	self.UpdateTime=""
	self.FullTextAddr=""

class Content:
    def __init__(self,site=None):
        self.__InitField()
        if site is not None:
            self.site = site

    def __InitField(self):
        self.Keys = ""
	self.site=""
	self.UpdateTime=""

class Task:
    def __init__(self,sw,year=None):
        self.__InitField()
        self.sw = sw
        if year is not None:
            self.year=year
            self.page=760
            self.BaseUrl=self.__GetBaseUrl(sw,year)
            self.TaskUrl=self.BaseUrl+"&pn=10"
            #self.TaskUrl=self.BaseUrl+"&page=10"

    def Next(self):
        self.page = self.page+10
        self.TaskUrl=self.BaseUrl+"&pn="+str(self.page)
        #self.TaskUrl=self.BaseUrl+"&page="+str(self.page)

    def __InitField(self):
        self.id = ""	
	self.sw=""
	self.year=""
	self.page=0
	self.BaseUrl=""
	self.TaskUrl=""

    def __GetBaseUrl(self,sw,year):
        url = ""
        url += "http://www.baidu.com/s?wd=inurl%3A%20(http%3A%2F%2Fwww.sex%20.com)&oq=inurl%3A%20(http%3A%2F%2Fwww.sex%20.com)&tn=baiduadv&ie=utf-8&rsv_pq=8b543c2800033cc4&rsv_t=e6d5c43AURwRezQz9LGpeCSRbkhgOyLMmgl4bmaC9iCY5c2sFOAlx6YJzJ5WddA&rsv_page=1"
        #url += "http://www.baidu.com/s?wd=site&oq=site&ie=utf-8&usm=1&rsv_idx=1&rsv_pq=9a73468e0002153b&rsv_t=8bcbbcvjidpZvnRZBRykFHYX80or22BoWxzY%2Fxpjr%2Bn2P5UpgurnL%2FMGvuI&rsv_page=1"
        #url += "http://www.sogou.com/web?query=site&sut=850&lkt=5%2C1445933425253%2C1445933426130&sst0=1445933426236&ie=utf8&p=40040100&dp=1&w=01019900&dr=1"
        #print url

        return url
