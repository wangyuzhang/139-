#coding:utf-8
"""
Action：WebDriver + Python 调用AutoIt例子（139邮箱写信页的附件上传操作）
Author：深圳-横放
Date：2014-07-02
"""

#调用autoit所需的模块
import win32api, win32pdhutil, win32con 
import win32com.client
from win32com.client import Dispatch

#webdriver的模块
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import unittest, time, re ,sys ,os

class Untitled(unittest.TestCase):
    username = "yfjelley@139.com" #账号
    password = "8143181123" #密码
    
    #初始化
    def setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(30)
        self.base_url = "http://mail.139.com"
        self.verificationErrors = []
        self.accept_next_alert = True
        self.driver.maximize_window()

    #判断元素是否出现
    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException, e: return False
        return True
    
    #通过xpath匹配对应的文本元素，判断文本元素是否出现
    def is_text_present(self,what):
        #Now = datetime.now()
        try:
            if what=="":
                pass
            else:
                self.driver.find_element_by_xpath("//*[contains(.,'%s')]"%what)
        except NoSuchElementException, e:
            #elapsedtime1 = datetime.now()-Now
            #print "waste time: "+str(elapsedtime1.seconds)+"."+str(elapsedtime1.microseconds)
            return False
        #elapsedtime1 = datetime.now()-Now
        #print "waste time: "+str(elapsedtime1.seconds)+"."+str(elapsedtime1.microseconds)
        return True
    
    #测试用例--发送一封测试邮件
    def test_untitled(self):
        driver = self.driver
        driver.get("http://mail.10086.cn/")
        
        #登录
        driver.find_element_by_id("txtUserName").clear()
        driver.find_element_by_id("txtUserName").send_keys(self.username.split('@')[0])
        driver.find_element_by_id("txtPassword").clear()
        driver.find_element_by_id("txtPassword").send_keys(self.password)
        driver.find_element_by_id("loginBtn").click()

        #写信页打开
        driver.find_element_by_css_selector("#btn_compose > span").click()#打开写信页
        time.sleep(2)

        #选择iframe
        iframe_xpath = "//iframe[contains(@id,'compose_')]"#写信页所在的iframe的xpath路径，通过火狐浏览器的firebug插件可以查看到是在一个iframe的页面内
        iframe = self.driver.find_element_by_xpath(iframe_xpath)
        self.driver.switch_to_frame(iframe)
        time.sleep(1)

        #编写邮件内容        
        self.driver.find_element_by_link_text(u"发给自己").click()
        mysubject = "这是一封测试邮件！ "
        mycontent = "这是一封测试邮件！ "*5

        #这里有个小技巧--对于不能正常输入文本的输入框，可以先用webdriver点击，然后再用AutoIt的send方法进行输入
        self.driver.find_element_by_css_selector("input.fl.addrText-input").click()
        autoit.Send(mysubject)
        
        autoit.MouseClick("left",448,452,1)#正文-坐标可以通过autoit自带的Autoit Window info来查看
        autoit.Send(mycontent)

        self.driver.find_element_by_id("uploadInput").click()#点击附件上传按钮
        #autoit.MouseClick("left", 288, 326, 1)#点击附件上传按钮--用上面的代码替代
        
        str_filepath = os.path.join(os.getcwd() + "\\1M.png") #获取附件路径
        #print str_filepath
        
        autoit.WinWait(u"打开", "", 5)
        autoit.WinActivate(u"打开")        
        autoit.ControlSetText(u"打开","","[CLASS:Edit; INSTANCE:1]",str_filepath)
        autoit.ControlClick(u"打开","",u"保存(&S)")
        autoit.ControlClick(u"打开","",u"打开(&O)") #附件上传动作

        #等待附件上传完成
        while self.is_text_present("(1,031.79K)") == False:
            sleep(1)
                    
        #点击发送
        self.driver.find_element_by_xpath("//*[@id=\"topSend\"]/span").click()
        time.sleep(10)


    #测试清理    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
      
    #调用autoit3
    try:
        autoit = Dispatch("AutoItX3.Control")
    except:
        print >> sys.stderr, u'AutoItX3 加载失败了！'
        
    unittest.main()