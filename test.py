# 协助Yvette爬取公司数据
# http://47.105.73.112/
# yvette.ho@power-engine.com.cn
# power168

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime
import xlwt

class taobao_infos:
    
    #对象初始化
    def __init__(self):
        url='http://47.105.73.112/'
        self.url=url

        options=webdriver.ChromeOptions() #配置 chrome 启动属性
        #options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2}) #不加载图片，加快访问速度
        options.add_experimental_option("excludeSwitches",['enable-automation']) # 此步骤很重要，设置为开发者模式，防止被各大网站识别出来使用了Selenium

        self.browser=webdriver.Chrome(executable_path=chromedriver_path,options=options)
        self.wait=WebDriverWait(self.browser,0.1) #超时时长为0.1s

    #输入用户名密码
    def login(self):
        self.browser.get(self.url)
        time.sleep(30)

    #打开需要爬取的页面
    def get_shop(self,shop_url):
        print("正在打开页面")
        self.browser.get(shop_url)
        time.sleep(3)
        try:
            self.browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/div[2]/div[1]/div/div[3]/div[2]/div/div[4]/button').click()
            print(f"已关闭筛选条件")
        except:
            print(f"无筛选条件")
        time.sleep(2)

        print(f"正在创建Excel文件")
        try:
            workbook = xlwt.Workbook(encoding = 'utf-8')
            worksheet = workbook.add_sheet('Yvette Worksheet')
            worksheet.write(0, 0, label = '姓名')
            worksheet.write(0, 1, label = '电话')
            worksheet.write(0, 2, label = '邮箱')
            print(f"成功创建Excel文件")
        except:
            print(f"未成功创建Excel文件")
        time.sleep(2)

        for i in range(2, 52):
            name_string = '//*[@id="app"]/div/div/div[2]/div[2]/div[2]/div[2]/div/div/div[1]/div[' + str(i) + ']/div[1]/div[2]/div[1]/a/span'
            tel_string = '//*[@id="app"]/div/div/div[2]/div[2]/div[2]/div[2]/div/div/div[1]/div[' + str(i) + ']/div[1]/div[3]/div[1]/span/a'
            mail_string = '//*[@id="app"]/div/div/div[2]/div[2]/div[2]/div[2]/div/div/div[1]/div[' + str(i) + ']/div[2]/div[1]/span/a'
            try:
                name = self.browser.find_element_by_xpath(name_string).text #获取名字
                worksheet.write(i-1, 0, label = name)
                print(f"成功保存" + name)
            except:
                print(f"未成功保存")
            time.sleep(1)
            try:
                self.browser.find_element_by_xpath(tel_string).click() #点击查看电话
                time.sleep(2)
                tel = self.browser.find_element_by_xpath(tel_string).text #获取电话
                worksheet.write(i-1, 1, label = tel)
                print(f"成功保存" + name + "的电话")
            except:
                print(f"未成功保存" + name + "的电话")
            time.sleep(1)
            try:
                self.browser.find_element_by_xpath(mail_string).click() #点击查看邮箱
                time.sleep(2)
                mail = self.browser.find_element_by_xpath(mail_string).text #获取邮箱
                worksheet.write(i-1, 2, label = mail)
                print(f"成功保存" + name + "的邮箱")
            except:
                print(f"未成功保存" + name + "的邮箱")
            time.sleep(1)

        workbook.save(r'E:/pythonworkspace/Yvette/Contact.xls')

    def gb(self):
        print("程序已运行完毕，可关闭浏览器")
        #self.browser.quit() #关闭浏览器

if __name__ == '__main__':
    chromedriver_path = r'F:\setupFiles\chromedriver.exe'  #谷歌chromedriver完整路径
    spider=taobao_infos()
    spider.login()
    shop_url="http://47.105.73.112/crm/candidate/list?gql="
    spider.get_shop(shop_url)

    spider.gb()