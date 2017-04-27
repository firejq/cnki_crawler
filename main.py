import random

import openpyxl
from pyquery import PyQuery
import time
import os
import openpyxl
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait

def getDriver(browser='chrome'):
    UserAgents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3072.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:46.0) Gecko/20100101 Firefox/46.0',
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.87 Safari/537.36 OPR/37.0.2178.32',
        ''
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/534.57.2 (KHTML, like Gecko) Version/5.1.7 Safari/534.57.2',
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36',
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.101 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586',
        'Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko',
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.106 BIDUBrowser/8.3 Safari/537.36',
        'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.9.2.1000 Chrome/39.0.2146.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 61; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36 Core/1.47.277.400 QQBrowser/9.4.7658.400',
        'Mozilla/5.0 (Linux; Android 5.0; SM-N9100 Build/LRX21V) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/37.0.0.0 Mobile Safari/537.36 MicroMessenger/6.0.2.56_r958800.520 NetType/WIFI',
        'Mozilla/5.0 (iPhone; CPU iPhone OS 7_1_2 like Mac OS X) AppleWebKit/537.51.2 (KHTML, like Gecko) Mobile/11D257 QQ/5.2.1.302 NetType/WIFI Mem/28'
    ]
    if browser == 'chrome':
        options = webdriver.ChromeOptions()
        options.add_argument('user-agent=' + random.choice(UserAgents))
        driver = webdriver.Chrome(executable_path='D:\\software\\web\\webDriver\\chromedriver_win32\\chromedriver.exe',
                                  chrome_options=options)
    elif browser == 'phantomjs':
        dcap = dict(DesiredCapabilities.PHANTOMJS)
        dcap["phantomjs.page.settings.userAgent"] = (
            random.choice(UserAgents)
        )
        driver = webdriver.PhantomJS(executable_path='D:\\software\\web\\phantomJS\\phantomjs-2.1.1-windows\\bin\\phantomjs.exe',
                                     desired_capabilities=dcap)
    return driver

def is_done(author_name):
    '''
    后期使用数据库查重，可代替此函数
    :param author_name: 
    :return: 
    '''
    path = os.getcwd() + os.sep + 'result'
    filename = 'result.xlsx'
    out_path = os.path.join(path, filename)
    if os.path.exists(out_path):
        wb = openpyxl.load_workbook(out_path)
        sheets_name = wb.get_sheet_names()
        if author_name in sheets_name:
            return True
        else:
            return False
    else:
        return False

def write_to_excel(res, author_name):
    path = os.getcwd() + os.sep + 'result'
    filename = 'result.xlsx'
    out_path = os.path.join(path, filename)


    if os.path.exists(out_path):
        wb = openpyxl.load_workbook(out_path)
        sheets_name = wb.get_sheet_names()
        if author_name in sheets_name:
            ws = wb.get_sheet_by_name(author_name)
            # print('该作者的论文信息已收集过，进入下一个作者的信息收集')
            # return
        else:
            ws = wb.create_sheet(author_name)
            titles = ['题名', '作者', '来源', '发表时间', '数据库']
            line = [title for title in titles]
            ws.append(line)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = author_name
        titles = ['题名', '作者', '来源', '发表时间', '数据库']
        line = [title for title in titles]
        ws.append(line)

    line = [r for r in res]
    ws.append(line)
    wb.save(out_path)
    # print('成功添加一条【' + author_name + '】的记录')

def scraping(driver, author_name, author_company='深圳大学'):
    if is_done(author_name):
        print('【' + author_name + '】的论文信息已收集过，进行下一个作者论文信息的收集')
        return
    else:
        try:
            author_name_element = driver.find_element_by_id(id_='au_1_value1')
            author_company_element = driver.find_element_by_id(id_='au_1_value2')
            author_name_element.send_keys(author_name)
            author_company_element.send_keys(author_company)
            # time.sleep(1)
            author_company_element.send_keys(Keys.ENTER)
        except NoSuchElementException:
            print('element cannot be found!')

        time.sleep(3)
        print('开始抓取【' + author_name + '】的论文信息')
        # 切换至frameResult
        driver.switch_to_frame('iframeResult')

        res_number = int(str(driver.find_element_by_css_selector('div.pagerTitleCell').text).split(' ')[2])
        if res_number <= 20:
            # 结果数量不足20条，只有一页
            # 抓取结果列表
            tbody = driver.find_element_by_xpath('//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody')
            # 点击所有的“显示全部作者”按钮
            showAlls = tbody.find_elements_by_class_name('showAll')
            for e in showAlls:
                e.click()
            # 遍历结果的每条记录
            trs = tbody.find_elements_by_tag_name('tr')
            del trs[0]
            for tr in trs:
                tds = tr.find_elements_by_tag_name('td')
                res = []  # 存放每一条结果记录
                for td in tds[1:6]:
                    res.append(td.text)
                print(res)
                # TODO 查重：在数据库中检查这条目是否以及存在
                # TODO 将该条记录写入excel
                write_to_excel(res, author_name=author_name)
        else:
            # 结果数量超过20条，不止一页
            # 获取最大页数
            max_page = int(driver.find_element_by_css_selector('div.TitleLeftCell').find_elements_by_tag_name('a')[-2].text)
            for i in range(1, max_page + 1):
                print('开始抓取第' + str(i) + '页')
                # 使用xpath定位抓取结果列表
                tbody = driver.find_element_by_xpath('//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody')
                # 点击所有的“显示全部作者”按钮
                showAlls = tbody.find_elements_by_class_name('showAll')
                for e in showAlls:
                    e.click()
                # 遍历结果的每条记录
                trs = tbody.find_elements_by_tag_name('tr')
                del trs[0]
                for tr in trs:
                    tds = tr.find_elements_by_tag_name('td')
                    res = []  # 存放每一条结果记录
                    for td in tds[1:6]:
                        res.append(td.text)
                    print(res)
                    # TODO 查重：在数据库中检查这条目是否以及存在
                    # TODO 将该条记录写入excel
                    write_to_excel(res, author_name=author_name)
                # 点击下一页按钮
                if i != max_page:
                    next_page = driver.find_element_by_css_selector('div.TitleLeftCell').find_elements_by_tag_name('a')[-1]
                    next_page.click()

        print('【' + author_name + '】的所有论文信息抓取完毕\n')


driver = getDriver('chrome')
url = 'http://kns.cnki.net/kns/brief/result.aspx?dbprefix=scdb&action=scdbsearch&db_opt=SCDB'
authors = [
    "陈思平", "陈昕", "汪天富", "谭力海", "彭珏", "但果", "叶继伦", "覃正笛",
    "张旭", "张会生", "钱建庭", "丁惠君", "刁现芬", "沈圆圆", "周永进", "孔湉湉",
    "陆敏华", "张新宇", "孙怡雯", "李乔亮", "齐素文", "徐海华", "倪东", "刘维湘",
    "李抱朴", "黄炳升", "徐敏", "雷柏英", "胡亚欣", "何前军", "郑介志", "常春起",
    "陈雯雯", "罗永祥", "黄鹏", "林静", "王倪传", "刘立", "张治国", "董磊"
]
for author in authors:
    # try:
    #     driver.get(url=url)
    #     time.sleep(1)
    #     # WebDriverWait(driver, 10).until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR, '.pageBar_bottom')))
    #     print('页面基本加载完毕')
    # except TimeoutError:
    #     print('timeout!')
    #     exit(0)
    driver.get(url=url)
    time.sleep(1)
    scraping(driver=driver, author_name=author)
print('名单上所有作者的论文信息抓取完毕')
exit()

# TODO 将名单提取到配置文件中
# TODO 优化目录结构
# TODO 去重
# TODO 提高性能--多线程 & 多进程 & 异步IO


