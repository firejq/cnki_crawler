import time
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

from util import util


def scraping(driver, author_name, author_company='深圳大学'):
    if util.is_done(author_name):
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

        time.sleep(4)

        # TODO 改用显式等待
        # try:
        #     wait = WebDriverWait(driver=driver, timeout=10)
        #     wait.until(expected_conditions.presence_of_element_located((By.XPATH, '//*[@id="iframeResult"]')))
        #     print('wait over')
        # except TimeoutError:
        #     print('超时')
        #     exit()

        print('开始抓取【' + author_name + '】的论文信息')
        # 切换至frameResult
        driver.switch_to_frame('iframeResult')
        # 获取结果记录条数
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
                util.write_to_excel(res, author_name=author_name)
        else:
            # 结果数量超过20条，不止一页
            # 获取最大页数
            max_page = int(
                driver.find_element_by_css_selector('div.TitleLeftCell').find_elements_by_tag_name('a')[-2].text)
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
                    util.write_to_excel(res, author_name=author_name)
                # 点击下一页按钮
                if i != max_page:
                    next_page = driver.find_element_by_css_selector('div.TitleLeftCell').find_elements_by_tag_name('a')[-1]
                    next_page.click()

        print('【' + author_name + '】的所有论文信息抓取完毕\n')