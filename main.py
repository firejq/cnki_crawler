import time
from util import util
from crawler.scraping import scraping

driver = util.getDriver('phantomjs')
# url = 'http://kns.cnki.net/kns/brief/result.aspx?dbprefix=scdb&action=scdbsearch&db_opt=SCDB'
url = 'http://kns.cnki.net/kns/brief/result.aspx?dbprefix=scdb'


for author in util.get_name_list():
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
driver.quit()
print('名单上所有作者的论文信息抓取完毕')


# TODO 去重
# TODO 提高性能--多线程 & 多进程 & 异步IO
