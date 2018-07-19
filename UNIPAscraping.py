from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from xlwt import Workbook

URL = 'https://portal.sa.dendai.ac.jp/up/faces/login/Com00505A.jsp'
ID = 'ID'
PASS = 'PASS'
YOBI = '1'
JIGEN = '1'


#ブラウザ起動
options = Options()
options.add_argument('--headless')
driver = webdriver.Chrome('chromedriver.exe',chrome_options=options)
driver.get(URL)

#Excel初期化
wb = Workbook()
ws = wb.add_sheet("YOBI="+YOBI+" JIGEM="+JIGEN)
ws.write(0,0,"YOBI="+YOBI+"JIGEM="+JIGEN)
ws.write(0,1,"授業コード")
ws.write(0,2,"授業名")
ws.write(0,3,"担当教員")
ws.write(0,4,"該当学期")
ws.write(0,5,"曜日")
ws.write(0,6,"時限")
ws.write(0,7,"教室番号")
ws.write(0,8,"単位数")
ws.write(0,9,"目的概要")
ws.write(0,10,"達成目標")
ws.write(0,11,"関連科目")
ws.write(0,12,"履修条件")
ws.write(0,13,"教科書名")
ws.write(0,14,"評価方法")
ws.write(0,15,"学習・教育目標との対応")
ws.write(0,16,"事前・事後学習")
ws.write(0,17,"E-Mail address")
ws.write(0,18,"質問への対応")
ws.write(0,19,"履修上での注意事項")
ws.write(0,20,"学習上の助言")
ws.write(0,21,"該当ユニット")
ws.write(0,22,"種類")

#ログイン
driver.find_element_by_id('form1:htmlUserId').send_keys(ID)
driver.find_element_by_id('form1:htmlPassword').send_keys(PASS)
driver.find_element_by_id('form1:login').click()

print('ログイン成功')
actions = ActionChains(driver)
actions.move_to_element(driver.find_element_by_id('menu5')).perform()
# menu = driver.find_element_by_id('menu5')
# driver.getMouse().mouseMove(menu)
driver.find_element_by_id('menuimg5-2').click()


print('シラバス検索ページ')
selector1 = Select(driver.find_element_by_id('form1:htmlYobi'))
selector2 = Select(driver.find_element_by_id('form1:htmlJigen'))
selector1.select_by_value(YOBI)
selector2.select_by_value(JIGEN)
driver.find_element_by_id('form1:search').click()


print('検索完了ページ')
table = driver.find_element_by_xpath("html/body/div/div/form[3]/table/tbody/tr[4]/td[2]/b/table/tbody")
trs = table.find_elements_by_class_name("rowClass1")
print(len(trs))
driver.save_screenshot("test.png")

# forループ
clasnum=0
for i in range(len(trs)):
    clasnum+=1;
    driver.find_element_by_id('form1:htmlKekkatable:'+str(i)+':edit').click()
    print("nowCaptureing"+str(i))
    table2 = driver.find_element_by_xpath("html/body/div/div/form[3]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/div/table/tbody/tr/td/table/tbody")
    trs2_1 = table2.find_elements_by_tag_name("th")
    trs2_2 = table2.find_elements_by_tag_name("td")
    number = 0
    for atd in trs2_2:
        print(atd.text)
        number+=1
        ws.write(clasnum,number,atd.text)

    driver.find_element_by_id('form1:back00').click()



wb.save("YOBI="+YOBI+" JIGEM="+JIGEN+".xls")
driver.quit()