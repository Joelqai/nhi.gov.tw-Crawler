# coding=utf-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from tqdm import tqdm
import time
import csv

NAME = "子宮肌瘤手術"

YEAR_LIST = [
    "9512", "9612", "9712", "9812", "9912", "10006", "10013", "10012", "10106", "10113", "10112", "10206", "10213", "10212", "10306", "10313", "10312", "10406", "10413", "10412", "10506", "10513", "10512", "10606"
]

YEAR_DICT = {
    "9512": "95年全年",
    "9612": "96年全年",
    "9712": "97年全年",
    "9812": "98年全年",
    "9912": "99年全年",
    "10006": "100年上半年",
    "10013": "100年下半年",
    "10012": "100年全年",
    "10106": "101年上半年",
    "10113": "101年下半年",
    "10112": "101年全年",
    "10206": "102年上半年",
    "10213": "102年下半年",
    "10212": "102年全年",
    "10306": "103年上半年",
    "10313": "103年下半年",
    "10312": "103年全年",
    "10406": "104年上半年",
    "10413": "104年下半年",
    "10412": "104年全年",
    "10506": "105年上半年",
    "10513": "105年下半年",
    "10512": "105年全年",
    "10606": "106年上半年",
    "10613": "106年下半年",
    "10612": "106年全年"
}

MYOMA_OPTION = [
    "471", "473", "494", "815"
]

MYOMA_DICT = {
    "471": "子宮肌瘤手術當次住院日數大於7日比率",
    "473": "子宮肌瘤手術後14日內因該手術相關診斷再住院率",
    "494": "子宮肌瘤手術後14日內有相關器官之損傷發生率",
    "815": "子宮肌瘤手術後14日內有相關器官之徵候發生率"
}

def dict2CSV(data, year, option):
    keys = data[0].keys()
    with open('{}/{}-{}.csv'.format(NAME, YEAR_DICT[year], MYOMA_DICT[option]), 'w', encoding="utf8") as f:
        dict_writer = csv.DictWriter(f, keys)
        dict_writer.writeheader()
        dict_writer.writerows(data)


def checkExist(year, option):
    import os.path
    if not os.path.isdir(NAME):
        os.makedirs(NAME)
    if os.path.isfile("{}/{}-{}.csv".format(NAME, YEAR_DICT[year], MYOMA_DICT[option])):
        return True
    return False

def getTableData(year, option):
    if checkExist(year, option):
        return

    DATA = []
    driver = webdriver.Firefox()
    try:
        driver.get("http://www1.nhi.gov.tw/mqinfo/SearchPro.aspx?Type=Myoma&List=4")

        driver.find_element_by_id("ContentPlaceHolder1_DropDA").click()
        driver.find_element_by_xpath("//option[@value='{}']".format(option)).click()
        time.sleep(2)
        driver.find_element_by_id("ContentPlaceHolder1_drop1").click()
        driver.find_element_by_xpath("//option[@value='{}']".format(year)).click()
        driver.find_element_by_id("ContentPlaceHolder1_RowBox").send_keys("000")
        driver.find_element_by_id("ContentPlaceHolder1_But_Query").click()
        time.sleep(60)

        table = driver.find_element_by_id("ContentPlaceHolder1_GV_List")
        first_line = True

        with tqdm(total=len(table.find_elements_by_tag_name("tr"))-1) as pbar:
            for row in table.find_elements_by_tag_name("tr"):
                if first_line:
                    first_line = False
                    continue
                column = row.find_elements_by_tag_name("td")
                info = {
                    "縣市別": column[1].text,
                    "醫事機構名稱": column[2].text,
                    "特約類別": column[3].text,
                    "分子": column[4].text,
                    "分母": column[5].text,
                    "院所指標值": column[6].text,
                    "所屬分區業務組指標值": column[7].text,
                    "全國指標值": column[8].text
                }
                DATA.append(info)
                pbar.update(1)
        dict2CSV(DATA, year, option)
    except Exception as e:
        print(e)
    driver.close()

for option in MYOMA_OPTION:
    for year in YEAR_LIST:
        print("{} {}".format(YEAR_DICT[year], MYOMA_DICT[option]))
        getTableData(year, option)
    