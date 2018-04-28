# coding=utf-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from tqdm import tqdm
import time
import csv

NAME = "急性心肌梗塞疾病"

YEAR_LIST = [
    "10006", "10013", "10012", "10106", "10113", "10112", "10206", "10213", "10212", "10306", "10313", "10312", "10406", "10413", "10412", "10506", "10513", "10512", "10606"
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

AMI_OPTION = [
    "1367", "1369", "1371", "1373", "1375", "1379", "1381", "1383", "1385", "1389", "1391", "1393", "1395", "1397", "1399", "1401", "1411", "1415"
]

AMI_DICT = {
    "1367": "AMI案件於住院期間執行血脂LDL檢查比率",
    "1369": "AMI案件住院期間給予Aspirin比率",
    "1371": "AMI案件住院期間給予ADP受體拮抗劑(如Clopidogrel類)比率",
    "1373": "AMI案件住院期間給予β-Blocker比率",
    "1375": "AMI案件住院期間給予ACE inhibitor或ARB比率",
    "1379": "AMI案件出院後3個月內使用Aspirin比率",
    "1381": "AMI案件出院後6個月內使用Aspirin比率",
    "1383": "AMI案件出院後9個月內使用Aspirin比率",
    "1385": "AMI案件出院後3個月內使用ADP受體拮抗劑(如Clopidogrel類) 比率",
    "1387": "AMI案件出院後6個月內使用ADP受體拮抗劑(如Clopidogrel類) 比率",
    "1389": "AMI案件出院後9個月內使用ADP受體拮抗劑(如Clopidogrel類) 比率",
    "1391": "AMI案件出院後3個月內使用β-Blocker比率",
    "1393": "AMI案件出院後6個月內使用β-Blocker比率",
    "1395": "AMI案件出院後9個月內使用β-Blocker比率",
    "1397": "AMI案件出院後3個月內使用ACE inhibitor或ARB比率",
    "1399": "AMI案件出院後6個月內使用ACE inhibitor或ARB比率",
    "1401": "AMI案件出院後9個月內使用ACE inhibitor或ARB比率",
    "1411": "AMI案件出院後3日內因主診斷為AMI或相關病情之急診返診比率",
    "1415": "AMI案件出院14天內(含)因主診斷AMI或相關病情之非計畫性再住院比率"
}

def dict2CSV(data, year, option):
    keys = data[0].keys()
    with open('{}/{}-{}.csv'.format(NAME, YEAR_DICT[year], AMI_DICT[option]), 'w', encoding="utf8") as f:
        dict_writer = csv.DictWriter(f, keys)
        dict_writer.writeheader()
        dict_writer.writerows(data)


def checkExist(year, option):
    import os.path
    if not os.path.isdir(NAME):
        os.makedirs(NAME)
    if os.path.isfile("{}/{}-{}.csv".format(NAME, YEAR_DICT[year], AMI_DICT[option])):
        return True
    return False


def getTableData(year, option):
    if checkExist(year, option):
        return

    DATA = []
    driver = webdriver.Firefox()
    try:
        driver.get("http://www1.nhi.gov.tw/mqinfo/SearchPro.aspx?Type=AMI&List=4")

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

        with tqdm(total=len(table.find_elements_by_tag_name("tr")-1)) as pbar:
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

for option in AMI_OPTION:
    for year in YEAR_LIST:
        print("{} {}".format(YEAR_DICT[year], AMI_DICT[option]))
        getTableData(year, option)
    