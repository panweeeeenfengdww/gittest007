import xlrd
from selenium import webdriver

def readExcel(sheet, startRow, endRow):
    a = xlrd.open_workbook(r"C:\Users\Maibenben\Desktop\data.xls")
    b = a.sheet_by_name(sheet)
    cols = b.ncols
    print(cols)
    lst = []
    for i in range(startRow, endRow):
        if cols > 1:
            lst1 = []
            for j in range(cols):
                c = b.cell_value(i, j)
                if type(c) == float:
                    c = int(c)
                lst1.append(c)
            lst.append(lst1)
        elif cols == 1:
            c = b.cell_value(i, 0)
            if type(c) == float:
                c = int(c)
            lst.append(c)
    return lst
def dongtaic(driver,ding):
    try:
        driver.implicitly_wait(2)
        driver.find_element_by_xpath('/html/body/div[1]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[2]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[3]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[4]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[5]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[6]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[7]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[8]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[9]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[10]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[11]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[12]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[13]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[14]'+ding).click()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[15]'+ding).click()
    except:
        a = 1

def dongtais(driver,ding,neirong):
    try:
        driver.implicitly_wait(2)
        driver.find_element_by_xpath('/html/body/div[1]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[2]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[3]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[4]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[5]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[6]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[7]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[8]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[9]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[10]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[11]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[12]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[13]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[14]'+ding).send_keys(neirong)
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[15]'+ding).send_keys(neirong)
    except:
        a = 1

def dongtaicc(driver,ding):
    try:
        driver.implicitly_wait(2)
        driver.find_element_by_xpath('/html/body/div[1]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[2]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[3]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[4]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[5]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[6]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.2)
        driver.find_element_by_xpath('/html/body/div[7]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[8]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[9]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[10]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[11]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[12]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[13]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[14]'+ding).clear()
    except:
        a = 1
    try:
        driver.implicitly_wait(0.1)
        driver.find_element_by_xpath('/html/body/div[15]'+ding).clear()
    except:
        a = 1