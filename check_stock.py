import xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")
driver = webdriver.Chrome(options=options)
driver.get("https://www.bourse.lu/issuer-securities/Turkey/26760")
x = 3
y = 0
workbook = xlsxwriter.Workbook(r"Where you want the file to be saved, add name with .xlsx")
worksheet = workbook.add_worksheet()
try:
    while x > 0:
        date = driver.find_element(By.XPATH, "//*[@id=\"section-content-col-1\"]/div/div/div/div[2]/ul/li[%s]/a/div[2]/"
                                             "div/div[1]/h3" % x).text
        name = driver.find_element(By.XPATH, "//*[@id=\"section-content-col-1\"]/div/div/div/div[2]/ul/li[%s]/a/div[2]/"
                                             "div/div[1]/p/span[1]" % x).text
        price = driver.find_element(By.XPATH, "//*[@id=\"section-content-col-1\"]/div/div/div/div[2]/ul/li[%s]/a/div[2]"
                                              "/div/div[1]/p/span[2]" % x).text
        percentage = driver.find_element(By.XPATH, "//*[@id=\"section-content-col-1\"]/div/div/div/div[2]/ul/li[%s]/a/d"
                                                   "iv[2]/div/div[2]/p[1]" % x).text
        other_date = driver.find_element(By.XPATH, "//*[@id=\"section-content-col-1\"]/div/div/div/div[2]/ul/li[%s]/a/d"
                                                   "iv[2]/div/div[2]/p[2]/b" % x).text
        worksheet.write(y, 0, date)
        worksheet.write(y, 1, name)
        worksheet.write(y, 2, price)
        worksheet.write(y, 3, percentage)
        worksheet.write(y, 4, other_date)
        x += 1
        y += 1
except:
    x = 0
workbook.close()
