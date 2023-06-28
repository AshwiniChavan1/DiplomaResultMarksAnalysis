from selenium import webdriver
from selenium.webdriver.remote import webelement
from selenium.webdriver.remote import webdriver as wd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
import regex as re

import pandas as pd

en_no_label = "Enrollment No."
percentage_label = "Percentage"
pass_fail_label = "P/F"

oop = {
    "title": "OBJECT ORIENTED PROGRAMMING USING C++",
    "percent_label": "OOP %",
    "pf_label": "OOP",
}

dsc = {"title": "DATA STRUCTURE USING ‘C’", "percent_label": "DSC %", "pf_label": "DSC"}

cg = {"title": "COMPUTER GRAPHICS", "percent_label": "CG %", "pf_label": "CG"}

dms = {
    "title": "DATABASE MANAGEMENT SYSTEM",
    "percent_label": "DMS %",
    "pf_label": "DMS",
}

dt = {"title": "DIGITAL TECHNIQUES", "percent_label": "DT %", "pf_label": "DT"}

subjects = [oop, dsc, cg, dms, dt]

xl_path = "D:\\repos\\python\\polytechnic-marks-auto\\output.xlsx"

polytechnic_url = "https://msbte.org.in/pcwebBTRes/pcResult01/pcfrmViewMSBTEResult.aspx"
enrollno_select_xpath = "//*[@id='ddlEnrollOrSeatNo']"
show_result_btn_xpath = "//*[@id='btnShowResult']"

content_xpath = "//*[@id='divContent']/div"
close_result_btn_xpath = "//*[@id='btnClose']"
result_grids_xpath = "//*[@id='divContent']/div/div[2]"
semester_cells_relative_xpath = "./table/tbody/tr[2]/td[7]"

percent_relative_xpath = "./div[contains(@id,'dvTotal')]/table/tbody/tr[2]/td[2]"
pass_fail_relative_xpath = "./div[contains(@id,'dvTotal')]/table/tbody/tr[3]/td[2]"
course_titles_relative_xpath = (
    "./div[contains(@id,'dvMain')]/table/tbody/tr/td[1]"  # skip first two rows
)

obt_percent_xpath = lambda i: f"./div[contains(@id,'dvMain')]/table/tbody/tr[{i}]/td[6]"


xl_file = pd.read_excel(xl_path, converters={en_no_label: int})

remaining = xl_file[
    (xl_file["DT"] != "PASS")
    & (xl_file["DT"] != "FAIL")
    & (xl_file[en_no_label] > 1000000000)
][en_no_label]

xl_file[en_no_label] = xl_file[en_no_label].astype(str)


def third_sem_grid(driver: wd.WebDriver) -> tuple[int, webelement.WebElement]:
    for i, grid in enumerate(driver.find_elements(By.XPATH, result_grids_xpath)):
        if (
            grid.find_element(By.XPATH, semester_cells_relative_xpath).text
            == "THIRD SEMESTER"
        ):
            return (i, grid)
    return None


def default_sem_grid(driver: wd.WebDriver) -> webelement.WebElement:
    return driver.find_element(By.XPATH, result_grids_xpath)


def select_enroll(driver: wd.WebDriver):
    enrollno_select = Select(driver.find_element(By.XPATH, enrollno_select_xpath))
    enrollno_select.select_by_visible_text("Enrollment No")


def put_enrollno(driver: wd.WebDriver, enrollno):
    enrollno_text_xpath = "//*[@id='txtEnrollOrSeatNo']"
    driver.find_element(By.XPATH, enrollno_text_xpath).send_keys(str(enrollno))


def wait_for(driver: wd.WebDriver, xpath: str):
    new_wait = WebDriverWait(driver, timeout=30)
    new_wait.until(EC.visibility_of_element_located(locator=(By.XPATH, xpath)))


def get_subject_row(grid: webelement.WebElement, title: str) -> int:
    for i, row in enumerate(grid.find_elements(By.XPATH, course_titles_relative_xpath)):
        if i < 2:
            continue
        if title in row.text:
            return i + 1
    return None


def get_results(driver: webdriver.Chrome) -> dict:
    i, grid = third_sem_grid(driver)
    result = dict()
    pass_fail = grid.find_element(By.XPATH, pass_fail_relative_xpath)
    result[pass_fail_label] = pass_fail.text
    percent = grid.find_element(By.XPATH, percent_relative_xpath).text
    result[percentage_label] = percent
    for subject in subjects:
        rownum = get_subject_row(grid, subject["title"])
        percent = grid.find_element(By.XPATH, obt_percent_xpath(rownum)).text
        percent = re.findall("[0-9]+",percent)[0].lstrip("0")
        percent = percent if percent!='' else '0'
        # percent = (
        #     grid.find_element(By.XPATH, obt_percent_xpath(rownum))
        #     .text.lstrip()
        #     .rstrip()
        #     .rstrip("*")
        #     .lstrip("0")
        # )
        result[subject["pf_label"]] = "PASS" if int(percent) >= 28 else "FAIL"
        result[subject["percent_label"]] = percent
    return result


def update_result(df: pd.DataFrame, enrollno: int, result: dict):
    i = df.index[df[en_no_label] == str(enrollno)].tolist()[0]
    for key in result.keys():
        df[key][i] = result[key]


try:
    driver = webdriver.Chrome(executable_path=".\chromedriver.exe")
    driver.get(polytechnic_url)
    for enrollno in remaining:
        select_enroll(driver=driver)
        print("enroll no", enrollno)
        put_enrollno(driver=driver, enrollno=enrollno)
        driver.find_element(By.ID, 'txtCaptchaHot').click()
        wait_for(driver, content_xpath)
        result = get_results(driver)
        print("result: ", result)
        print(xl_file.columns)
        update_result(xl_file, enrollno, result)
        xl_file.to_excel(xl_path, index=False)
        driver.find_element(By.XPATH, close_result_btn_xpath).click()
finally:
    driver.close()
    driver.quit()
