from driver_service import get_driver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import pandas as pd
from typing import List, Dict


def get_data_from_excel(filename: str) -> List[Dict[str, any]]:
    """Возвращает данные из excel с именованными столбцами в формате списка словарей"""
    df = pd.read_excel(filename)
    df.columns = df.columns.str.strip()
    return df.to_dict(orient="records")


def main():
    paste_data = get_data_from_excel("challenge.xlsx")  # получаем из excel данные для подстановки
    driver = get_driver()
    driver.get("https://rpachallenge.com/")
    start_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".col.s12.m12.l12.btn-large.uiColorButton"))
    )
    start_button.click()  # находим и кликаем кнопку старта
    actions = ActionChains(driver)
    for paste_data_item in paste_data:
        form = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "form"))
        )
        main_div = form.find_element(By.CSS_SELECTOR, ".row")
        fields = main_div.find_elements(By.TAG_NAME, "rpa1-field")
        for item in fields:  # Находим все инпуты, прокликиваем и вставляем данные
            item.click()
            param_name = item.get_attribute("ng-reflect-dictionary-value")
            actions.send_keys(paste_data_item[param_name]).perform()
        button = form.find_element(By.CLASS_NAME, "btn.uiColorButton")
        button.click()  # находим и прокликиваем кнопку подтверждения формы
    time.sleep(10)  # задержка для просмотра результата
    driver.quit()


if __name__ == "__main__":
    main()
