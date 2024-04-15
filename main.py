from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import pandas as pd
from typing import List, Dict
import undetected_chromedriver
from selenium.webdriver import ChromeOptions
import os
import glob
import shutil
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, NoSuchElementException
import logging

logging.basicConfig(filename='rpa_test.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                    encoding='utf-8')
ERROR_TEXT = "Программа завершилась с ошибкой, обратитесь к администратору"


def get_driver():
    """Создает драйвер"""
    logging.info('Начали создавать драйвер')
    # mac/linux chromedriver path = ./chromedriver
    driver = undetected_chromedriver.Chrome(driver_executable_path="chromedriver.exe")
    driver.maximize_window()
    logging.info('Закончили создание драйвера')
    return driver


def get_data_from_excel(filename: str) -> List[Dict[str, any]]:
    """Возвращает данные из excel с именованными столбцами в формате списка словарей"""
    logging.info('Начали считывать данные для вставки из excel')
    df = pd.read_excel(filename)
    df.columns = df.columns.str.strip()
    result = df.to_dict(orient="records")
    logging.info('Завершили считывание и преобразование данных из excel')
    return result


def download_file(download_link):
    """Скачивает файл"""
    logging.info("Начали скачивать файл")
    try:
        download_link.click()
        logging.info("Кликнули на кнопку скачивания файла с данными")
    except ElementClickInterceptedException:
        logging.error("Не получилось кликнуть на кнопку скачивания, кнопка не кликабельна")
        return None
    filename = move_last_downloaded_excel()
    attempt_count = 0
    while filename is None and attempt_count < 10:
        attempt_count += 1
        time.sleep(0.35)
        filename = move_last_downloaded_excel()
    if filename is None:
        logging.error(f"Не получилось найти файл с данными, было сделано {attempt_count} попыток")
        return None
    logging.info("Успешно скачали и переместили файл с данными")
    return filename


def move_last_downloaded_excel():
    """Перемещаем файл из загрузок в текузую папку"""
    # Путь к стандартной папке "Загрузки" на Windows
    logging.info('Начали копирование файла с данными из загрузок в папку проекта')
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
    # Поиск всех файлов с расширением .xlsx в папке "Загрузки"
    excel_files = glob.glob(os.path.join(downloads_folder, 'challenge*.xlsx'))
    if not excel_files:  # проверка на наличие файлов
        return None
    excel_files_filtered = [f for f in excel_files if os.path.getctime(f) > time.time() - 3 * 60]
    if not excel_files_filtered:  # проверка, что есть новые файлы, подходящие под наши требования
        return None
    # Сортировка файлов по времени последнего изменения (самый последний файл будет последним в списке)
    excel_files.sort(key=os.path.getmtime, reverse=True)
    # Выбор последнего файла
    last_downloaded_excel = excel_files[0]
    # Копирование последнего файла в текущую директорию
    shutil.copy(last_downloaded_excel, os.getcwd())
    # Удаление последнего файла из папки "Загрузки"
    os.remove(last_downloaded_excel)
    # Получение имени файла
    file_name = os.path.basename(last_downloaded_excel)
    logging.info('Копирование файла из загрузок прошло успешно')
    return file_name


def get_form_field(driver):
    try:
        form = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "form"))
        )
        logging.info(f'Начали искать форму')
    except TimeoutException:
        logging.error('Не получилось найти форму для вставки данных code1. Работа прекращена')
        print(ERROR_TEXT)
        return None, None
    try:
        main_div = form.find_element(By.CSS_SELECTOR, ".row")
    except NoSuchElementException:
        logging.error('Не получилось найти форму для вставки данных code2. Работа прекращена')
        print(ERROR_TEXT)
        return None, None
    fields = main_div.find_elements(By.TAG_NAME, "rpa1-field")
    if len(fields) == 0:
        logging.error('Не получилось найти форму для вставки данных code3. Работа прекращена')
        print(ERROR_TEXT)
        return None, None
    logging.info('Нашли все поля формы')
    return form, fields


def is_data_paste_right(data, driver):
    logging.info('Начинаем проверку правильности заполнения формы')
    form, fields = get_form_field(driver)
    if form is None or fields is None:
        return None
    for field in fields:
        param_name = field.get_attribute("ng-reflect-dictionary-value")
        try:
            input_element = field.find_element(By.TAG_NAME, "input")
        except NoSuchElementException:
            logging.error(f'Не получилось найти форму {param_name} для проверки её заполнения. Работа прекращена')
            print(ERROR_TEXT)
            return None
        value = input_element.get_attribute('value')
        if str(data[param_name.strip()]).strip() != value.strip():
            logging.error('Проверка завершена, форма заполнена неправильно')
            return False
    logging.info('Проверка формы завершена, все ОК')
    return True


def main():
    logging.info('Старт программы')
    driver = get_driver()
    driver.get("https://rpachallenge.com/")
    logging.info('Загрузили страницу')
    # Скачиваем файл и забираем из него данные
    try:
        logging.info('Начинаем поиск ссылки для скачивания файла с данными')
        download_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a"))
        )
    except TimeoutException:
        logging.error('Не получилось найти ссылку для загрузки файла с данными, работа прекращена')
        print(ERROR_TEXT)
        return
    data_filename = download_file(download_link)
    if data_filename is None:
        logging.error('Не получилось загрузить файл, работа прекращена')
        print(ERROR_TEXT)
        return
    try:
        logging.info('Начинаем поиск кнопки для начала работы')
        start_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button"))
        )
        logging.info('Нашли кнопку для старта')
    except TimeoutException:
        logging.error('Не получилось найти ссылку для загрузки файла с данными, работа прекращена')
        print(ERROR_TEXT)
        return
    try:
        start_button.click()  # находим и кликаем кнопку старта
        logging.info('Кликнули на кнопку для старта')
    except ElementClickInterceptedException:
        logging.error("Не получилось кликнуть на кнопку начало работы, кнопка не кликабельна. Работа завершена")
        print(ERROR_TEXT)
        return
    paste_data = get_data_from_excel(data_filename)  # получаем из excel данные для подстановки
    actions = ActionChains(driver)
    logging.info('Начинаем вставлять данные из excel в формы')
    for paste_data_item in paste_data:
        logging.info(f'Начали работу для набора данных {str(paste_data_item)}')
        form, fields = get_form_field(driver)
        if form is None or fields is None:
            logging.error('Не получилось найти поля или форму. Работа программы прекращена')
            print(ERROR_TEXT)
            return
        logging.info('Начинаем заполнение полей формы')
        for item in fields:  # Находим все инпуты, прокликиваем и вставляем данные
            param_name = item.get_attribute("ng-reflect-dictionary-value")
            try:
                item.click()
            except ElementClickInterceptedException:
                logging.error(f"Не получилось кликнуть по полю {param_name}. Работа прекращена")
                print(ERROR_TEXT)
                return
            actions.send_keys(paste_data_item[param_name]).perform()
            logging.info(f'Заполнили поле {param_name}')
        # Проверяем вставились ли данные
        is_data_paste_right_value = is_data_paste_right(paste_data_item, driver)
        if is_data_paste_right_value is None:
            logging.error('Поля заполнены неправильно, работа программы прекращена')
            print(ERROR_TEXT)
            return
        if not is_data_paste_right_value:
            logging.error('Поля заполнены неправильно, работа программы прекращена')
            print(ERROR_TEXT)
            return
        try:
            button = form.find_element(By.CLASS_NAME, "btn.uiColorButton")
            logging.info('Нашли кнопку подтверждения формы')
        except NoSuchElementException:
            logging.error('Не получилось найти кнопку подтверждения формы, работа прекращена')
            print(ERROR_TEXT)
            return
        try:
            button.click()  # находим и прокликиваем кнопку подтверждения формы
        except ElementClickInterceptedException:
            logging.error(f"Не получилось кликнуть по кнопке подтверждения формы. Работа прекращена")
            print(ERROR_TEXT)
            return
        logging.info('Подтвердили форму')
    logging.info("Завершили вставку")
    try:
        logging.info('Ищем результат работы программы')
        result = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, ".message2"))
        )
        print(result.text)
        logging.info(f'Результат работы программы: {result.text}')
    except TimeoutException:
        logging.error('Результат работы программы не найден, работа прекращена')
        print(ERROR_TEXT)
        return
    if os.path.exists(data_filename):  # проверяем существует ли файл
        os.remove(data_filename)  # если да, то удаляем его
        logging.info(f"Файл {data_filename} успешно удалён.")
    else:
        logging.info(f"Файл {data_filename} не существует.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.error('Произошла ошибка: неизвестная ошибка %s', e, exc_info=True)
        print(ERROR_TEXT)
