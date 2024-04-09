# код вставляет информацию из excel в поля на сайте https://rpachallenge.com/
## Использованные технологии:
1. Selenium - непосредственно вставка данных в web-форму
2. Pandas - обработка данных из excel
## Требования к софту
python v >= 3.10, браузер googleChrome, windows<br>
## Инструкция по запуску
1. Скачать chromedriver, подходящий под вашу версию Chrome и добавить его в папку с проектом. Версию браузера можно посмотреть [вот так](https://www.bolshoyvopros.ru/questions/2494554-kak-uznat-kakaja-u-menja-stoit-versija-google-chrome.html) Сейчас скачана версия для браузера версии 123.0.6312.... Скачать драйвер для вашей версии можно [вот тут](https://googlechromelabs.github.io/chrome-for-testing/) или для старых версить [тут](https://chromedriver.chromium.org/downloads)
2. Открыть командную строку windows и перейти в папку с проектом
3. Ввести команды
``` shell
python -m venv venv
venv\Scripts\activate.bat
pip install -r requirements.txt
python main.py
``` 

