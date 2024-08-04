import scrapy
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from scrapy.http import HtmlResponse
from openpyxl import Workbook
import logging
import time
import json
from typing import List, Dict


class TnvedSpider(scrapy.Spider):
    name = 'tnved_spider'
    allowed_domains = ['tnved.info']
    start_urls = ['https://tnved.info/TnvedTree/']
    base_url = 'https://tnved.info/TnvedTree/'

    def __init__(self, *args, **kwargs) -> None:
        """
        Инициализация паука.
        """
        super(TnvedSpider, self).__init__(*args, **kwargs)
        try:
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            logging.info("Webdriver успешно инициализирован.")
        except Exception as e:
            logging.error(f"Ошибка при инициализации webdriver: {e}")

    def parse(self, response: HtmlResponse) -> None:
        """
        Метод для парсинга страницы. Разворачивает элементы дерева и собирает данные.

        Args:
            response (HtmlResponse): Ответ на запрос страницы.
        """
        try:
            self.driver.get(response.url)
            logging.info(f"Открытие страницы: {response.url}")

            # Разворачиваем все элементы дерева на первом уровне
            first_level_buttons = self.driver.find_elements(By.CSS_SELECTOR, '.css-1ganlhb')
            for button in first_level_buttons:
                self.driver.execute_script("arguments[0].click();", button)
                time.sleep(1)

            # Разворачиваем все товарные группы до товарных позиций
            second_level_buttons = self.driver.find_elements(By.CSS_SELECTOR, 'ul.MuiBox-root.css-optkkl .css-1ganlhb')
            for button in second_level_buttons:
                self.driver.execute_script("arguments[0].click();", button)
                time.sleep(1)

            # HTML-код страницы
            body = self.driver.page_source
            response = HtmlResponse(url=response.url, body=body, encoding='utf-8')
            logging.info("HTML-код страницы получен.")

            data: List[Dict[str, str]] = []

            sections = response.css('li.MuiBox-root')
            logging.info(f"Найдено {len(sections)} секций.")

            for section in sections:
                section_title = section.css('div.css-1ngz1cb::text').get().strip() if section.css(
                    'div.css-1ngz1cb::text') else "N/A"
                section_description = section.css('div.css-1ia32ta::text').get().strip() if section.css(
                    'div.css-1ia32ta::text') else "N/A"
                logging.info(f"Секция: {section_title} - {section_description}")

                groups = section.css('ul.MuiBox-root.css-optkkl > li.MuiBox-root.css-1y3glo5')
                logging.info(f"Секция: {section_title} - Найдено {len(groups)} групп")

                for group in groups:
                    group_html = group.get()
                    logging.info(f"Group HTML: {group_html}")

                    # CSS-селектор
                    group_title_css = group.css('div.css-nycptu > div.css-mya6i8::text').get().strip() if group.css(
                        'div.css-nycptu > div.css-mya6i8::text') else "N/A"
                    logging.info(f"Группа (CSS): {group_title_css}")

                    # XPath-селектор
                    group_title_xpath = group.xpath('.//div[contains(@class, "css-nycptu")]/div[contains(@class, "css-mya6i8")]/text()').get().strip() if group.xpath(
                        './/div[contains(@class, "css-nycptu")]/div[contains(@class, "css-mya6i8")]/text()') else "N/A"
                    logging.info(f"Группа (XPath): {group_title_xpath}")

                    group_title = group_title_css if group_title_css != "N/A" else group_title_xpath
                    group_description = group.css('div.css-1ia32ta::text').get().strip() if group.css(
                        'div.css-1ia32ta::text') else "N/A"
                    logging.info(f"Группа: {group_title} - {group_description}")

                    items = group.css('ul.MuiBox-root.css-optkkl > li.MuiBox-root.css-1y3glo5')
                    logging.info(f"Группа: {group_title} - Найдено {len(items)} элементов")

                    for item in items:
                        self.process_item(item, data, section_title, section_description, group_title, group_description)

            self.save_to_xlsx(data)

            with open('tnved_data.json', 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)

        except Exception as e:
            logging.error(f"Ошибка при парсинге страницы: {e}")
        finally:
            self.driver.quit()
            logging.info("Webdriver успешно завершил работу.")

    def process_item(self, item, data, section_title, section_description, group_title, group_description):
        """
        Обрабатывает элемент, включая вложенные элементы.

        Args:
            item: текущий элемент для обработки.
            data: список для хранения данных.
            section_title: название секции.
            section_description: описание секции.
            group_title: название группы.
            group_description: описание группы.
        """
        item_html = item.get()
        logging.info(f"Item HTML: {item_html}")

        item_title_css = item.css('div.css-nycptu > div.css-mya6i8::text').get().strip() if item.css(
            'div.css-nycptu > div.css-mya6i8::text') else "N/A"
        logging.info(f"Элемент (CSS): {item_title_css}")

        item_title_xpath = item.xpath('.//div[contains(@class, "css-nycptu")]/div[contains(@class, "css-mya6i8")]/text()').get().strip() if item.xpath(
            './/div[contains(@class, "css-nycptu")]/div[contains(@class, "css-mya6i8")]/text()') else "N/A"
        logging.info(f"Элемент (XPath): {item_title_xpath}")

        item_title_alt_css = item.css('div.css-nycptu > div.css-1gn4is1::text').get().strip() if item.css(
            'div.css-nycptu > div.css-1gn4is1::text') else "N/A"
        logging.info(f"Элемент (CSS Alt): {item_title_alt_css}")

        item_title = item_title_css if item_title_css != "N/A" else item_title_xpath if item_title_xpath != "N/A" else item_title_alt_css

        if item_title.isdigit() and len(item_title) > 4:
            item_title = item_title[:4]

        item_description = item.css('div.css-1ia32ta::text').get().strip() if item.css(
            'div.css-1ia32ta::text') else "N/A"
        logging.info(f"Элемент: {item_title} - {item_description}")

        notes_link = item.css('div[aria-label="Примечания"] > a::attr(href)').get()
        if notes_link:
            notes_link = self.base_url + notes_link
        explanation_link = item.css('div[aria-label="Пояснения"] > a::attr(href)').get()
        if explanation_link:
            explanation_link = self.base_url + explanation_link
        logging.info(f"Примечания: {notes_link}, Пояснения: {explanation_link}")

        data.append({
            'Раздел': section_title,
            'Описание раздела': section_description,
            'Товарная группа': group_title,
            'Описание товарной группы': group_description,
            'Товарная позиция': item_title,
            'Описание товарной позиции': item_description,
            'Примечания': notes_link,
            'Пояснения': explanation_link
        })

        # Рекурсивно обрабатываем вложенные элементы
        nested_items = item.css('ul.MuiBox-root.css-optkkl > li.MuiBox-root.css-1y3glo5')
        logging.info(f"Найдено {len(nested_items)} вложенных элементов для {item_title}")

        for nested_item in nested_items:
            self.process_item(nested_item, data, section_title, section_description, group_title, group_description)

    def save_to_xlsx(self, data: List[Dict[str, str]]) -> None:
        """
        Сохраняет собранные данные в файл Excel и выравнивает ширину колонок.

        Args:
            data (List[Dict[str, str]]): Список данных для сохранения в Excel файл.
        """
        wb = Workbook()
        ws = wb.active
        ws.append(['Раздел', 'Описание раздела', 'Товарная группа', 'Описание товарной группы', 'Товарная позиция',
                   'Описание товарной позиции'])

        for item in data:
            ws.append(
                [item['Раздел'], item['Описание раздела'], item['Товарная группа'], item['Описание товарной группы'],
                 item['Товарная позиция'], item['Описание товарной позиции']])

        # Выравнивание ширины колонок по содержимому
        for column in ws.columns:
            max_length = 0
            column_name = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_name].width = adjusted_width

        wb.save('tnved_data.xlsx')
        logging.info("Данные успешно сохранены в tnved_data.xlsx")

    def close(self, reason: str) -> None:
        """
        Закрывает драйвер при завершении работы паука.

        Args:
            reason (str): Причина закрытия.
        """
        try:
            self.driver.quit()
            logging.info("Webdriver успешно закрыт.")
        except Exception as e:
            logging.error(f"Ошибка при закрытии webdriver: {e}")
