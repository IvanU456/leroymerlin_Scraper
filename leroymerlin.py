import random
import requests
from bs4 import BeautifulSoup
import openpyxl
import datetime


proxies_list = [
    '193.23.253.115:7687',
    '193.23.253.217:7789',
    '193.23.253.103:7675',
    '193.23.253.127:7699',
    '193.23.253.93:7665',
    '193.23.253.43:7615',
    '193.23.253.212:7784',
    '193.23.253.242:7814',
    '193.23.253.239:7811',
    '193.23.253.226:7798'
]

now = datetime.datetime.now()
t = now.strftime("%d-%m-%Y %H.%M")
URL = input('Введите ссылку на категорию: ').strip()
print('Парсинг ссылок на товары')
HEADERS = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36', 'accept': '*/*'}
FILE = f'товары {t}.xlsx'


def get_html(url, params=None):
    proxies = {
        'https': f'http://puwympav:61xs8f4j4vcr@{random.choice(proxies_list)}'
    }
    r = requests.get(url, headers=HEADERS, params=params, proxies=proxies)
    return r


def get_pages_count(html):
    soup = BeautifulSoup(html, 'html.parser')
    pagination = soup.find('div', class_='s1pmiv2e_plp').find_all_next('a', class_='l7pdtbg_plp')
    if pagination:
        return int(pagination[-2].get_text())
    else:
        return 1


def get_links(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_='phytpj4_plp largeCard')
    links = []
    for item in items:
        link = item.find('a').get('href')
        links.append('https://leroymerlin.ru/'+link)
    return links


def get_content(link, html):
    soup = BeautifulSoup(html, 'html.parser')
    card = []
    name = soup.find('div', class_='product-detailed-page').get('data-product-name')
    price = soup.find(itemprop="price").get("content").replace('.', ',')
    rating = soup.find(itemprop="ratingValue").get("content")
    avaliable = soup.find('div', class_='product-detailed-page').get('data-product-is-available')
    def_list = soup.find('uc-pdp-section-layout').find_next('p').get_text(strip=True)
    if avaliable == 'available':
        avaliable = 'В наличие'
    else:
        avaliable = 'Нет в наличие'
    if def_list == '':
        def_list = soup.find('uc-pdp-section-layout').find_next('p').find_next('p').get_text(strip=True)
    if def_list == 'СКАЧАТЬ ИНСТРУКЦИЮ':
        def_list = soup.find('uc-pdp-section-layout').find_next('p').find_next('p').find_next('p').get_text(strip=True)
    if def_list == 'Цены и наличие товаров на сайте и в гипермаркетах могут различаться. Пожалуйста, уточняйте стоимость и наличие товаров в конкретном магазине.':
        def_list = 'Нет описания'
    card.append({
        'link': link,
        'price': price,
        'name': name,
        'rating': rating,
        'avaliable': avaliable,
        'def_list': def_list
    })
    return card


def save_file(items, path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['A1'] = 'Ссылка'
    sheet['B1'] = 'Цена'
    sheet['C1'] = 'Название'
    sheet['D1'] = 'Рейтинг'
    sheet['E1'] = 'Наличие'
    sheet['F1'] = 'Характеристики'
    row = 2
    for item in items:
        sheet[row][0].value = item['link']
        sheet[row][1].value = item['price']
        sheet[row][2].value = item['name']
        sheet[row][3].value = item['rating']
        sheet[row][4].value = item['avaliable']
        sheet[row][5].value = item['def_list']
        row += 1
    wb.save(path)


def main():
    html = get_html(URL)
    if html.status_code == 200:
        links = []
        crads = []
        pages_count = get_pages_count(html.text)
        for page in range(1, pages_count + 1):
            html = get_html(URL, params={'page': page})
            links.extend(get_links(html.text))
        print(f'Получено {len(links)} товаров')
        i = 0
        try:
            for link in links:
                i += 1
                print(f"Парсинг {i} из {len(links)}")
                html = get_html(link)
                crads.extend(get_content(link, html.text))
            save_file(crads, FILE)
        except AttributeError:
            for link in links[i:]:
                i += 1
                print(f"Парсинг {i} из {len(links)}")
                html = get_html(link)
                crads.extend(get_content(link, html.text))
            save_file(crads, FILE)
    else:
        print("Error")


if __name__ == '__main__':
    main()
