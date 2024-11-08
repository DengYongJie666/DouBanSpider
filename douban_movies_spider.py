import pandas as pd
import requests
import random
import time
from bs4 import BeautifulSoup


header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/92.0.4515.131 Safari/537.36 SLBrowser/8.0.0.12022 SLBChan/105 '
    }

sess = requests.session()


def get_movie_info(url):
        try:
            response = sess.get(url, headers=header, timeout=5)
            source_code = response.status_code
            if source_code == 200:
                soup = BeautifulSoup(response.text,'html.parser')
            else:
                print(f"请求失败，状态码: {source_code}")
                return None
        except Exception as e:
            print(e)
            return None

        year_span = soup.find('span',{'class': 'year'})
        if year_span:
            year = year_span.text.strip('()')
        else:
            year = '未知'

        director_span = soup.find('span',{'class': 'pl'},string='导演')
        if director_span:
            director = director_span.find_next_sibling('span',{'class': 'attrs'}).text.strip()
        else:
            director = '未知'

        screenwriter_span = soup.find('span',{'class': 'pl'},string='编剧')
        if screenwriter_span:
            screenwriter = screenwriter_span.find_next_sibling('span',{'class': 'attrs'}).text.strip()
        else:
            screenwriter = '未知'

        actors_span = soup.find('span',{'class': 'pl'},string='主演')
        if actors_span:
            actors = actors_span.find_next_sibling('span',{'class': 'attrs'}).text.strip()
        else:
            actors = '未知'

        genres_span = soup.find('span',{'class': 'pl'},string='类型:')
        if genres_span:
            genres = ', '.join([span.text for span in genres_span.find_next_siblings('span',{'property': 'v:genre'})])
        else:
            genres = '未知'

        country_span = soup.find('span',{'class': 'pl'},string='制片国家/地区:')
        if country_span:
            country = country_span.next_sibling
        else:
            country = '未知'

        return {
            'year': year,
            'director': director,
            'screenwriter': screenwriter,
            'actors': actors,
            'genres': genres,
            'country': country
        }


def get_movie_info_from_top_250():
    page_list = []
    name_title = []
    website_dict = {}
    for i in range(0,250,25):
        url = f"https://movie.douban.com/top250?start={i}&filter="
        time.sleep(random.random() * 5)
        try:
            response = sess.get(url, headers=header, timeout=5)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text,'html.parser')
                # 获取电影名称和链接
                movie_items = soup.find_all('div',{'class': 'item'})
                for item in movie_items:
                    title = item.find('span',{'class': 'title'}).text.strip()
                    link = item.find('a')['href']
                    name_title.append(title)
                    website_dict[title] = {'link': link}
                page_list.append(soup)
            else:
                print(f"请求失败，状态码: {response.status_code}")
        except Exception as e:
            print(e)

        print(f"下载分页:{url}")
    print(f"共下载了 {len(page_list)} 页")
    name_title = [title for title in name_title if title]
    total_movies = len(name_title)
    print(f"共需处理 {total_movies} 个电影")

    for idx, (title, info) in enumerate(website_dict.items(), start=1):
        link = info['link']
        time.sleep(random.random() * 5)
        movie_info = get_movie_info(link)
        if movie_info:
            website_dict[title].update(movie_info)
            print(f"正在处理第 {idx} 个电影: {title} ({idx}/{total_movies})")
        else:
            print(f"处理第 {idx} 个电影失败: {title} ({idx}/{total_movies})")

    ordered_website_list = [website_dict[title] for title in name_title]
    print(name_title)
    # print(len(name_title))
    print(ordered_website_list)
    # print(len(ordered_website_list))

    return name_title,ordered_website_list


def save_to_excel(name_title, ordered_movie_info):
    data = {
        '电影名称': name_title,
        '网站链接': [info['link'] for info in ordered_movie_info],
        '年份': [info['year'] for info in ordered_movie_info],
        '导演': [info['director'] for info in ordered_movie_info],
        '编剧': [info['screenwriter'] for info in ordered_movie_info],
        '主演': [info['actors'] for info in ordered_movie_info],
        '类型': [info['genres'] for info in ordered_movie_info],
        '国家': [info['country'] for info in ordered_movie_info]
    }

    df = pd.DataFrame(data)
    df.to_excel('movies_top250.xlsx', index=False)
    print("数据已写入Excel文件: movies_top250.xlsx")


if __name__ == '__main__':
    name_title, ordered_movie_info = get_movie_info_from_top_250()
    save_to_excel(name_title, ordered_movie_info)




