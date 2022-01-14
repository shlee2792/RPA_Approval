import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from dateutil import relativedelta

# input_date = datetime.today().strftime('%Y%m%d')


# titles = []
# links = []
# res = requests.get("https://news.nate.com/rank/interest?sc=sisa&p=day&date={}".format(input_date))

# res.raise_for_status()
# res.encoding = None
# html = res.text

# soup = BeautifulSoup(html, 'html.parser')
# # ranking_list = soup.find('div', {'class': 'postRankSubjectList f_clear'}).findAll('dt')
# test_list = soup.select('#newsContents > div > div.postRankSubjectList.f_clear ')

# test_list_2 = soup.select('#postRankSubject')

# for test in test_list:
#     print(test)
    

