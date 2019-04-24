import pandas as pd
import requests as requests
import urllib
from bs4 import BeautifulSoup
from geopy.geocoders import Nominatim

data_loc = 'largestcompanies.xlsx'
data = pd.read_excel(data_loc)
companies = list(data['Company Name'])
userAgent = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
search = "https://www.google.com/search?q="
geolocator = Nominatim(user_agent='kirk_kenney')


rank = 0


for company in companies:
    rank += 1
    hq_search = BeautifulSoup(requests.get(search + urllib.parse.quote_plus(company) + '+headquarters', headers=userAgent).text, 'html.parser')
    hq = hq_search.find('div', {'class': 'Z0LcW'}).get_text()
    location = geolocator.geocode(hq)
    lat = location.latitude
    lon = location.longitude
    # data.loc[company, 'Headquarters'] = hq
    data.loc[company, 'Latitude'] = lat
    data.loc[company, 'Longitude'] = lon
    data.loc[company, 'Rank'] = rank
    data.to_excel(data_loc, index=False)
    print(company + ' added')
    # coord_search = BeautifulSoup(requests.get(search + urllib.parse.quote_plus(hq) + '+coordinates', headers=userAgent).text, 'html.parser')
    # coord = coord_search.find('div', {'class': 'Z0LcW'}).get_text()
    # lat, lon = coord.split(', ')
    # print(company + ': ' + 'LAT ' + lat + ', LON ' + lon)

# import pandas as pd
# import requests as requests
# import urllib
# from bs4 import BeautifulSoup
# datasetLocation = "largestcompanies.xlsx"
# df = pd.read_excel(datasetLocation, "Sheet1")
# userAgent = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}
# searchURL = "https://www.google.com/search?q="
# for x in range(0, len(df.index)):
#     companyName = df.iloc[x, 0]
#     search1 = BeautifulSoup(requests.get(searchURL + urllib.parse.quote_plus(companyName) + '+headquarters ', headers=userAgent).text, 'html.parser')
#     headquarters = search1.find('div', {"class": "Z0LcW"}).get_text()
#     if not headquarters:
#         headquarters = search1.find('div', {"class": "desktop-title-subcontent"}).get_text()
#     search2 = BeautifulSoup(requests.get(searchURL + urllib.parse.quote_plus(headquarters) + '+coordinates', headers=userAgent).text, 'html.parser')
#     coordinates = search2.find('div', {"class": "Z0LcW"}).get_text()
#     df.loc[x, 'Location'] = headquarters
#     df.loc[x, 'Coordinates'] = coordinates
#     print(df.iloc[x, 0] + " | " + headquarters + " | " + coordinates)
# df.to_excel(datasetLocation, index=False)
