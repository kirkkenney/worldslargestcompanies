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
    # Google HQ for each company in data set
    hq_search = BeautifulSoup(requests.get(search + urllib.parse.quote_plus(company) + '+headquarters', headers=userAgent).text, 'html.parser')
    # get values from 1st result displayed in Google search
    hq = hq_search.find('div', {'class': 'Z0LcW'}).get_text()
    # get latitutde and longitude of city where company HQ is located
    location = geolocator.geocode(hq)
    lat = location.latitude
    lon = location.longitude
    # append latitude to Excel sheet
    data.loc[company, 'Latitude'] = lat
    # append longitude to Excel sheet
    data.loc[company, 'Longitude'] = lon
    # append rank integer to Excel sheet
    data.loc[company, 'Rank'] = rank
    # save amended Excel sheet
    data.to_excel(data_loc, index=False)
