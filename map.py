import folium
from folium.features import DivIcon
import pandas as pd


data = pd.read_excel('largestcompanies.xlsx')
company = list(data['Company Name'])
values = list(data['Market value (USD billions)'])
location = list(data['Headquarters'])
latitude = list(data['Latitude'])
longitude = list(data['Longitude'])
rank = list(data['Rank'])
world_gdp = 88000.00
totals = []
total_value = 0

html = '''<div style="font-size:20px; text-align:center; font-weight:bold; margin:5px 0">%s</div> <br>
<span style="font-size:15px">HQ: %s</span> <br>
<span style="font-size:15px">Rank: %s</span> <br>
<span style="font-size:15px">Market value: $%s billion</span> <br>
<span style="font-size:15px">Proportion Of World GDP: %s</span>
'''

legend_html = '''<div style="position:fixed; border:solid 1px black;
background-color:rgba(255,255,255,0.8); bottom:50px; left:50px; z-index:999;
padding: 15px 10px">
<div style="font-size: 22px">Total Market Value: $%sbillion</div>
<div style="font-size: 22px"> Proportion Of World GDP: %s</div>
</div>"
'''


map = folium.Map(location=[0, 0], zoom_start=2)
fg_markers = folium.FeatureGroup(name='World\'s Largest Companies')
for company, value, loc, lat, lon, rank in zip(company, values, location, latitude, longitude, rank):
    totals.append(value)
    percent = value/world_gdp
    percent = '{:.2%}'.format(percent)
    iframe = folium.IFrame(html=html % (company, loc, rank, value, percent), width=300, height=150)
    fg_markers.add_child(folium.CircleMarker(location=[lat, lon],
        popup=folium.Popup(iframe), fill=True,
        color='black', fill_color='blue',
        fill_opacity=0.6, radius=12, tooltip='Rank #' + str(rank)))

for total in totals:
    total_value += total

proportion = total_value/world_gdp
proportion = '{:.2%}'.format(proportion)
print(total_value)
print(proportion)

map.add_child(fg_markers)
map.get_root().html.add_child(folium.Element(legend_html % (total_value, proportion)))
map.save('index.html')
