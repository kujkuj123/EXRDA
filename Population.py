from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import xlwings
geolocator = Nominatim(user_agent="geoapiExercises")

TU_name_list=['TU Delft','TU Eindhoven','University of Twente', 'Wageningen University']

TU_coordinates=[]

for i in range(4):
    TU_coordinates.append([geolocator.geocode(TU_name_list[i]).latitude,geolocator.geocode(TU_name_list[i]).longitude])

def distance(start,end):
    return geodesic(start, end).kilometers

tot_population = 17.28*10**6
dutch_population  = 48428

workbook = xlwings.Book('C:\\Users\\danie\\Desktop\\Pop\\Population_offline.xlsx')
sht = workbook.sheets('First')
city_name_list = sht.range('A1:A352').value
population_city_list = sht.range('C1:C352').value
temp1 = sht.range('D1:D352').value
temp2 = sht.range('E1:E352').value

population_student_list=[]

for i in range(len(population_city_list)):
    population_student_list.append(round((population_city_list[i]/tot_population)*dutch_population))

city_coordinate_list=[]

for i in range(len(city_name_list)):
    city_coordinate_list.append([temp1[i],temp2[i]])

distance_list=[]

for i in range(len(city_name_list)):
    distance_list.append(
        [distance(city_coordinate_list[i], TU_coordinates[0]), distance(city_coordinate_list[i], TU_coordinates[1]),
         distance(city_coordinate_list[i], TU_coordinates[2]), distance(city_coordinate_list[i], TU_coordinates[3])])

Delft_dutch_students=0
Eindhoven_dutch_students=0
Twente_dutch_students=0
Wageningen_dutch_students=0

for i in range(len(city_name_list)):
    index_min = min(range(len(distance_list[i])), key=distance_list[i].__getitem__)
    if index_min==0:
        Delft_dutch_students += population_student_list[i]
    elif index_min==1:
        Eindhoven_dutch_students += population_student_list[i]
    elif index_min==2:
        Twente_dutch_students += population_student_list[i]
    elif index_min==3:
        Wageningen_dutch_students += population_student_list[i]

print('Delft:',Delft_dutch_students)
print('Ein: ',Eindhoven_dutch_students)
print('Twente: ',Twente_dutch_students)
print('Wagen: ',Wageningen_dutch_students)



