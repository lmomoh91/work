from Car import car
from my_electric_car import ElectricCar

my_bettle = car('volkswagen', 'beetle', 2019)
print(my_bettle.get_description_name())

my_tesla = ElectricCar('tesla', 'roadster', 2019)
print(my_tesla.get_description_name())
