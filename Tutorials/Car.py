class Car:
    """A simple attempt to represent a car."""
    def __init__(self, make, model, year):
        self.make = make 
        self.model = model 
        self.year = year
        self.odometer = 0 

    def get_description_name(self):
        long_name = f"{self.year} {self.make} {self.model}"
        return long_name.title()
    
    def odometer_reading(self):
        reading = f"{self.get_description_name()} has {self.odometer} miles"
        return reading



    
    
    
my_new_car = Car('Jaguar', 'F-Type', '2019')
print(my_new_car.get_description_name())
print(my_new_car.odometer_reading())