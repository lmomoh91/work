class Car:
    """A simple attempt to represent a car."""
    def __init__(self, make, model, year):
        self.make = make 
        self.model = model 
        self.year = year
        self.odometer_reading = 0 
        self.amount_of_gas = 50 

    def get_description_name(self):
        long_name = f"{self.year} {self.make} {self.model}"
        return long_name.title()
    
    def read_odometer(self):
        reading = f"{self.get_description_name()} has {self.odometer_reading} miles"
        return reading

    def update_odometer(self, mileage):
        """
        Set the odometer reading to a given value.
        Reject if someone attempts to roll back the odometer
        """
        if mileage >= self.odometer_reading:
            self.odometer_reading = mileage
        else:
            print(f"You cant roll back the odometer")
    
    def increment_odometer(self, miles):
        """Add the given amount to the odometer reading."""
        self.odometer_reading += miles

    def fill_gas_tank(self):
        self.amount_of_gas = 100

class Battery:
    

    """
    A simple attempt to model a car battery for an electric car.
    """
    def __init__(self, battery_size = 75):
        self.battery_size = battery_size

    def describe_battery(self):
        print(f"This car has a {self.battery_size}--kWh battery.")
        

class ElectricCar(Car):
    """
    Same Aspects as car, but specifically for Electric Cars
    """

    def __init__(self, make, model, year):
        """Initialize attributes of the parent class."""
        super().__init__(make, model, year)
        self.battery = Battery()
    
    def fill_gas_tank(self):
        print(f"This is an Electric Car. There is no gas tank!")
    
    
my_tesla = ElectricCar('tesla', 'model s', 2019)

print(my_tesla.get_description_name())

my_tesla.battery.describe_battery()