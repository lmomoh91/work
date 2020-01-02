"""Set of classes that represent gas and electric cars."""

class car:
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
        reading = f"{self.get_description_name()} has {self.odometer_reading} miles on it."
        return print(reading)

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



    