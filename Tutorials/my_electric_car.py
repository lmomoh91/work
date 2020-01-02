from Car import car

class Battery:
    """
    A simple attempt to model a car battery for an electric car.
    """
    def __init__(self, battery_size = 75):
        self.battery_size = battery_size

    def describe_battery(self):
        print(f"This car has a {self.battery_size}--kWh battery.")

    def get_range(self):
        """Print a statement about the range this battery provides."""
        if self.battery_size == 75:
            range = 260     
        elif self.battery_size == 100:
            range = 315
        print(f"This car can go about {range} miles on a full charge")

    def upgrade_battery(self):
        if self.battery_size != 100:
            self.battery_size = 100
            print(f"Battery Upgraded!!!")
        else:
            pass


class ElectricCar(car):
    """
    Same Aspects as car, but specifically for Electric Cars
    """

    def __init__(self, make, model, year):
        """Initialize attributes of the parent class."""
        super().__init__(make, model, year)
        self.battery = Battery()
    
    def fill_gas_tank(self):
        print(f"This is an Electric Car. There is no gas tank!")