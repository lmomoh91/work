class restaurant:
    def __init__(self,name,cuisine_type):
        self.name = name
        self.cuisine_type = cuisine_type
        self.number_served = 0

    def describe_restaurant(self):
        print(f"{self.name.title()} serves {self.cuisine_type.title()} style food.")

    def open_restaurant(self):
        print(f"{self.name.title()} is now open!")

    def reset_number_served(self):
        self.number_served = 0

    def set_number_served(self, num_served):
        self.number_served = num_served

    def increment_number_served(self,served):
        self.number_served += served

class IceCreamStand(restaurant):

    def __init__(self,name,cuisine_type):
        super().__init__(name,cuisine_type)
        self.flavors = ['Vanilla', 'Strawberry', 'Chocolate']
        pass

    def ListFlavors(self):
        print(self.flavors)
        pass
    
        
