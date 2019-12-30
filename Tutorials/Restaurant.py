class restaurant:
    def __init__(self,name,cuisine_type):
        self.name = name
        self.cuisine_type = cuisine_type
        self.number_served = 0

    def describe_restaurant(self):
        print(f"{self.name.title()} serves {self.cuisine_type.title()} style food.")

    def open_restaurant(self):
        print(f"{self.name.title()} is now open!")

    def set_number_served(self, num_served):
        self.number_served = num_served

    def increment_number_served(self,served):
        self.number_served += served
    
        
bobs = restaurant("bobs burgers", "american")
bobs.describe_restaurant()
bobs.open_restaurant()
print(bobs.number_served)
bobs.set_number_served(8)
print(bobs.number_served)
bobs.increment_number_served(2)
print(bobs.number_served)