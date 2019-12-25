class restaurant:
    def __init__(self,name,cuisine_type):
        self.name = name
        self.cuisine_type = cuisine_type

    def describe_restaurant(self):
        print(f"{self.name.title()} serves {self.cuisine_type.title()} style food.")

    def open_restaurant(self):
        print(f"{self.name.title()} is now open!")
        
        
bobs = restaurant("bobs burgers", "american")
bobs.describe_restaurant()
bobs.open_restaurant()

luigis = restaurant(f"Luigi's", "Italian")
luigis.describe_restaurant()
luigis.open_restaurant()


Leo = restaurant(f"Leo", "African")
Leo.describe_restaurant()
Leo.open_restaurant()