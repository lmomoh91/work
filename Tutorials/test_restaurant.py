from Restaurant import *

bobs = restaurant("bobs burgers", "american")
bobs.describe_restaurant()
bobs.open_restaurant()
print(bobs.number_served)
bobs.set_number_served(8)
print(bobs.number_served)
bobs.increment_number_served(2)
print(bobs.number_served)
bobs.reset_number_served()

coldstone = IceCreamStand('ColdStone','Dessert')
coldstone.ListFlavors()