from random import randint

class Die:
    def __init__ (self,sides=6):
        self.sides = int(sides)
    pass

    def roll_die(self, number_of_rolls):
        for i in range(int(number_of_rolls)):
            print(randint(1,self.sides))

        print('Done')
        pass



Die(6).roll_die(5)

Die(10).roll_die(10)

Die(20).roll_die(10)