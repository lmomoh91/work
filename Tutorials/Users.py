class User:
    def __init__(self, FirstName, LastName, Age, Sex):
        self.FirstName = FirstName
        self.LastName = LastName
        self.Age = Age
        self.Sex = Sex
        self.login_attempts = 0


    def describe_user(self):
        description = f"{self.FirstName} {self.LastName} is a {self.Age} year old {self.Sex}."
        return description

    def greet_user(self):
        print(f"Hello {self.FirstName}!")

    def increment_login_attempts(self):
        self.login_attempts += 1

    def reset_login_attempts(self):
        self.login_attempts = 0




class privileges:

    def __init__(self):
        self.privilages = ["can add post", "can delete posts", "can ban user"]
        pass

    def show_privilages(self):
        print(self.privilages)


class Admin(User):

    def __init__(self, FirstName, LastName, Age, Sex):
        super().__init__(FirstName,LastName,Age,Sex)
        self.privilages = privileges()
        



user0 = User('Doug','Funny','17','Male')
print(user0.describe_user())
user0.increment_login_attempts()
user0.increment_login_attempts()
print(user0.login_attempts)
user0.reset_login_attempts()
print(user0.login_attempts)
user1 = Admin("Cousin", "Skeeter", '25', 'Male')
user1.privilages.show_privilages()