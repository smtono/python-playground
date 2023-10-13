"""
this is one of the modules of all time
"""

class Module1:
    """module1"""
    
    def __init__(self, args) -> None:
        """
        module1 this is a bunch of args and the next 3 args are set attributes 1 2 3
        """
        self.text = args[0:-3]
        self.setattr1 = args[-1]
        self.setattr2 = args[-2]
        self.settattr3 = args[-3]    

    def execute(self):
        """execut
        """
        print(self.text)
        print("babababababbasdjgkwhagsduijkvhbwauigjdksv,")
        print(f"{self.setattr1}, {self.setattr2}, {self.settattr3}")
    