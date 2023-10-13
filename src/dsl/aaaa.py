"""
dsl prototype
"""

from dsl.modules.module1 import Module1


class Parser:
    """
    this parses stuff
    """
    def __init__(self, stupid: [str]) -> None:
        for blep in stupid:
            beep = blep.split(" ")
            self.command = beep[0]
            self.args = beep[1:]

    def execute(self):
        #importlib
        if self.command == "module1":
            Module1(self.args).execute()

if __name__ == "__main__":
    test = "module1 blah blah blah blah blah 1 2 3"
    Parser([test]).execute()
       