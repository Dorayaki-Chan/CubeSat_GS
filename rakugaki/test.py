class Test:
    def __init__(self):
        print(self.abc("asdf"))
    
    def abc(self, a):
        return "hello!" + a

if __name__ == '__main__':
    y = Test()