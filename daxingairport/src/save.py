import os

def resout():
    print(__name__)

if __name__ == "__main__":
    print(os.getcwd())
    print(os.path.dirname(__file__))
