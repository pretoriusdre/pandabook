

class Hyperlink():
    def __init__(self, url, name):
        self.url = url
        self.name = name
        
    def __repr__(self):
        return f"Hyperlink(name='{self.name}', url='{self.url}')"