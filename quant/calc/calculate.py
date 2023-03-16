class Calculate:
    def __init__(self, numlst: list = [1, 2, 3]):
        self.numlist = numlst

    def add(self) -> float:
        summation = 0
        for ea_num in self.numlist:
            summation += ea_num
        return summation

    def multiply(self) -> float:
        product = 1
        for ea_num in self.numlist:
            product *= ea_num
        return product
