class Hello:
    def __init__(self, name: str, age: int):
        self.name = name
        self.age = age

    def hello(self) -> str:
        hw = f"Hello! {self.name}, your age is {self.age}"
        return hw
