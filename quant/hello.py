class Hello:
    def __init__(self, name: str = "Chris", age: int = 40):
        self.name = name
        self.age = age

    def say_hello(self) -> str:
        hw = f"Hello! {self.name}, your age is {self.age}"
        return hw
