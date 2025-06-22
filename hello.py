def greet(name):
    return f"Hello, {name}!"

def add_numbers(a, b):
    return a + b

if __name__ == "__main__":
    print(greet("World"))
    result = add_numbers(5, 3)
    print(f"5 + 3 = {result}")