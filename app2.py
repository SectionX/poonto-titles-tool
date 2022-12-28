#decorator test
from time import perf_counter

def dec(func):
    print("created function: " + repr(func) )
    def wrapper(*args):
        print("Inside Wrapper")
        return func(*args)
    return wrapper



@dec
def my_sum(n):
    if n == 0:
        return 0
    return n + my_sum(n-1)


print(my_sum(4))