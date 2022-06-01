import math

i = 0
while True:

    power = pow(i, i)
    fact = math.factorial(i)
    
    if power > fact:
        print(i, 'Power bigger!')
    else:
        print(i, 'Factorial bigger!')

    i += 1