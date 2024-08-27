import math


def get_specific_letter_by_increment(increment_to_A):
    length = ord('Z') - ord('A') + 1
    times = math.floor(increment_to_A / length)
    mod = increment_to_A % length
    ending = chr(ord('A') + mod)
    starting = 'A' * times
    return starting + ending