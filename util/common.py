#!/usr/bin/env python3
# -*- coding: utf-8 -*-
' common model '
__author__ = 'Yuechen Yang'

from functools import reduce
import types
import math

def char2num(s):
    return {'0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9}[s]

def str2int(s):
    print(map(char2num, s))
    return reduce(lambda x, y: x * 10 + y, map(char2num, s))


def my_round(number, accuracy_level = 2):
    if isinstance(number, float):
    #if number.isdigit():
        if(math.isnan(number)):
            return .0
        # 设置精度取两位小数
        return round(number + 0.0001, accuracy_level)
    else:
        return 0


if __name__ == '__main__':
    print(str2int('157'))