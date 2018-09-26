# -*- coding: utf-8 -*-
from __future__ import division
from numpy.random import randn
import numpy as np

# -*- coding: utf-8 -*-
###通用函数
arr = np.arange(10)
print(np.sqrt(arr))
print(np.exp(arr))

x = randn(8)
y = randn(8)
print(x)
print(y)
np.maximum(x, y)  # 元素级最大值

arr = randn(7) * 5
print(arr)

np.modf(arr)

###利用数组进行数据处理
# 向量化
points = np.arange(-5, 5, 0.01)  # 1000 equally spaced points
xs, ys = np.meshgrid(points, points)
ys
