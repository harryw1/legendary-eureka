import numpy as np

def test():
    np.random.seed(0)
    return np.random.rand(3, 3)

print(test())
