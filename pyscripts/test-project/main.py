import matplotlib.pyplot as plt
import numpy
import pandas

# Create and test pandas, numpy, plt
df = pandas.DataFrame(numpy.random.randn(10, 2), columns=['A', 'B'])
df['C'] = df['A'] + df['B']
df['D'] = df['A'] - df['B']
df['E'] = df['A'] * df['B']
df['F'] = df['A'] / df['B']
print(df)

plt.plot(df['A'], df['B'])
plt.show()
