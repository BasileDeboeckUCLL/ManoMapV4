import numpy as np
import matplotlib.pyplot as plt

def plot_topographic_data():
    # Dummy data for example purposes
    x = np.linspace(-5, 5, 100)
    y = np.linspace(-5, 5, 100)
    X, Y = np.meshgrid(x, y)
    Z = np.sin(np.sqrt(X**2 + Y**2))

    plt.figure()
    cp = plt.contourf(X, Y, Z)
    plt.colorbar(cp)
    plt.title('Topographic Contour Plot')
    plt.xlabel('X axis')
    plt.ylabel('Y axis')
    plt.show()