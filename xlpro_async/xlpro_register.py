
import time
import matplotlib.pyplot as plt
import numpy as np

def dummy(a:float, b:float):
    time.sleep(3)
    return a * b

def create_figure():
    fig, ax = plt.subplots()
    
    xdata = np.linspace(0, 10, 101)
    ydata = np.sin(xdata)

    ax.plot(xdata, ydata)

    return fig

def add_numbers(a:float, b:float):
    return a + b


if __name__ == "__main__":
    fig = create_figure()
    plt.show()