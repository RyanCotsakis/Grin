import matplotlib.pyplot as plt
import numpy as np

def main():
    fig, ax = plt.subplots()
    ax.pcolormesh(np.random.random((100, 100)), cmap='gray')

    ClickToDrawPoints(ax).show()

class ClickToDrawPoints(object):
    def __init__(self, ax):
        self.ax = ax
        self.fig = ax.figure
        self.xy = []
        self.points = ax.plot([], [], color='red')
        self.fig.canvas.mpl_connect('button_press_event', self.on_click)

    def on_click(self, event):
        if event.inaxes is None:
            return
        self.xy.append([event.xdata, event.ydata])
        self.ax.draw_artist(self.points)
        self.fig.canvas.blit(self.fig.bbox)

    def show(self):
        plt.show()

main()