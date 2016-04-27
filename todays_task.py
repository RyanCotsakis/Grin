"""
A simple example of an animated plot
"""
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import threading
import time
import sys

plt.ion()
fig, ax = plt.subplots()
ax.plot([1.1,2.2],[1,2])

def update():
	while True:
		ymin, ymax = ax.get_ylim()
		print ymax
		sys.stdout.flush()
		ax.set_ylim([ymin - .5, ymax + .5])
		time.sleep(2)

try:
   upthr = threading.Thread(target = update)
   upthr.start()
except:
   raise Exception("Unable to start update thread.")

plt.show(block = True)

#not accessible:
ani = animation.FuncAnimation(fig, animate, np.arange(1, 200), init_func=init,
    interval=25, blit=True)