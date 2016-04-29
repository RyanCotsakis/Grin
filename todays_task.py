import matplotlib.pyplot as plt
import threading
import time
import sys

plt.ion()
fig, ax = plt.subplots()
vPlot, = ax.plot([1.1,2.2],[1,2])

def update():
	ymin, ymax = 0,3
	count = 0
	while True:
		print count
		sys.stdout.flush()

		ax.set_ylim([ymin - count, ymax + count]) #zoom out -- this is the line that causes problems

		time.sleep(.1) #decrease this to .01 and it'll throw the error
		count += 1

		if count > 10000:
			break

try:
   upthr = threading.Thread(target = update)
   upthr.start()
except:
   raise Exception("Unable to start update thread.")


plt.show(block = True)