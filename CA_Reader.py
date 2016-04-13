import serial
import sys
import os
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import matplotlib as mpl
import time
import thread
import FileDialog
import tkFileDialog
import Tkinter
import ctypes

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch


#final vars
VscaleJump = 10 #voltage scale goes from 0 to (max voltage rounded up to nearest Iscalejump)
IscaleJump = 1
timeOut = 30 #seconds until program automatically pauses
header_string = "Time\t\tV\t\tA\t\tWh\t\tAh\tCA_Ah"
maxFreq = .9 #number of seconds between data points

#lists
voltages = []
currents = []
times = []
wattHours = [0]
ampHours = [0]

#modifiable global vars. animate() can only access changes to lists.
paused = [False,0] #bool, time when paused was initiated
cursorIndex = [-1] #-1 if no cursor
zoom = [0,0,0,0,0,0,False] #0: x1; 1: y1; 2: x2; 3: y2; 4: xclicked; 5: yclicked; 6: clicked?
mouseLoc = [0,0] #x,y of mouse location when held click
scroll = [False,False,False,False,0] #up, down, left, right being held; time when pushed
ahAxes = [False] #amp hour axes? Or time axes?

print "Enter COM port number:"
sys.stdout.flush()
COMport = raw_input()
try:
	int(COMport)
	COMport = "COM" + COMport
except:
	"do nothing"

print "Enter battery serial number:"
sys.stdout.flush()
serialNumber = raw_input()

print "" #new line

#SETUP
mpl.rcParams['toolbar'] = 'None' #gets rid of default toolbar

#prevent sleep
ES_SYSTEM_REQUIRED = 0x00000001
ES_CONTINUOUS = 0x80000000
ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)

#matplotlib
fig = plt.figure()
Vgraph = fig.add_subplot(111)
Igraph = Vgraph.twinx()

startTime = [time.time()] #start time at 0, global modifiable variable (not actually a list)

#make pdf
def pdf(filename):
	doc = SimpleDocTemplate(filename,
			pagesize=letter,
			rightMargin=72,leftMargin=72,
			topMargin=72,bottomMargin=18)
	styles=getSampleStyleSheet()

	Story=[]
	spacing = 3

	Story.append(Image("grin_header.jpg", 6.5*inch, 1*inch))
	Story.append(Spacer(1, spacing))

	ptext = '<font size=18>Your Battery Test Results:</font>'
	Story.append(Paragraph(ptext, styles["Normal"]))
	Story.append(Spacer(1, 3*spacing))

	ptext = '<font size=12>Time of Test: %s</font>' %time.ctime()
	Story.append(Paragraph(ptext, styles["Normal"]))
	Story.append(Spacer(1, spacing))

	ptext = '<font size=12>Pack Serial #: %s</font>' %serialNumber
	Story.append(Paragraph(ptext, styles["Normal"]))
	Story.append(Spacer(1, spacing))

	ptext = "<font size=12>Load Current: %.2f A</font>" %(sum(currents)/len(currents))
	Story.append(Paragraph(ptext, styles["Normal"]))
	Story.append(Spacer(1, spacing))

	ptext = '<font size=12>Test Capacity: %.2f Ah</font>' %ampHours[len(ampHours)-1]
	Story.append(Paragraph(ptext, styles["Normal"]))
	Story.append(Spacer(1, spacing))

	ptext = '<font size=12>Cutoff Voltage: %.2f V</font>' %voltages[max([len(voltages)-3,0])]
	Story.append(Paragraph(ptext, styles["Normal"]))
	Story.append(Spacer(1, spacing))

	Story.append(Image(serialNumber + ".jpg", 7*inch, 5*inch))
	Story.append(Spacer(1, spacing))

	doc.build(Story)

#BUTTONS
#pause/play button
def pause(self):
	cursorIndex[0] = -1
	thisTime = time.time()
	paused[0] ^= True
	if paused[0]:
		paused[1] = thisTime
	if not paused[0]:
		startTime[0] += (thisTime - paused[1]) #start time was also ticking while plot was paused.
	zoom[6] = False
bpause = mpl.widgets.Button(plt.axes([0.65,0.92,0.09,0.05]),"Pause")
bpause.on_clicked(pause)

#reset to default zoom and no cusor button
def reset(self):
	cursorIndex[0] = -1
	zoom[2], zoom[3] = zoom[0], zoom[1]
	zoom[6] = False
breset = mpl.widgets.Button(plt.axes([0.75,0.92,0.15,0.05]),"Reset View")
breset.on_clicked(reset)

#clear plot
def clear_plot(self):
	for i in range(len(times)):
		del voltages[0]
		del currents[0]
		del times[0]
		if len(wattHours) > 1:
			del wattHours[1]
			del ampHours[1]
	startTime[0] = time.time()
	if paused[0]:
		paused[1] = startTime[0]
	reset(self)
	f.seek(0)
	f.write(header_string+ "\n")
	f.truncate()
	os.system("cls")
	print header_string
bclear = mpl.widgets.Button(plt.axes([0.49,0.92,0.15,0.05]),"Clear Plot")
bclear.on_clicked(clear_plot)

#save
def save(self):
	reset(self)
	ahAxes[0] = True
	root = Tkinter.Tk()
	root.withdraw()
	filename = tkFileDialog.asksaveasfilename()
	filename = filename.split(".")
	filename = filename[0]
	if filename != "":
		plt.savefig(serialNumber + ".jpg")
		command = "copy \"" + serialNumber + ".txt\" \"" + filename + ".txt\" >NUL"
		os.system(command)
		pdf(filename + ".pdf")
		print "...SAVED DATA..."
bsave = mpl.widgets.Button(plt.axes([0.33,0.92,0.15,0.05]),"Save Data")
bsave.on_clicked(save)

#xaxis
def xaxis(self):
	ahAxes[0] ^= True
	cursorIndex[0] = -1
	zoom[2], zoom[3] = zoom[0], zoom[1]
	zoom[6] = False
bxaxis = mpl.widgets.Button(plt.axes([0.63,0.02,0.2,0.04]),"Change to Ah")
bxaxis.on_clicked(xaxis)


def startFile():
	global f
	f = open(serialNumber + ".txt",'w')

	print header_string
	sys.stdout.flush()
	f.write(header_string + "\n")

	return f

def timeOutProcess():
	print "\nPAUSED. No data to CA for " + str(timeOut) + " seconds.\n"
	sys.stdout.flush()
	paused[0] = True
	paused[1] = time.time() - timeOut

def readCA():
	try:
		ca = serial.Serial(
		    port=COMport,
		    baudrate=9600,
		    parity=serial.PARITY_NONE,
		    stopbits=serial.STOPBITS_ONE,
		    bytesize=serial.EIGHTBITS
		)
		ca.isOpen()
	except:
		raise Exception("COM Port: '" + COMport + "' unavailable")

	f = startFile()
	prevTime = time.time()

	while True:
		try:
			thisTime = time.time()
		except:
			break
		while not paused[0] and thisTime - prevTime >= maxFreq: #read data
			line = ca.readline(100)
			values = line.split()

			try:
				CAah = float(values[0])
				voltage = float(values[1])
				current = float(values[2])
				elapsed = time.time()-startTime[0] # calculating time.time() frequently for accuracy
			except:
				break
			
			#print and write
			if len(values) > 0:
				mins = elapsed // 60
				secs = elapsed - 60 * mins
				hours = 0
				while mins >= 60:
					hours += 1
					mins -= 60
					sys.stdout.flush()
				length = len(times)
				if hours > 0:
					string = "%i:%02i:%04.1f" %(hours,mins,secs)
				else:
					string = "%02i:%04.1f\t" %(mins,secs)
				string += "\t%.2f\t\t%.2f\t\t%.1f\t\t%.2f\t%.2f" %(voltage,current,wattHours[length-1],ampHours[length-1],CAah)
				print string
				sys.stdout.flush()
				f.write(string + "\n")
				f.flush()

			#plot
			try:
				length = len(times)
				if length > 0:
					ah = current*(thisTime-prevTime)/3600
					wh = wattHours[length-1]+voltage *ah
					ah += ampHours[length-1]
					if ah < 0:
						ah = 0
						wh = 0
					wattHours.append(wh)
					ampHours.append(ah)
				times.append(elapsed/60)
				voltages.append(voltage)
				currents.append(current)
				prevTime = thisTime
			except:
				break

		ca.readline(100) #clear buffer while paused

try:
   CAthr = thread.start_new_thread(readCA,())
except:
   print "Error: unable to start thread"


#MATPLOTLIB ACTION HANDLING
#start cursor, mark for first zoom corner
def on_press(event):
	if event.xdata != None:
		if not ahAxes[0]:
			for t in times:
				if t > event.xdata:
					cursorIndex[0] = times.index(t)
					break
				else:
					cursorIndex[0] = len(times)-1
		else:
			for t in ampHours:
				if t > event.xdata:
					cursorIndex[0] = ampHours.index(t)
					break
				else:
					cursorIndex[0] = len(ampHours)-1
		zoom[6] = True
	zoom[4], zoom[5] = event.xdata, event.ydata
	mouseLoc[0], mouseLoc[1] = zoom[4], zoom[5]

#zoom to box
def on_release(event):
	if event.xdata != None:
		x = event.xdata
		y = event.ydata
		if x != zoom[4] and y !=zoom[5] and zoom[4] != None and zoom[5] != None:
			zoom[0],zoom[1] = zoom[4],zoom[5]
			zoom[2],zoom[3] = x,y
	zoom[6] = False

def on_key(event):
	Imin, Imax = Igraph.get_ylim()
	vertMove = (Imax-Imin)/3
	scroll[4] = time.time()
	if event.key == "up":
		zoom[1] += vertMove
		zoom[3] += vertMove
		scroll[0] = True
	elif event.key == "down":
		zoom[1] -= vertMove
		zoom[3] -= vertMove
		scroll[1] = True
	elif event.key == "right" and cursorIndex[0] < len(times)-1 and cursorIndex[0] != -1:
		cursorIndex[0] += 1
		scroll[3] = True
	elif event.key == "left" and cursorIndex[0] > 0:
		cursorIndex[0] -= 1
		scroll[2] = True

def on_key_release(event):
	scroll[1] = False
	scroll[2] = False
	scroll[3] = False
	scroll[0] = False


def on_motion(event):
	if zoom[6]:
		mouseLoc[0], mouseLoc[1] = event.xdata, event.ydata

def animate(i):
	Vgraph.clear()
	Igraph.clear()

	thisTime = time.time()
	#axes labels
	if not paused[0]:
		elapsed = (thisTime - startTime[0])/60
	else:
		elapsed = (paused[1] - startTime[0])/60
	hourSwitch = 90 #minutes until the x axis is in hours

	if not ahAxes[0]:
		if elapsed <= hourSwitch:
			timeUnit = "mins"
			if elapsed < 1 and (zoom[0] == zoom[2] or zoom[1] == zoom[3]):
				Vgraph.set_xlim([0,1])
		else:
			timeUnit = "h"
			newLabels = ["%0.1f" %(number/60) for number in Vgraph.get_xticks()]
			Vgraph.set_xticklabels(newLabels)
		Vgraph.set_xlabel("Time (" + timeUnit +")")
	else:
		Vgraph.set_xlabel("Amp Hours")
	Vgraph.set_ylabel("Voltage (V)", color = "b")
	Igraph.set_ylabel("Current (A)", color = "r")

	#AXES LIMITS:
	Vmin, Vmax = Vgraph.get_ylim()
	Imin, Imax = Igraph.get_ylim()
	tmin, tmax = Vgraph.get_xlim()

	#hold scroll left and right
	scrollDelay = .5 #seconds until scroll initiated
	if (scroll[3] or scroll[2]) and thisTime - scroll[4] > scrollDelay: #right or left
		numInBetween = 0
		if not ahAxes[0]:
			for t in times:
				if t > tmin:
					numInBetween+=1
				if t > tmax:
					break
		else:
			for t in ampHours:
				if t > tmin:
					numInBetween+=1
				if t > tmax:
					break
		movement = numInBetween//60+1
		if scroll[2] and cursorIndex[0]-movement >= 0:
			cursorIndex[0] -= movement
		elif scroll[3] and cursorIndex[0]+movement < len(times):
			cursorIndex[0] += movement

	#hold scroll vertical
	vertMove = (Imax-Imin)/3
	horiMove = (tmax-tmin)/3
	if (scroll[1] or scroll[0]) and thisTime - scroll[4] > scrollDelay: #down or up
		if scroll[0]:
			zoom[1] += vertMove
			zoom[3] += vertMove
		else:
			zoom[1] -= vertMove
			zoom[3] -= vertMove

	#make sure cursor is in the window
	if cursorIndex >= 0 and cursorIndex < len(times):
		horiMove = (tmax-tmin)/3
		if not ahAxes[0]:
			if times[cursorIndex[0]] > tmax:
				zoom[0] += horiMove
				zoom[2] += horiMove
			elif times[cursorIndex[0]] < tmin:
				zoom[0] -= horiMove
				zoom[2] -= horiMove
		else:
			if ampHours[cursorIndex[0]] > tmax:
				zoom[0] += horiMove
				zoom[2] += horiMove
			elif ampHours[cursorIndex[0]] < tmin:
				zoom[0] -= horiMove
				zoom[2] -= horiMove
		
	#default zoom or modified zoom?
	if zoom[0] == zoom[2] or zoom[1] == zoom[3]:
		Vmin = 0
		Imin = 0
		if len(currents) > 0:
			Vmax = VscaleJump*((max(voltages)+VscaleJump/2)//VscaleJump+1)
			cur_max = max(currents)
		else:
			Vmax = VscaleJump
			cur_max = 0
		if cur_max > 0:
			Imax = IscaleJump*((cur_max+IscaleJump/2)//IscaleJump+1)
		else:
			Imax = IscaleJump
		if len(times) > 0:
			Vgraph.set_ylim([Vmin,Vmax])
			Igraph.set_ylim([Imin,Imax])
		else:
			Vgraph.set_ylim([0,VscaleJump])
			Igraph.set_ylim([0,IscaleJump])
	else:
		#commit zoomed axes
		y1 = (zoom[1]-Imin)*(Vmax-Vmin)/(Imax-Imin)+Vmin
		y2 = (zoom[3]-Imin)*(Vmax-Vmin)/(Imax-Imin)+Vmin
		Vgraph.set_ylim([min([y1,y2]),max(y1,y2)])
		Igraph.set_ylim([min([zoom[1],zoom[3]]),max(zoom[1],zoom[3])])
		Vgraph.set_xlim([min([zoom[0],zoom[2]]),max(zoom[0],zoom[2])])

	#pause/play
	if paused[0]:
		bpause.label.set_text("Resume")
	else:
		bpause.label.set_text("Pause")

	#xaxis
	if ahAxes[0]:
		bxaxis.label.set_text("Change to Time")
	else:
		bxaxis.label.set_text("Change to Ah")

	#ANNOTATIONS
	length = len(times)
	if length > 0:
		Vgraph.annotate("V: %.2fV" %voltages[length-1], xy=(0.8,0.95), xycoords = "axes fraction", color = "b")
		Vgraph.annotate("I: %.2fA" %currents[length-1], xy=(0.8,0.91), xycoords = "axes fraction", color = "r")
		Vgraph.annotate("Wh: %.1fWh" %wattHours[length-1], xy=(0.8,0.87), xycoords = "axes fraction")
		Vgraph.annotate("Ah: %.2fAh" %ampHours[length-1], xy=(0.8,0.83), xycoords = "axes fraction")
	if cursorIndex[0] != -1:
		if not ahAxes[0]:
			t = times[cursorIndex[0]]
		else:
			t = ampHours[cursorIndex[0]]
		v = voltages[cursorIndex[0]]
		c = currents[cursorIndex[0]]
		Vgraph.plot([t],[v],'bo')
		Igraph.plot([t],[c],'ro')
		Igraph.plot([t,t],[Imin,Imax],"k--")
		Vgraph.annotate("Cursor:", xy=(0.01,0.95), xycoords = "axes fraction")
		Vgraph.annotate("%.2fs" %(60*times[cursorIndex[0]]), xy=(0.01,0.91), xycoords = "axes fraction")
		Vgraph.annotate("%.2fV" %v, xy=(0.01,0.87), xycoords = "axes fraction", color = "b")
		Vgraph.annotate("%.2fA" %c, xy=(0.01,0.83), xycoords = "axes fraction", color = "r")
		if cursorIndex[0] > 0:
			Vgraph.annotate("%.1fWh" %wattHours[cursorIndex[0]], xy=(0.01,0.79), xycoords = "axes fraction")
			Vgraph.annotate("%.2fAh" %ampHours[cursorIndex[0]], xy=(0.01,0.75), xycoords = "axes fraction")
	if zoom[6]:
		boxType = "k-."
		Igraph.plot([zoom[4],zoom[4]],[zoom[5],mouseLoc[1]],boxType) #old vert
		Igraph.plot([zoom[4],mouseLoc[0]],[zoom[5],zoom[5]],boxType) #old hori
		Igraph.plot([mouseLoc[0],mouseLoc[0]],[zoom[5],mouseLoc[1]],boxType) #new vert
		Igraph.plot([zoom[4],mouseLoc[0]],[mouseLoc[1],mouseLoc[1]],boxType) #new hori

	#timeOut
	if length > 0 and thisTime - startTime[0] >= times[length-1]*60 + timeOut and not paused[0]:
		timeOutProcess()

	try:
		#PLOT
		alph = 0.7
		if ahAxes[0]:
			Vgraph.plot(ampHours,voltages,"b", alpha = alph)
			Igraph.plot(ampHours,currents,"r", alpha = alph)
		else:
			Vgraph.plot(times,voltages,"b", alpha = alph)
			Igraph.plot(times,currents,"r", alpha = alph)
	except:
		"Try again"

		
cid = fig.canvas.mpl_connect('button_press_event', on_press)
cid = fig.canvas.mpl_connect('button_release_event', on_release)
cid = fig.canvas.mpl_connect('key_press_event', on_key)
cid = fig.canvas.mpl_connect('motion_notify_event', on_motion)
cid = fig.canvas.mpl_connect('key_release_event', on_key_release)
ani = animation.FuncAnimation(fig, animate, interval = 10)
plt.show()
f.close()
os.system("del \"" + serialNumber + ".txt\" \"" + serialNumber + ".jpg\"")

