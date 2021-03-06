import matplotlib.pyplot as plt
import matplotlib.animation as animation
import matplotlib as mpl
import reportlab.platypus as plat
import serial
import sys
import os
import time
import threading
import FileDialog
import tkFileDialog
import ctypes
import win32gui
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from Tkinter import *
from functools import partial
from multiprocessing import Process, Pipe

"""
CA_READER

This script provides a user interface to open CA_Readers that
read data from Cycle Analysts and displays that data on
interactive plots.

Author: Ryan Cotsakis, GRIN TECH

"""


# --- HELPER FUNCTION ---

def pdf(filename, serialNumber, load, capacity, cutoffVoltage):
		doc = SimpleDocTemplate(filename,
				pagesize=letter,
				rightMargin=72,leftMargin=72,
				topMargin=72,bottomMargin=18)
		styles=getSampleStyleSheet()

		Story=[]
		spacing = 3
		im = plat.Image("grin_header.jpg", 7*inch, 1*inch)
		Story.append(im)
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

		ptext = "<font size=12>Load Current: %.2f A</font>" %load
		Story.append(Paragraph(ptext, styles["Normal"]))
		Story.append(Spacer(1, spacing))

		ptext = '<font size=12>Test Capacity: %.2f Ah</font>' %capacity
		Story.append(Paragraph(ptext, styles["Normal"]))
		Story.append(Spacer(1, spacing))

		ptext = '<font size=12>Cutoff Voltage: %.2f V</font>' %cutoffVoltage
		Story.append(Paragraph(ptext, styles["Normal"]))
		Story.append(Spacer(1, spacing))

		im = plat.Image(serialNumber + ".jpg", 7*inch, 5*inch)
		Story.append(im)
		Story.append(Spacer(1, spacing))

		doc.build(Story)


# --- NEW PROCESS ---
def CA_Reader(conn, COM, serialNumber):

	#final vars
	VscaleJump = 10 #voltage scale goes from 0 to (max voltage rounded up to nearest Iscalejump)
	IscaleJump = 1
	timeOut = 30 #seconds until program automatically pauses
	header_string = "Time\t\tV\t\tA\t\tWh\t\tAh\tCA_Ah"

	#lists
	voltages = []
	currents = []
	times = []
	wattHours = [0]
	ampHours = [0]

	#modifiable global vars. animate() can only access changes to lists.
	paused = [False,0] #bool, time when paused was initiated
	cursorIndex = [-1] #-1 if no cursor
	zoom = [0,0,0,0,0,0,False,False] #0: x1; 1: y1; 2: x2; 3: y2; 4: xclicked; 5: yclicked; 6: clicked? 7: update zoom?
	mouseLoc = [0,0] #x,y of mouse location when held click
	scroll = [False,False,False,False,0] #up, down, left, right being held; time when pushed
	ahAxes = [False] #amp hour axes? Or time axes?

	try:
		ca = serial.Serial(
			    port=COM,
			    baudrate=9600,
			    parity=serial.PARITY_NONE,
			    stopbits=serial.STOPBITS_ONE,
			    bytesize=serial.EIGHTBITS
			)
		ca.isOpen()
	except:
		raise Exception(COM + " unavailable.")

	# --- SETUP ---

	#prevent sleep
	ES_SYSTEM_REQUIRED = 0x00000001
	ES_CONTINUOUS = 0x80000000
	ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)

	startTime = [time.time()] #start time at 0, global modifiable variable (not actually a list)

	#matplotlib
	plt.ion()
	mpl.rcParams['toolbar'] = 'None' #gets rid of default toolbar
	fig = plt.figure(serialNumber)
	Vgraph = fig.add_subplot(111)
	Igraph = Vgraph.twinx()
	Vgraph.set_ylabel("Voltage (V)", color = "b")
	Igraph.set_ylabel("Current (A)", color = "r")


	# --- PLOT FUNCTIONS ---

	#BUTTONS
	#pause/play button
	def pause(self):
		cursorIndex[0] = -1
		thisTime = time.time()
		paused[0] ^= True
		if paused[0]:
			paused[1] = thisTime
			conn.send(["Paused"])
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
	bclear = mpl.widgets.Button(plt.axes([0.49,0.92,0.15,0.05]),"Clear Plot")
	bclear.on_clicked(clear_plot)

	#save
	def save(self):
		reset(self)
		ahAxes[0] = True
		root = Tk()
		root.withdraw()
		filename = tkFileDialog.asksaveasfilename()
		filename = filename.split(".")
		filename = filename[0]
		if filename != "":
			plt.savefig(serialNumber + ".jpg")
			command = "copy \"" + serialNumber + ".txt\" \"" + filename + ".txt\" >NUL"
			os.system(command)
			conn.send(["Saved", filename + ".pdf", serialNumber, sum(currents)/len(currents), ampHours[len(ampHours)-1], voltages[max([len(voltages)-3, 0])]])
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
		f.write(header_string + "\n")
		return f

	def timeOutProcess():
		paused[0] = True
		paused[1] = time.time() - timeOut
		conn.send(["Timed Out"])

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
		zoom[7] = True

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


	def changeLim(): #if there is no sleep between changes in the axes limits, there will be an error thrown.
		time.sleep(.1)

	# --- THREAD THAT UPDATES WHEN DATA RECEIVED ---

	def readCA():

		f = startFile()
		prevTime = time.time()
		alph = 0.7
		vPlot, = Vgraph.plot([],[],"b", alpha = alph)
		iPlot, = Igraph.plot([],[],"r", alpha = alph)

		while True:
			try:
				thisTime = time.time()
			except:
				break
			while not paused[0]: #read data
				line = ca.readline(100)
				values = line.split()

				try:
					CAah = float(values[0])
					voltage = float(values[1])
					current = float(values[2])
					elapsed = time.time()-startTime[0] # calculating time.time() frequently for accuracy
				except:
					break
				
				#write and send
				if len(values) > 0:
					mins = elapsed // 60
					secs = elapsed - 60 * mins
					hours = 0
					while mins >= 60:
						hours += 1
						mins -= 60
					length = len(times)
					if hours > 0:
						string = "%i:%02i:%04.1f" %(hours,mins,secs)
					else:
						string = "%02i:%04.1f\t" %(mins,secs)
					string += "\t%.2f\t\t%.2f\t\t%.1f\t\t%.2f\t%.2f" %(voltage,current,wattHours[length-1],ampHours[length-1],CAah)
					try:
						conn.send(["Active", string])
					except:
						return
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

				#matplotlib

				if ahAxes[0]:
					changePlot.acquire()
					vPlot.set_data(ampHours,voltages)
					iPlot.set_data(ampHours,currents)
					changePlot.release()
				else:
					changePlot.acquire()
					vPlot.set_data(times,voltages)
					iPlot.set_data(times,currents)
					changePlot.release()

				#default zoom
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
						changeLim()
						changePlot.acquire()
						Vgraph.set_ylim([Vmin,Vmax])
						changePlot.release()
						changeLim()
						changePlot.acquire()
						Igraph.set_ylim([Imin,Imax])
						changePlot.release()
					else:
						changeLim()
						changePlot.acquire()
						Vgraph.set_ylim([0,VscaleJump])
						changePlot.release()
						changeLim()
						changePlot.acquire()
						Igraph.set_ylim([0,IscaleJump])
						changePlot.release()

					if mins == 0:
						changeLim()
						changePlot.acquire()
						Vgraph.set_xlim([0,1])
						changePlot.release()

			ca.readline(100) #clear buffer while paused


	# --- THREAD THAT UPDATES QUICKLY ---

	def animate():
		volAn = Vgraph.annotate("", xy=(0.8,0.95), xycoords = "axes fraction", color = "b")
		curAn = Vgraph.annotate("", xy=(0.8,0.91), xycoords = "axes fraction", color = "r")
		whAn = Vgraph.annotate("", xy=(0.8,0.87), xycoords = "axes fraction")
		ahAn = Vgraph.annotate("", xy=(0.8,0.83), xycoords = "axes fraction")

		bdot, = Vgraph.plot([],[],'bo')
		rdot, = Igraph.plot([],[],'ro')
		curLine, = Igraph.plot([],[],"k--")
		legTit = Vgraph.annotate("Cursor:", xy=(0.01,0.95), xycoords = "axes fraction")
		legTime = Vgraph.annotate("", xy=(0.01,0.91), xycoords = "axes fraction")
		legVol = Vgraph.annotate("", xy=(0.01,0.87), xycoords = "axes fraction", color = "b")
		legCur = Vgraph.annotate("", xy=(0.01,0.83), xycoords = "axes fraction", color = "r")
		legWh = Vgraph.annotate("", xy=(0.01,0.79), xycoords = "axes fraction")
		legAh = Vgraph.annotate("", xy=(0.01,0.75), xycoords = "axes fraction")

		boxType = "k-."
		oldVert, = Igraph.plot([],[],boxType) #old vert
		oldHori, = Igraph.plot([],[],boxType) #old hori
		newVert, = Igraph.plot([],[],boxType) #new vert
		newHori, = Igraph.plot([],[],boxType) #new hori

		while True:
			#Vgraph.clear()
			#Igraph.clear()

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
				else:
					timeUnit = "h"
					newLabels = ["%0.1f" %(number/60) for number in Vgraph.get_xticks()]
					Vgraph.set_xticklabels(newLabels)
				Vgraph.set_xlabel("Time (" + timeUnit +")")
			else:
				Vgraph.set_xlabel("Amp Hours")

			#AXES LIMITS:
			changePlot.acquire()
			Vmin, Vmax = Vgraph.get_ylim()
			Imin, Imax = Igraph.get_ylim()
			tmin, tmax = Vgraph.get_xlim()
			changePlot.release()

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
			if cursorIndex[0] >= 0 and cursorIndex[0] < len(times):
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

			#zoom to zoom window
			if zoom[7] and not (zoom[0] == zoom[2] or zoom[1] == zoom[3]):

				#commit zoomed axes
				y1 = (zoom[1]-Imin)*(Vmax-Vmin)/(Imax-Imin)+Vmin
				y2 = (zoom[3]-Imin)*(Vmax-Vmin)/(Imax-Imin)+Vmin
				changeLim()
				Vgraph.set_ylim([min([y1,y2]),max(y1,y2)])
				changeLim()
				Igraph.set_ylim([min([zoom[1],zoom[3]]),max(zoom[1],zoom[3])])
				changeLim()
				Vgraph.set_xlim([min([zoom[0],zoom[2]]),max(zoom[0],zoom[2])])
				zoom[7] = False

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

			length = len(times)
			condition = not cursorIndex[0] == -1

			#PUT THAT STUFF ON THAT GRAPH!!
			bdot.set_visible(condition)
			rdot.set_visible(condition)
			legTit.set_visible(condition)
			legTime.set_visible(condition)
			legVol.set_visible(condition)
			legCur.set_visible(condition)
			curLine.set_visible(condition)
			legWh.set_visible(condition)
			legAh.set_visible(condition)
			oldVert.set_visible(zoom[6])
			oldHori.set_visible(zoom[6])
			newVert.set_visible(zoom[6])
			newHori.set_visible(zoom[6])

			if length > 0:
				volAn.set_text("V: %.2fV" %voltages[length-1])
				curAn.set_text("I: %.2fA" %currents[length-1])
				whAn.set_text("Wh: %.1fWh" %wattHours[length-1])
				ahAn.set_text("Ah: %.2fAh" %ampHours[length-1])
			if condition:
				if not ahAxes[0]:
					t = times[cursorIndex[0]]
				else:
					t = ampHours[cursorIndex[0]]
				v = voltages[cursorIndex[0]]
				c = currents[cursorIndex[0]]
				bdot.set_data([t],[v])
				rdot.set_data([t],[c])
				curLine.set_data([t,t],[Imin,Imax])
				legTime.set_text("%.2fs" %(60*times[cursorIndex[0]]))
				legVol.set_text("%.2fV" %v)
				legCur.set_text("%.2fA" %c)
				legWh.set_text("%.1fWh" %wattHours[cursorIndex[0]])
				legAh.set_text("%.2fAh" %ampHours[cursorIndex[0]])
			if zoom[6]:
				oldVert.set_data([zoom[4],zoom[4]],[zoom[5],mouseLoc[1]]) #old vert
				oldHori.set_data([zoom[4],mouseLoc[0]],[zoom[5],zoom[5]]) #old hori
				newVert.set_data([mouseLoc[0],mouseLoc[0]],[zoom[5],mouseLoc[1]]) #new vert
				newHori.set_data([zoom[4],mouseLoc[0]],[mouseLoc[1],mouseLoc[1]]) #new hori


			#timeOut
			if length > 0 and thisTime - startTime[0] >= times[length-1]*60 + timeOut and not paused[0]:
				timeOutProcess()

			#sleep for x ms
			x = 30
			time.sleep(x/1000)

			
	# --- INITIATE THE THREADS ---

	cid = fig.canvas.mpl_connect('button_press_event', on_press)
	cid = fig.canvas.mpl_connect('button_release_event', on_release)
	cid = fig.canvas.mpl_connect('key_press_event', on_key)
	cid = fig.canvas.mpl_connect('motion_notify_event', on_motion)
	cid = fig.canvas.mpl_connect('key_release_event', on_key_release)

	changePlot = threading.Lock()

	try:
	    read_ca_thr = threading.Thread(target = readCA)
	    read_ca_thr.start()
	except:
	    raise Exception("Unable to start readCA thread.")

	try:
	    anithr = threading.Thread(target = animate)
	    anithr.start()
	except:
	    raise Exception("Unable to start animate thread.")
	
	plt.show(block = True)

	conn.send(["Closed"])
	f.close()
	os.system("del \"" + serialNumber + ".txt\" \"" + serialNumber + ".jpg\"")

# --- USER INTERFACE, MAIN PROCESS ---

class UserInterface(Frame):
	processes = 0

	def __init__(self,parent):
		Frame.__init__(self,parent)
		self.parent = parent
		self.initUI()

	def initUI(self):
		self.parent.title("CA Reader")

		self.menubar = Menu(self.parent)
		self.comMenu = Menu(self.menubar, postcommand = self.openComMenu, tearoff = 0)
		self.itemsInMenu = 0
		self.menubar.add_cascade(label = "Select COM Port", menu = self.comMenu)
		self.parent.config(menu = self.menubar)
		self.comNum = ""

		Label(text = "COM Port:").grid(row = 0, column = 0)
		self.COMlabel = StringVar()
		self.COMlabel.set("Not Selected")
		Label(textvariable = self.COMlabel).grid(row = 0, column = 1)

		Label(text = "Battery Serial #:").grid(row = 1, column = 0)
		self.serEntry = Entry()
		self.serEntry.grid(row = 1, column = 1)

		Button(text = "Launch Test", command = self.start).grid(row = 0, column = 2, rowspan = 2)

		Label(text = "Processes:").grid(row = 2, column = 0, columnspan = 5)
		Label(text = "Serial").grid(row = 3, column = 0)
		Label(text = "Elapsed").grid(row = 3, column = 1)
		Label(text = "Amp-Hours").grid(row = 3, column = 2)
		Label(text = "Voltage").grid(row = 3, column = 3)
		Label(text = "Current").grid(row = 3, column = 4)
		Label(text = "Status").grid(row = 3, column = 5)




	def start(self):
		serNum = self.serEntry.get()
		if serNum != "" and self.comNum != "":

			if __name__ == '__main__':
				parent_conn, child_conn = Pipe()
				process = Process(target = CA_Reader, args = (child_conn,self.comNum, serNum))
				process.start()

			stats = [self.comNum + ": " + serNum[-4:]]
			for i in range(5):
				newVar = StringVar()
				stats.append(newVar)
				newVar.set("")
				if i == 4:
					newVar.set("Launching")

			try:
			    threading.Thread(target = self.displayStats, args = (parent_conn,process,serNum,stats)).start()
			except:
			    raise Exception("Unable to start displayStats thread.")

			self.processes += 1
			Button(text = stats[0], command = lambda: self.bringToFront(serNum)).grid(row = 3+self.processes, column = 0)
			Label(textvariable = stats[1]).grid(row = 3+self.processes, column = 1)
			Label(textvariable = stats[2]).grid(row = 3+self.processes, column = 2)
			Label(textvariable = stats[3]).grid(row = 3+self.processes, column = 3)
			Label(textvariable = stats[4]).grid(row = 3+self.processes, column = 4)
			Label(textvariable = stats[5]).grid(row = 3+self.processes, column = 5)

			self.COMlabel.set("Not Selected")
			self.comNum = ""

	def bringToFront(self,serial):
		cb = lambda x,y: y.append(x)
		wins = []
		win32gui.EnumWindows(cb,wins)

		for win in wins:
			txt = win32gui.GetWindowText(win)
			if serial == txt:
				win32gui.ShowWindow(win, 3) #https://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
				return

	def selectCom(self, name):
		self.comNum = name
		self.COMlabel.set(name)

	def openComMenu(self):
		self.comMenu.delete(0,self.itemsInMenu)
		self.itemsInMenu = 0
		for com in range(255):
			COMport = "COM%i" %com
			goodCom = True
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
				goodCom = False
			if goodCom:
				self.itemsInMenu+=1
				self.comMenu.add_command(label = COMport, command = partial(self.selectCom,COMport))
		if self.itemsInMenu == 0:
			self.comMenu.add_command(label = "No ports found")
			self.itemsInMenu=1

	def displayStats(self, conn, process, serial, out):
		while True:
			inp = conn.recv()
			if inp[0] == "Active":
				inp = inp[1].split()
				out[1].set(inp[0])
				out[2].set(inp[4])
				out[3].set(inp[1])
				out[4].set(inp[2])
				out[5].set("Active")
			else:
				if inp[0] == "Saved":
					pdf(inp[1],inp[2],inp[3],inp[4],inp[5])
				else:
					out[3].set("N/A")
					out[4].set("N/A")
				out[5].set(inp[0])
				if inp[0] == "Closed":
					break

if __name__ == '__main__':
	root = Tk()
	root.geometry("500x300+100+100")
	UserInterface(root)
	root.mainloop()