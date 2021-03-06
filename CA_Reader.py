import matplotlib.pyplot as plt
import matplotlib.animation as animation
import matplotlib as mpl
import reportlab.platypus as plat
import serial
import sys
import os
import time
import thread
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

This script provides a user interface to interact with Cycle Analysts and displays their data on
interactive plots.

Author: Ryan Cotsakis, GRIN TECH

"""

#this function produces a pdf file with the discharge stats.
#The function locates two files in the directory of the executable,
#those are 'serialNumber'.jpg, and grin_header.jpg
#load - average current throughout discharge (possibly change to capacity/times[len(times)-1]*60)
#filename - directory included in filename
#capacity - total Ah
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




#This is a new process. Many instances of this process may be running at any time.
#COM - serial port
#serialNumber - window title
#conn - pipe to the main window
def CA_Reader(conn, COM, serialNumber):

	#final vars. these vars do not change throughout execution
	VscaleJump = 10 #voltage scale goes from 0 to (max voltage rounded up to nearest Iscalejump)
	IscaleJump = 1
	timeOut = 30 #seconds until program automatically pauses
	header_string = "Time\t\tV\t\tA\t\tWh\t\tAh\tCA_Ah"
	hourSwitch = 90 #minutes until the x axis is in hours

	#lists
	voltages = []
	currents = []
	times = []
	wattHours = [0]
	ampHours = [0]

	#modifiable global vars. the animate function copies instances of existing variables when it is called.
	#wrapping the variables in lists allows the function to access the changes, as the list instance is the same.
	paused = [False,0] #Is the plot currently paused?; time.time() when paused was initiated.
	cursorIndex = [-1] #-1 if cursor not visible, else the index of the x coordinate selected
	zoom = [0,0,0,0,0,0,False] #0: x1 --- 1: y1 --- 2: x2 --- 3: y2 --- 4: x coordinate of mouse clicke --- 5: y coordinate of mouse click --- 6: currently clicked?
	mouseLoc = [0,0] #x, y of mouse location when click held down
	scroll = [False,False,False,False,0] #scroll up? scroll down? scroll left? scroll right? time.time() when mouse was clicked.
	ahAxes = [False] #amp hour axis? If not, use time axis

	#open serial port passed to the CA_Reader function (COM)
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

	#SETUP
	mpl.rcParams['toolbar'] = 'None' #gets rid of default toolbar

	#prevent system sleep. This will interupt data collection
	ES_SYSTEM_REQUIRED = 0x00000001
	ES_CONTINUOUS = 0x80000000
	ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)

	#matplotlib initiation
	fig = plt.figure(serialNumber)
	Vgraph = fig.add_subplot(111)
	Igraph = Vgraph.twinx() #Current and voltage on the same graph

	startTime = [time.time()] #start time at 0, global modifiable variable (not actually a list)

	#BUTTONS
	#pause/play button
	def pause(self):
		cursorIndex[0] = -1
		thisTime = time.time()
		paused[0] ^= True #binary operation ^ swtiches paused[0] between true and false
		if paused[0]: #just now paused
			paused[1] = thisTime
			conn.send(["Paused"])
		if not paused[0]: #just now unpaused
			startTime[0] += (thisTime - paused[1]) #start time was also ticking while plot was paused.
		zoom[6] = False #we have let go of the click. this may be redundant
	#create button with this function
	bpause = mpl.widgets.Button(plt.axes([0.65,0.92,0.09,0.05]),"Pause")
	bpause.on_clicked(pause)

	#reset to default zoom and no cusor
	def reset(self):
		cursorIndex[0] = -1
		zoom[2], zoom[3] = zoom[0], zoom[1] # zoom to box with area = 0. this will unzoom
		zoom[6] = False #unclick mouse
	#create button
	breset = mpl.widgets.Button(plt.axes([0.75,0.92,0.15,0.05]),"Reset View")
	breset.on_clicked(reset)

	#clear plot
	def clear_plot(self):
		#delete all of the lists
		for i in range(len(times)):
			del voltages[0]
			del currents[0]
			del times[0]
			if len(wattHours) > 1:
				del wattHours[1]
				del ampHours[1]
		startTime[0] = time.time()
		if paused[0]:
			paused[1] = startTime[0] #time when paused initiated. only has an effect if already paused when plot was cleared.
		reset(self) #user defined function ^
		#clear text doc
		f.seek(0)
		f.write(header_string+ "\n")
		f.truncate()
	#create button
	bclear = mpl.widgets.Button(plt.axes([0.49,0.92,0.15,0.05]),"Clear Plot")
	bclear.on_clicked(clear_plot)

	#save to pdf and txt doc
	def save(self):
		reset(self)
		ahAxes[0] = True #default view for pdfs

		root = Tk()
		root.withdraw() #hides pointless tk window
		filename = tkFileDialog.asksaveasfilename()
		#get rid of file extension
		filename = filename.split(".")
		filename = filename[0]

		if filename != "":
			#fig screencap -- smile! :-)
			plt.savefig(serialNumber + ".jpg")
			#copy txt doc
			command = "copy \"" + serialNumber + ".txt\" \"" + filename + ".txt\" >NUL"
			os.system(command)
			#communicate
			conn.send(["Saved", filename + ".pdf", serialNumber, sum(currents)/len(currents), ampHours[len(ampHours)-1], voltages[max([len(voltages)-3, 0])]])
	#create button widget
	bsave = mpl.widgets.Button(plt.axes([0.33,0.92,0.15,0.05]),"Save Data")
	bsave.on_clicked(save)

	#xaxis switch units
	def xaxis(self):
		#(see pause function above)
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
		paused[1] = time.time() - timeOut #paused when the CA cut out. timeout declared at top of this script.
		conn.send(["Timed Out"])

	#NEW THREAD - Collects data
	def readCA():

		f = startFile()
		prevTime = time.time()

		while True:
			while not paused[0]: #read data
				line = ca.readline(100) #100 bytes
				values = line.split()

				try:
					CAah = float(values[0])
					voltage = float(values[1])
					current = float(values[2])
					thisTime = time.time()
					elapsed = thisTime-startTime[0]
				except: #the elements in the line read could not be cast to floats.
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

					conn.send(["Active", string])

					try:
						f.write(string + "\n")
						f.flush()
					except:
						break

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
					times.append(elapsed/60) #times is array of minutes elapsed not seconds for nice x axis ticks.
					voltages.append(voltage)
					currents.append(current)
					prevTime = thisTime
				except:
					break

			ca.readline(100) #clear buffer while paused
			time.sleep(.05)

	try:
	   CAthr = thread.start_new_thread(readCA,())
	except:
	   raise Exception("Unable to start readCA thread.")


	#MATPLOTLIB ACTION HANDLING

	#start cursor, mark for first zoom corner
	def on_press(event): #on mouse click
		if event.xdata != None: #if inside the plot
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
						cursorIndex[0] = ampHours.index(t) #only difference is use times array
						break
					else:
						cursorIndex[0] = len(ampHours)-1
			zoom[6] = True #consider the mouse clicked.
		zoom[4], zoom[5] = event.xdata, event.ydata
		mouseLoc[0], mouseLoc[1] = zoom[4], zoom[5]

	#zoom to box
	def on_release(event):
		if event.xdata != None:
			x = event.xdata
			y = event.ydata
			if x != zoom[4] and y !=zoom[5] and zoom[4] != None and zoom[5] != None: #commit zoom
				zoom[0],zoom[1] = zoom[4],zoom[5] #one corner of box where mouse was clicked, opposite where released.
				zoom[2],zoom[3] = x,y
		zoom[6] = False #not clicked.

	def on_key(event): #possible scroll command
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
		scroll[0] = False
		scroll[1] = False
		scroll[2] = False
		scroll[3] = False


	def on_motion(event):
		if zoom[6]: #if the mouse is clicked
			mouseLoc[0], mouseLoc[1] = event.xdata, event.ydata

	#this is a special function. Matplotlib uses it to update the axes.
	#i -- frame number. implicit += 1 after each call
	def animate(i):
		#start fresh. This may or may not be one of the main reasons why the graph updates so slowly.
		Vgraph.clear()
		Igraph.clear()

		#profiling
		thisTime = time.time()

		if not paused[0]:
			elapsed = (thisTime - startTime[0])/60
		else:
			elapsed = (paused[1] - startTime[0])/60

		#change axes labels. this usually produces no change, so theres certainly a clever way of avoiding this update every time.
		if not ahAxes[0]:
			if elapsed <= hourSwitch:
				timeUnit = "mins"
				if elapsed < 1 and (zoom[0] == zoom[2] or zoom[1] == zoom[3]): # if less than one minute passed, and default zoom:
					Vgraph.set_xlim([0,1])
			else:
				timeUnit = "h"
				newLabels = ["%0.1f" %(number/60) for number in Vgraph.get_xticks()]
				Vgraph.set_xticklabels(newLabels)
			Vgraph.set_xlabel("Time (" + timeUnit +")")
		else:
			Vgraph.set_xlabel("Amp Hours")

		#never changes
		Vgraph.set_ylabel("Voltage (V)", color = "b")
		Igraph.set_ylabel("Current (A)", color = "r")



		#AXES LIMITS:
		Vmin, Vmax = Vgraph.get_ylim()
		Imin, Imax = Igraph.get_ylim()
		tmin, tmax = Vgraph.get_xlim()

		#hold scroll left and right
		scrollDelay = .5 #seconds until scroll initiated
		if (scroll[3] or scroll[2]) and thisTime - scroll[4] > scrollDelay: #right or left

			#this counts the number of data points in the current view
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

			movement = numInBetween//60+1 #move one more than one sixtyth <-- probably spelt wrong
			if scroll[2] and cursorIndex[0]-movement >= 0: #if left
				cursorIndex[0] -= movement
			elif scroll[3] and cursorIndex[0]+movement < len(times): #if right
				cursorIndex[0] += movement

		#hold scroll vertical
		vertMove = (Imax-Imin)/3
		horiMove = (tmax-tmin)/3
		if (scroll[1] or scroll[0]) and thisTime - scroll[4] > scrollDelay: #down or up
			#note that if the zoom is defaluted, it will remain default after this call.
			if scroll[0]:
				zoom[1] += vertMove
				zoom[3] += vertMove
			else:
				zoom[1] -= vertMove
				zoom[3] -= vertMove

		#make sure cursor is in the window
		if cursorIndex[0] >= 0 and cursorIndex[0] < len(times): #not too sure if the second condition is necessary.
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
		if zoom[0] == zoom[2] or zoom[1] == zoom[3]: #default if zoom area is 0
			Vmin = 0
			Imin = 0
			if len(currents) > 0:
				Vmax = VscaleJump*((max(voltages)+VscaleJump/2)//VscaleJump+1) #this could probably be simplified.
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
		#ANNOTATE FORMAT: axes.annotate(text, coordinates, typeOfCoordinates (I used exclusively axes fraction))
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
			Vgraph.plot([t],[v],'bo') #blue dot
			Igraph.plot([t],[c],'ro') #red dot
			Igraph.plot([t,t],[Imin,Imax],"k--") #dashed line. k indicates black, -- indicates line type.
			Vgraph.annotate("Cursor:", xy=(0.01,0.95), xycoords = "axes fraction")
			Vgraph.annotate("%.2fs" %(60*times[cursorIndex[0]]), xy=(0.01,0.91), xycoords = "axes fraction")
			Vgraph.annotate("%.2fV" %v, xy=(0.01,0.87), xycoords = "axes fraction", color = "b")
			Vgraph.annotate("%.2fA" %c, xy=(0.01,0.83), xycoords = "axes fraction", color = "r")
			if cursorIndex[0] > 0:
				Vgraph.annotate("%.1fWh" %wattHours[cursorIndex[0]], xy=(0.01,0.79), xycoords = "axes fraction")
				Vgraph.annotate("%.2fAh" %ampHours[cursorIndex[0]], xy=(0.01,0.75), xycoords = "axes fraction")
		if zoom[6]:
			boxType = "k-."
			Igraph.plot([zoom[4],zoom[4]], [zoom[5],mouseLoc[1]], boxType) #old vert
			Igraph.plot([zoom[4],mouseLoc[0]], [zoom[5],zoom[5]], boxType) #old hori
			Igraph.plot([mouseLoc[0],mouseLoc[0]], [zoom[5],mouseLoc[1]], boxType) #new vert
			Igraph.plot([zoom[4],mouseLoc[0]], [mouseLoc[1],mouseLoc[1]], boxType) #new hori

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

	#this connects the event handling functions above to actual events
	cid = fig.canvas.mpl_connect('button_press_event', on_press)
	cid = fig.canvas.mpl_connect('button_release_event', on_release)
	cid = fig.canvas.mpl_connect('key_press_event', on_key)
	cid = fig.canvas.mpl_connect('motion_notify_event', on_motion)
	cid = fig.canvas.mpl_connect('key_release_event', on_key_release)

	#heres the line that starts the funky animate function using matplotlibs animate feature.
	ani = animation.FuncAnimation(fig, animate, interval = 10)

	plt.show() #this is a blocking method.
	#resume on exit window.
	conn.send(["Closed"])
	f.close()
	os.system("del \"" + serialNumber + ".txt\" \"" + serialNumber + ".jpg\"") #windows command to delete text doc.



#THIS IS THE MAIN PROCESS
class UserInterface(Frame):
	processes = 0 #not including self. Number of CA's being read.

	#TKinter tutorials will tell more a bout this than I can.
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

		padding = 20
		Label(text = "Processes:").grid(row = 2, column = 0, columnspan = 5)
		Label(text = "Serial").grid(row = 3, column = 0, padx = padding)
		Label(text = "Elapsed").grid(row = 3, column = 1, padx = padding)
		Label(text = "Amp-Hours").grid(row = 3, column = 2, padx = padding)
		Label(text = "Voltage").grid(row = 3, column = 3, padx = padding)
		Label(text = "Current").grid(row = 3, column = 4, padx = padding)
		Label(text = "Status").grid(row = 3, column = 5, padx = padding)


	def start(self): #start new process (CA_Reader)
		serNum = self.serEntry.get()
		if serNum != "" and self.comNum != "":

			if __name__ == '__main__':
				parent_conn, child_conn = Pipe()
				process = Process(target = CA_Reader, args = (child_conn, self.comNum, serNum))
				process.start()

			stats = [self.comNum + ": " + serNum[-4:]] #label on button that opens the CA_Reader window.

			for i in range(5):
				newVar = StringVar()
				stats.append(newVar)
				newVar.set("")
				if i == 4:
					newVar.set("Launching")
			thread.start_new_thread(self.displayStats,(parent_conn,stats)) # new thread devoted to updating the stats for this CA_Reader.
			self.processes += 1
			Button(text = stats[0], command = lambda: self.bringToFront(serNum)).grid(row = 3+self.processes, column = 0)
			Label(textvariable = stats[1]).grid(row = 3+self.processes, column = 1)
			Label(textvariable = stats[2]).grid(row = 3+self.processes, column = 2)
			Label(textvariable = stats[3]).grid(row = 3+self.processes, column = 3)
			Label(textvariable = stats[4]).grid(row = 3+self.processes, column = 4)
			Label(textvariable = stats[5]).grid(row = 3+self.processes, column = 5)

			self.COMlabel.set("Not Selected")
			self.comNum = ""

	#maximizes the window with title "serial"
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

	#display available coms.
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

	#seperate thread devoted to updating the stats
	def displayStats(self, conn, out):
		while True: #dont worry, theres a break down below
			inp = conn.recv() 
			if inp[0] == "Active":
				inp = inp[1].split() #this is exactly whats getting sent to the text document, split into substrings of course.
				out[1].set(inp[0]) #elapsed
				out[2].set(inp[4]) #ah
				out[3].set(inp[1]) #vol
				out[4].set(inp[2]) #current
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

#start main thread
if __name__ == '__main__':
	root = Tk()
	root.geometry("600x300+100+100")
	UserInterface(root)
	root.mainloop()