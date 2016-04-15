from Tkinter import *
import sys
import serial
import os
import subprocess
import thread
import win32gui

class UserInterface(Frame):
	processes = 0

	def __init__(self,parent):
		Frame.__init__(self,parent)
		self.parent = parent
		self.initUI()

	def initUI(self):
		self.parent.title("CA Reader")

		Label(text = "Select COM port:").grid(row = 0, column = 0)
		self.coms = Listbox(selectmode = SINGLE, height = 3)
		self.coms.grid(row = 0, column = 1)
		self.findCAs()
		self.coms.bind("<<ListboxSelect>>", self.selectCom) 
		Button(text = "Refresh List", command = self.findCAs).grid(row = 1, column = 1)
		self.comNum = ""

		Label(text = "Battery Serial No.").grid(row = 0, column = 2)
		self.serEntry = Entry()
		self.serEntry.grid(row = 0, column = 3)

		Button(text = "Launch Test", command = self.start).grid(row = 1, column = 2, columnspan = 2)

		Label(text = "Processes:").grid(row = 2, column = 0, columnspan = 5)
		Label(text = "Serial").grid(row = 3, column = 0)
		Label(text = "Elapsed").grid(row = 3, column = 1)
		Label(text = "Amp-Hours").grid(row = 3, column = 2)
		Label(text = "Voltage").grid(row = 3, column = 3)
		Label(text = "Current").grid(row = 3, column = 4)

	def start(self):
		serNum = self.serEntry.get()
		if serNum != "" and self.comNum != "":
			process = subprocess.Popen(["python", "CA_Reader.py", self.comNum, serNum], stdout = subprocess.PIPE)
			self.coms.delete(self.listIndex)
			stats = [self.comNum + ": " + serNum[-4:]]
			for i in range(4):
				newVar = StringVar()
				stats.append(newVar)
				newVar.set("")
			thread.start_new_thread(self.displayStats,(process,serNum,stats))
			self.processes += 1
			Button(text = stats[0], command = lambda: self.bringToFront(serNum)).grid(row = 3+self.processes, column = 0)
			Label(textvariable = stats[1]).grid(row = 3+self.processes, column = 1)
			Label(textvariable = stats[2]).grid(row = 3+self.processes, column = 2)
			Label(textvariable = stats[3]).grid(row = 3+self.processes, column = 3)
			Label(textvariable = stats[4]).grid(row = 3+self.processes, column = 4)
		self.comNum = ""
	def bringToFront(self,serial):
		cb = lambda x,y: y.append(x)
		wins = []
		win32gui.EnumWindows(cb,wins)

		for win in wins:
			txt = win32gui.GetWindowText(win)
			sys.stdout.flush()
			if serial == txt:
				win32gui.SetForegroundWindow(win)
				return

	def selectCom(self, val):
		sender = val.widget
		self.listIndex = sender.curselection()
		try:
			self.comNum = sender.get(self.listIndex)
		except:
			self.comNum = ""

	def displayStats(self, process, serial, out):
		while process.poll() is None:
			inp = process.stdout.readline().split()
			if len(inp) > 4:
				out[1].set(inp[0])
				out[2].set(inp[4])
				out[3].set(inp[1])
				out[4].set(inp[2])
		out[3].set("N/A")
		out[4].set("N/A")

	def findCAs(self):
		sys.stdout.flush()
		self.coms.delete(0, END)
		for com in range(100):
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
				self.coms.insert(END, COMport)

root = Tk()
root.geometry("500x300+100+100")
UserInterface(root)
root.mainloop()