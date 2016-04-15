from Tkinter import *
import sys

class Tester(Frame):
	def __init__(self,parent):
		Frame.__init__(self,parent)
		self.parent = parent
		self.initUI()
	
	def initUI(self):
		self.parent.title("Grin")

		Label(text = "Yo, B!").pack()

		self.greeting = StringVar()
		self.greeting.set("Hello!")
		Label(textvariable = self.greeting).pack()

		Entry().pack()

		Button(text = "push me", command = self.say_hello).pack()

		self.clicked = BooleanVar()
		cb = Checkbutton(text = "show title", variable = self.clicked, command = self.check)
		cb.select()
		cb.pack()

		self.lb = Listbox()
		for i in ["one", "two", "three"]:
			self.lb.insert(END, i)
		self.lb.bind("<<ListboxSelect>>", self.onSelect)    
		self.lb.pack()

		menu = Menu(self.parent)
		file = Menu(menu)
		file.add_command(label = "show laugh", command = self.laugh)
		menu.add_cascade(label = "file", menu = file)
		self.parent.config(menu = menu)

	def laugh(self):
		Label(text = "HAHAH!").pack()
		self.lb.insert(END, "HAHAH!!")

	def onSelect(self, val):
		sender = val.widget
		i = sender.curselection()
		self.greeting.set(sender.get(i))

	def say_hello(self):
		print "hello"
		sys.stdout.flush()

	def check(self):
		if self.clicked.get():
			self.parent.title("Grin")
		else:
			self.parent.title("")


root = Tk()
root.geometry("500x300+100+100")
Tester(root)
root.mainloop()