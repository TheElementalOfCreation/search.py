import tkinter as tk
import tkinter.filedialog as tkFile

def askopenfilename(master = None, **options):
	dest = False
	if master == None:
		master = tk.Tk()
		dest = True
	a = tkFile.Open(master, **options).show()
	if dest:
		master.destroy()
	return a

def asksaveasfilename(master = None, **options):
	dest = False
	if master == None:
		master = tk.Tk()
		dest = True
	a = tkFile.SaveAs(master, **options).show()
	if dest:
		master.destroy()
	return a

def askopenfile(master = None, mode = 'r', **options):
	dest = False
	if master == None:
		master = tk.Tk()
		dest = True
	filename = tkFile.Open(master, **options).show()
	if dest:
		master.destroy()
	if filename:
		return open(filename, mode)
	return None

def askopenfiles(master = None, mode = 'r', **options):
	dest = False
	if master == None:
		master = tk.Tk()
		dest = True
	files = askopenfilenames(master, **options)
	if dest:
		master.destroy()
	if files:
		ofiles = []
		for filename in files:
			ofiles.append(open(filename, mode))
		files = ofiles
	return files

def asksaveasfile(master = None, mode = 'w', **options):
	dest = False
	if master == None:
		master = tk.Tk()
		dest = True
	filename = tkFile.SaveAs(master, **options).show()
	if dest:
		master.destroy()
	if filename:
		return open(filename, mode)
	return None

def askdirectory(master = None, **options):
	dest = False
	if master == None:
		master = tk.Tk()
		dest = True
	a = tkFile.Directory(**options).show()
	if dest:
		master.destroy()
	return a

def askopenfilenames(master = None, **options):
	dest = False
	if master == None:
		master = tk.Tk()
		dest = True
	options['multiple'] = 1
	files = tkFile.Open(master, **options).show()
	if dest:
		master.destroy()
	return files


askdirectory.__doc__ = tkFile.askdirectory.__doc__
askopenfile.__doc__ = tkFile.askopenfile.__doc__
askopenfiles.__doc__ = tkFile.askopenfiles.__doc__
askopenfilename.__doc__ = tkFile.askopenfilename.__doc__
askopenfilenames.__doc__ = tkFile.askopenfilenames.__doc__
asksaveasfile.__doc__ = tkFile.asksaveasfile.__doc__
asksaveasfilename.__doc__ = tkFile.asksaveasfilename.__doc__
