import json
import os
import progressbar
import shutil
import tkinter as tk

from creatorUtils import path


INT_SEARCH = 0
INT_COPY   = 1
RANGE_5    = tuple(range(5))

exList = []
g = None

def reset():
	global exList, g
	g = {
		'interrupt': None,
		'files': None,
		'filepaths': None,
		'remaining': None,
	}
	exList = []

reset()

def tryToCopy(source, output, return_error = False):
	es = []
	err = None
	ms = 'Could not copy file '+ source + '\nReason'
	ms2 = ':\n'
	same = True
	for x in RANGE_5:
		try:
			shutil.copyfile(source, output)
			os.stat(output)
			return
		except Exception as e:
			es.append(e)
			err = e.errno
			if x == 4:
				for y in es:
					if y.errno != err:
						if same:
							same = False
							ms += 's'
						ms2 += '{}\n'.format(y)
				if return_error:
					return ms + ms2
				else:
					print(ms + ms2)

def recordExceptions(exception):
	global exList
	if exception.errno == 2:
		exList.extend(exception)


def handleWalkExceptions(files, filepaths):
	global g, exList
	a = []
	if exList != []:
		g['interrupt'] = INT_SEARCH
		for x in exList:
			try:
				# Error processing
				a.append(x.filename)
			except:
				pass
		g['remaining'] = a
	g['files'] = files
	g['filepaths'] = filepaths



def loadPartialSession(sessionFilePath):
	global g
	with open(sessionFilePath, 'r') as f:
		a = json.load(f)

def start():
	global exList, g
	while True:
		try:
			a = input('Enter folder path you want to search (You may just drag the folder onto this window):\n')
			a = a.replace('"', '').replace("'", '')
			f = input('Enter destination folder (Again, you can drag and drop):\n')
			f = f.replace('\\', '/').replace('"', '').replace("'", '')
			try:
				os.stat(f)
			except Exception as e:
				raise Exception('Could not access destination directory. Reason: {0}'.format(e))
			if f[-1] != '/':
				f += '/'
			probar = progressbar.ProgressBar()
			try:
				print('Starting Search. Depending on the amount of subdirectories, this may take a while...')
				b, c, d = path.getall(a, True, ['pdf'], progressBar = progressbar.ProgressBar(), onerror = recordExceptions)
				handleWalkExceptions(c, b)
			except Exception as e:
				raise Exception('Could not search directory. Reason: {0}'.format(e))
			for x in probar(b):
				tryToCopy(x, f + x.split('/')[-1])

			os.remove('Session filenames.out')
			os.remove('Session paths.out')
		except:
			raise

def gTest():
	a = tk.Tk()
	Graphical(a, a).pack()
	a.mainloop()



class Graphical(tk.Frame):
	def __init__(self, master, TkInstance, *args, **kwargs):
		# Setup graphics and variables
		ttk.Frame.__init__(self, master, *args, **kwargs)
		self.__Tk = TkInstance
		self.__canceler = canceler.Canceler()
		self.__requests = []
		self.__style = ttk.Style()
		self.__currFile = tk.StringVar(self, 'Current File:\tNone')
		self.__stateVar = tk.StringVar(self, 'Status:\t\tReady')
		self.__selection = tk.StringVar(self, '')
		self.__dest = tk.StringVar(self, '')
		self.__extVar = tk.StringVar(self, 'pdf')
		self.__l1 = ttk.Label(self, text = 'Path of input file or directory:')
		self.__selectEntry = ttk.Entry(self, width = 80, textvariable = self.__selection)
		self.__sFolderButton = ttk.Button(self, text = 'Browse...', command = self.getSFolder)
		self.__l2 = ttk.Label(self, text = 'Path of output directory:')
		self.__destEntry = ttk.Entry(self, width = 80, textvariable = self.__dest)
		self.__dFolderButton = ttk.Button(self, text = 'Browse...', command = self.getDFolder)
		self.__extLabel = ttk.Label(self, text = "File type to search for")
		self.__extBox = ttk.Entry(self, width = 20, textvariable = self.__extVar)
		self.__controlFrame = ttk.Frame(self)
		self.__start = ttk.Button(self.__controlFrame, text = 'Start', command = self.thread)
		self.__cancel = ttk.Button(self.__controlFrame, text = 'Cancel', command = self.cancel)
		self.__progress = ttk.Progressbar(self, value = 100)
		self.__stateLabel = ttk.Label(self, textvariable = self.__stateVar)
		self.__sizeL = ['Consolas', 12]
		self.__style.configure('.', font = self.__sizeL)
		self.__selectEntry.configure(font = self.__sizeL)
		self.__destEntry.configure(font = self.__sizeL)
		# Position graphics
		self.__l1.grid(column = 0, row = 0, sticky = 'w')
		self.__selectEntry.grid(column = 0, row = 1, sticky = 'ew')
		self.__fileButton.grid(column = 1, row = 1, sticky = 'ew')
		self.__sFolderButton.grid(column = 2, row = 1, sticky = 'ew')
		self.__l2.grid(column = 0, row = 2, sticky = 'w')
		self.__destEntry.grid(column = 0, row = 3, sticky = 'ew')
		self.__deriveButton.grid(column = 1, row = 3, sticky = 'ew')
		self.__dFolderButton.grid(column = 2, row = 3, sticky = 'ew')
		self.__extLabel.grid(column = 0, row = 4, sticky = 'w')
		self.__extBox.grid(column = 0, row = 5, sticky = 'w')
		self.__controlFrame.grid(column = 0, row = 6, columnspan = 3, sticky = 'ew')
		self.__start.pack(side = 'left', expand = True, fill = 'x')
		self.__cancel.pack(side = 'right', expand = True, fill = 'x')
		self.__progress.grid(column = 0, row = 7, columnspan = 3, sticky = 'ew')
		self.__stateLabel.grid(column = 0, row = 8, sticky = 'w')
		self.bind_all('<Control-plus>', self.inc)
		self.bind_all('<Control-equal>', self.inc)
		self.bind_all('<Control-minus>', self.dec)
		self.progress_ready()
		self.after(100, self.handle_loop)

	def getSFolder(self):
		self.__selection.set(tkFile.askdirectory(self.__Tk))

	def getDFolder(self):
		self.__dest.set(tkFile.askdirectory(self.__Tk))

	def progress_ready(self):
		self.__stateVar.set('Status:\t\tReady')
		self.__currFile.set('Current File:\tNone')
		try:
			self.__progress.stop()
		except:
			pass;
		self.__progress.configure(maximum = 100, value = 100, mode = 'determinate')
		self.__selectEntry.configure(state = tk.NORMAL)
		self.__sFolderButton.configure(state = tk.NORMAL)
		self.__start.configure(state = tk.NORMAL)
		self.__destEntry.configure(state = tk.NORMAL)
		self.__dFolderButton.configure(state = tk.NORMAL)
		self.__cancel.configure(state = tk.DISABLED)

	def progress_indeterminate(self):
		self.__progress.configure(mode = 'indeterminate', maximum = 100, value = 0)
		self.__progress.start(15)

	def cancel(self, event = None):
		if tkmb.askyesno('Cancel', 'Are you sure you want to cancel the current operation?'):
			self.__canceler.cancel()
			self.__stateVar.set('Status:\t\tCanceling...')

	def initEnv(self):
		self.__c = False
		self.__selectEntry.configure(state = tk.DISABLED)
		self.__fileButton.configure(state = tk.DISABLED)
		self.__sFolderButton.configure(state = tk.DISABLED)
		self.__start.configure(state = tk.DISABLED)
		self.__destEntry.configure(state = tk.DISABLED)
		self.__deriveButton.configure(state = tk.DISABLED)
		self.__dFolderButton.configure(state = tk.DISABLED)
		self.__cancel.configure(state = tk.NORMAL)

	def thread(self):
		threading.Thread(target = Graphical.search, args = (self,)).start()

	def request(self, *args):
		self.__requests.append(args)

	def handle_loop(self):
		if len(self.__requests) > 0:
			a = self.__requests.pop(0)
			if len(a) > 1:
				a[0](*a[1:])
			else:
				a[0]()
		self.after(100, self.handle_loop)

	def inc(self, event = None):
		self.__sizeL[1] += 1
		self.__style.configure('.', font = self.__sizeL)
		self.__selectEntry.configure(font = self.__sizeL)
		self.__destEntry.configure(font = self.__sizeL)

	def dec(self, event = None):
		if self.__sizeL[1] > 1:
			self.__sizeL[1] -= 1
			self.__style.configure('.', font = self.__sizeL)
			self.__selectEntry.configure(font = self.__sizeL)
			self.__destEntry.configure(font = self.__sizeL)

	def search(self, event = None):
		self.__stateVar.set('Status:\t\tInitializing Environment...')
		self.progress_indeterminate()
		self.initEnv()
		if self.__canceler.get():
			self.progress_ready()
			return
		self.__dest.set(self.__destEntry.get())
		self.__selection.set(self.__selectEntry.get())
		a = self.__selection.get().replace('"', '').replace("'", '')
		f = self.__dest.get().replace('\\', '/').replace('"', '').replace("'", '')
		try:
			os.stat(a)
		except Exception as e:
			self.request(tkmb.showerror, 'ERROR', 'Could not access input directory. Reason: {0}'.format(e))
			self.progress_ready()
			return
		try:
			os.stat(f)
		except:
			try:
				os.mkdir(f)
			except Exception as e:
				self.request(tkmb.showerror, 'ERROR', 'Could not access or create destination directory. Reason: {0}'.format(e))
				self.progress_ready()
				return
		if f[-1] != '/':
			f += '/'
		self.__stateVar.set('Status:\t\tSearching for {0} files...'.format(self.__extBox.get()))
		b, c, d = path.getall(a, True, [self.__extBox.get()], onerror = recordExceptions)
		handleWalkExceptions(c, b)

if __name__ == '__main__':
	start()
