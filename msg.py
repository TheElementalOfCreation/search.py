# Note: In python, a "#" indicates the start of a comment and text enclosed in triple quotes
# is a comment that can be retrieved from the python command line with help(command).
# For example, if you are trying to get help for the command "test()" in the script "pro.py"
# you would first import the script with "import pro" and then use the command
# "help(pro.test)".

# ----Start importing necessary libraries---- #

import base64
import os
import re
import sys
import shutil
import subprocess
import threading
import time
import tkinter as tk
import zipfile

import bs4 as BeautifulSoup
import compressed_rtf
import extract_msg
import progressbar
import tkFile


from tkinter import ttk, messagebox as tkmb

from creatorUtils import canceler
from creatorUtils.path import * # "from path import *" means to import all functionas and varibles from "path.py"

# ----End importing necessary libraries---- #

# ----Early define globals---- #
# These variables MUST be defined before any of the functions because some function defaults depend on these
ex = None
# Determine the directory for the other resources.
# This MUST be the first global variable in this list.
# "__file__" is a variable in every python script that
# contains the path of that file.
resDir = get_short_path_name('/'.join(os.path.abspath(__file__).split('\\')[0:-1]))
# Define path to error folder
error = get_short_path_name(resDir + '/ERROR')

#----Define Attachment class----#
class Attachment(extract_msg.Attachment):
    def __init__(self, msg, dir_):
        extract_msg.Attachment.__init__(self, msg, dir_)
    def saveEmbededMessage(self, contentId = False, json = False, useFileName = False, raw = False, customPath = None, customFilename = None):
        processFile(self.msg.path, ''.join(i for i in self.msg.filename.split('/').pop().split('.')[0] if i not in r'\/:*?"<>|'), self.__prefix)


#----Start defining necessary functions----#

def pause():
    raw_input('Press enter to continue...')

def zipFolder(folderpath, outfilepath):
    z = zipfile.ZipFile(outfilepath, 'w')
    if not isinstance(folderpath, str):
        folderpath = folderpath.decode('utf-8')
    paths = []
    for x in os.walk(folderpath):
        for y in x[1]:
            paths.append(x[0] + '/' + y)
        for y in x[2]:
            paths.append(x[0] + '/' + y)
    paths.sort()
    for x in paths:
        z.write(x, x)
    z.close()

def callNode(filename, filepath, foldername, outfile = error + '/node.out'):
    """
    Calls the Node.js Script that deencapsulates
    the html embeded in a .rtf file. The Node.js
    script tries to read a file called "out.rtf"
    and writes the html code to "output.html".

    Ensure that the current working directory
    contains "out.rtf" before calling this. If
    you are running this as a script, this will
    be done automatically.
    """
    fil1 = open(outfile, 'w')
    a = subprocess.call([nodePath, rtf2htmlPath], stdout = fil1, stderr = fil1)
    fil1.close()
    if a != 0:
        raise NodeError(filename, filepath, foldername, outfile)

def getHeader(messageObject):
    """
    Gets the necessary header information from the
    Message object
    """
    # Dictionary that will contain the header of the output file.
    header = {'From:': '', 'To:': '', 'Sent:': '', 'Cc:': '', 'Subject:': ''}
    header['From:'] = messageObject.sender
    header['Sent:'] = messageObject.date
    header['To:'] = messageObject.to
    header['Cc:'] = messageObject.cc
    header['Subject:'] = messageObject.subject
    return header

def getFormattedHeader(header):
    """
    Formats the header and saves the values to the dictionary "formatted".
    """
    # Dictionary that will contain the formatted header strings.
    formatted = {'From:': '', 'To:': '', 'Sent:': '', 'Cc:': '', 'Subject:': ''}
    for x in headerVals:
        a = header[x]
        if x == 'Cc:' or x == 'To:':
            if a != None:
                while True:
                    if a.find('<') == -1:
                        break
                    a = a.replace(a[a.find('<'):a.find('>')+1], '')
                a = a.replace('\r\n\t', '')
                a = a.replace(' , ', ', ')
                a = a.replace(',', ';')
                a = a.replace(' ;', ';')
        if x == 'Sent:':
            try:
                b = a.split(' ')
                b[0] = dayOfWeek[(b[0])]
                b[2] = month[b[2]]
                if int(b[4][0:2]) <= 12:
                    g = 'AM'
                else:
                    g = 'PM'
                b[4] = str(int(b[4][0:2])%12) + b[4][2:5]
                c = [b[0], b[2], b[1] + ',', b[3], b[4], g, 'GMT', b[5]]
                a = ' '.join(c)
            except:
                print('Message date error! Setting sent date as "Null"...')
                a = 'Null'
        if a == None:
            a = '<b>' + x + '</b>&nbsp;<br>'
        else:
            a = '<b>' + x + '</b>&nbsp;' + a + '<br>'
        formatted[x] = a
    return formatted

def callToPdf(name):
    """
    Calls to the program to convert the html file to a pdf file.
    """
    subprocess.call([toPdf, '-q', '-O', 'Landscape', 'output.html', name + ' message.pdf'])
    # wkhtmltopdf.wkhtmltopdf(*['output.html', os.getcwd() + '\\' + name + ' message.pdf'], **{'Orientation' : 'Landscape'})

def mkdir(name):
    """
    Tries to make a directory called "name". If it fails,
    it will add " (#)", where "#" is a number, to the name
    and try again up to 20. If it succedes, it returns the
    name of the directory it created.
    """
    try:
        os.mkdir(name)
        return name
    except:
        name += ' ({})'
        for x in range(2, 20):
            try:
                os.mkdir(name.format(x))
                return name.format(x)
            except Exception:
                if x == 19:
                    raise

def addHeader(headerString):
    """
    This function adds the formatted header
    to the html file before it is converted.
    """
    htmlFileIn = open('output.html', 'rb')
    text = htmlFileIn.read()
    htmlFileIn.close()
    text = text.replace(b'vlink="#954F72"><div class=WordSection1>', b'vlink="#954F72"><div class=WordSection1>' + headerString.encode('utf-8'))
    htmlFileOut = open('output.html', 'wb')
    htmlFileOut.write(text)
    htmlFileOut.close()

def embedImages():
    if debug:
        shutil.copy('output.html', 'outputOLD.html')
    file1 = open('output.html', 'rb')
    con = file1.read()
    file1.close()
    try:
        bs = BeautifulSoup.BeautifulSoup(con.decode('utf8'), features = "lxml")
    except UnicodeDecodeError:
        bs = BeautifulSoup.BeautifulSoup(con, fromEncoding = 'windows-1252', features = "lxml")
    if bs.find('meta', {'http-equiv': 'Content-Type'}) is None:
        bs.find('head').insert(1, BeautifulSoup.Tag(parser = bs, name = 'meta', attrs = {'http-equiv': 'Content-Type', 'content': 'text/html; charset=utf-8'}))
    tagsTemp = bs.findAll('img')
    tags = []
    for x in tagsTemp:
        h = x.get('src')
        if h != 'None':
            if len(h) > 3:
                if h[0:4] == 'cid:':
                    tags.append(x)
    for x in tags:
        src = x['src']
        with open(src[4:], 'rb') as emb:
            stream = emb.read()
        data = 'data:image;base64,' + base64.b64encode(stream).decode('utf8')
        x['src'] = data
        if not debug:
            os.remove(src[4:])
    con = bs.prettify()
    file1 = open('output.html', 'wb')
    file1.write(con.encode('utf8'))
    file1.close()

def processFile(filepath, filename, prefix = '', _canceler = canceler.FAKE):
    messa = extract_msg.Message(filepath, prefix, Attachment) # Create Message object.
    header = getHeader(messa) # Gets the header.
    formatted = getFormattedHeader(header) # The next few lines format the header
    headerString = '<p class=MsoNormal>'
    for u in headerVals:
        headerString = headerString + formatted[u]
    headerString = headerString + '</p><br>'
    ofilename = filename
    filename = mkdir(filename)
    os.chdir(filename)
    messa.save_attachments(True) # Saves the attatchments
    if messa.htmlBody != None: # Has html body?
        with open('output.html', 'wb') as o:
            o.write(messa.htmlBody)
    else:
        rtfContents = compressed_rtf.decompress(messa.compressedRtf) # Read contents, decompress them, and store them in a variable.
        with open('out.rtf', 'wb') as rtfFile:
            rtfFile.write(rtfContents)
        callNode(ofilename, filepath, filename)
        if not debug:
            os.remove('out.rtf')
    addHeader(headerString)
    embedImages()
    callToPdf(filename)
    if not debug:
        os.remove('output.html')
    os.chdir('..') # Move back to the parent directory
    # Note, if the current directory is "C:\path\to\current\directory\" then ".."
    # refers to the parent directory, "C:\path\to\current\" in this case.

def gTest():
    a = tk.Tk()
    Graphical(a, a).pack()
    a.mainloop()

# ----End Defining functions---- #

# ----Start defining global variables---- #

# Are we trying to collect debug information?
with open(resDir + '/debug/debug.txt', 'r') as f:
    debug = f.read()[0] != '0'
# Define the path for the node.js interpreter.
nodePath = get_short_path_name(resDir + '/nodejs/node.exe')
# Define the path for the rtf2html script.
rtf2htmlPath = get_short_path_name(resDir + '/nodejs/rtf2html.js')
# Define the path for the output directory.
outPath = ''
# Tuple (unchangable array) that contains the index names for header values
# in the order they will be retrieved.
headerVals = (
    'From:',
    'Sent:',
    'To:',
    'Cc:',
    'Subject:',
)
# Dictionary of short day names to long day names.
dayOfWeek = {
    'Sun,': 'Sunday,',
    'Mon,': 'Monday,',
    'Tue,': 'Tuesday,',
    'Wed,': 'Wednesday,',
    'Thu,': 'Thursday,',
    'Fri,': 'Friday,',
    'Sat,': 'Saturday,',
}
# Dictionary of short month names to long month names.
month = {
    'Jan': 'January',
    'Feb': 'February',
    'Mar': 'March',
    'Apr': 'April',
    'May': 'May',
    'Jun': 'June',
    'June': 'June',
    'Jul': 'July',
    'July': 'July',
    'Aug': 'August',
    'Sep': 'September',
    'Oct': 'October',
    'Nov': 'November',
    'Dec': 'December',
}
# Defines the path to program to convert html to pdf.
toPdf = get_short_path_name(resDir + '/wkhtmltopdf/wkhtmltopdf.exe')
# The next few lines are defining specific exceptions so they can be used later.
winExceptions = [
    WindowsError(2, 'Cannot find file.'),
    WindowsError(3, 'Cannot find path.'),
    WindowsError(5, 'Permission denied.'),
    WindowsError(15, 'Cannot find drive.'),
    WindowsError(17, 'Cannot copy file across drives.'),
    WindowsError(19, 'Location is write protected.'),
    WindowsError(32, 'File or directory busy.'),
    WindowsError(183, 'File already Exists.'),
]

# ----End defining global variables---- #

# ----Define NodeError class---- #
class NodeError(Exception):
    """
    Created when a class to Node.js fails. `out`
    is the path of the error file that was generated
    by the last call to Node.js. `filename` and
    `filepath` are the name and path, respectively,
    of the file that was being worked on when the
    error occured. `foldername` is the name of the
    folder in the temp directory where the file was
    being worked on.
    """
    def __init__(self, filename, filepath, foldername, out = error + '/node.out'):
        with open(out, 'r') as o:
            self.__nodeErrorLog = o.read()
            o.close()
        self.__filename = filename
        self.__foldername = foldername
        self.__filepath = filepath
        self.__message = ('\n\n' + '-'*40 + '\n!CRITICAL!: Node.js ran into an error.\nFile: {}\nPath: {}\n\n The log of this error is as follows:\n{}\n\n' + '-'*20 + '\n').format(self.filename, self.filepath, self.nodeErrorLog)
        Exception.__init__(self, self.__message)

    @property
    def nodeErrorLog(self):
        return self.__nodeErrorLog

    @property
    def filepath(self):
        return self.__filepath

    @property
    def filename(self):
        return self.__filename

    @property
    def foldername(self):
        return self.__foldername

    @property
    def message(self):
        return self.__message



# ----Define graphical interface classes---- #

class Graphical(ttk.Frame):
    def __init__(self, master, TkInstance, initialInput = None, initialOutput = None, **kwargs):
        # Setup graphics and variables
        ttk.Frame.__init__(self, master, **kwargs)
        self.__Tk = TkInstance
        self.__canceler = canceler.Canceler()
        self.__style = ttk.Style()
        self.__sizeL = ['Consolas', 12]
        self.__requests = []
        self.__style.configure('.', font = self.__sizeL)
        self.__currFile = tk.StringVar(self, 'Current File:\tNone')
        self.__stateVar = tk.StringVar(self, 'Status:\t\tReady')
        self.__selection = tk.StringVar(self, initialInput if isinstance(initialInput, stringType) else '')
        self.__dest = tk.StringVar(self, initialOutput if isinstance(initialOutput, stringType) else '')
        self.__l1 = ttk.Label(self, text = 'Path of input file or directory:')
        self.__selectEntry = ttk.Entry(self, width = 80, textvariable = self.__selection)
        self.__fileButton = ttk.Button(self, text = 'Browse for file', command = self.getSFile)
        self.__sFolderButton = ttk.Button(self, text = 'Browse for folder', command = self.getSFolder)
        self.__l2 = ttk.Label(self, text = 'Path of output directory:')
        self.__destEntry = ttk.Entry(self, width = 80, textvariable = self.__dest)
        self.__deriveButton = ttk.Button(self, text = 'Derive from input', command = self.derive)
        self.__dFolderButton = ttk.Button(self, text = 'Browse for folder', command = self.getDFolder)
        self.__controlFrame = ttk.Frame(self)
        self.__start = ttk.Button(self.__controlFrame, text = 'Start', command = self.processFiles)
        self.__cancel = ttk.Button(self.__controlFrame, text = 'Cancel', command = self.cancel)
        self.__progress = ttk.Progressbar(self, value = 100)
        self.__stateLabel = ttk.Label(self, textvariable = self.__stateVar)
        self.__currentLabel = ttk.Label(self, textvariable = self.__currFile)
        # Position graphics
        self.__l1.grid(column = 0, row = 0, sticky = 'w')
        self.__selectEntry.grid(column = 0, row = 1, sticky = 'ew')
        self.__fileButton.grid(column = 1, row = 1, sticky = 'ew')
        self.__sFolderButton.grid(column = 2, row = 1, sticky = 'ew')
        self.__l2.grid(column = 0, row = 2, sticky = 'w')
        self.__destEntry.grid(column = 0, row = 3, sticky = 'ew')
        self.__deriveButton.grid(column = 1, row = 3, sticky = 'ew')
        self.__dFolderButton.grid(column = 2, row = 3, sticky = 'ew')
        self.__controlFrame.grid(column = 0, row = 4, columnspan = 3, sticky = 'ew')
        self.__start.pack(side = 'left', expand = True, fill = 'x')
        self.__cancel.pack(side = 'right', expand = True, fill = 'x')
        self.__progress.grid(column = 0, row = 5, columnspan = 3, sticky = 'ew')
        self.__stateLabel.grid(column = 0, row = 6, sticky = 'w')
        self.__currentLabel.grid(column = 0, row = 7, sticky = 'w')
        self.bind_all('<Control-plus>', self.inc)
        self.bind_all('<Control-equal>', self.inc)
        self.bind_all('<Control-minus>', self.dec)
        self.after(100, self.handle_loop)
        self.progress_ready()

    def getSFile(self):
        self.__selection.set(tkFile.askopenfilename(self.__Tk))

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
            pass
        self.__progress.configure(maximum = 100, value = 100, mode = 'determinate')
        self.__selectEntry.configure(state = tk.NORMAL)
        self.__fileButton.configure(state = tk.NORMAL)
        self.__sFolderButton.configure(state = tk.NORMAL)
        self.__start.configure(state = tk.NORMAL)
        self.__destEntry.configure(state = tk.NORMAL)
        self.__deriveButton.configure(state = tk.NORMAL)
        self.__dFolderButton.configure(state = tk.NORMAL)
        self.__cancel.configure(state = tk.DISABLED)

    def progress_indeterminate(self):
        self.__progress.configure(mode = 'indeterminate', maximum = 100, value = 0)
        self.__progress.start(15)

    def derive(self):
        a = self.__selectEntry.get().replace('\\', '/')
        if a[-1] == '/':
            a = a[:-1]
        try:
            os.stat(a)
        except:
            return
        b = '/'.join(a.split('/')[:-1])
        self.__dest.set(b)

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
        os.chdir(resDir)
        try:
            os.mkdir('temp') # Create temporary working directory.
        except: # If "temp" already exists in the resource directory, an exception will have been raised.
            shutil.rmtree(resDir + '/temp')
            time.sleep(1)
            os.mkdir('temp')
        os.chdir('temp')
        self.__root = os.getcwd()

    def thread(self):
        threading.Thread(target = Graphical.processFiles, args = (self,)).start()

    def processFiles(self):
        self.__stateVar.set('Status:\t\tInitializing Environment...')
        self.progress_indeterminate()
        self.initEnv()
        if self.__canceler.get():
            self.progress_ready()
            return
        a = self.__selection.get().replace('"', '').replace("'", '')
        if a == '':
            self.request(tkmb.showerror, 'ERROR', 'Please specify an input.')
            self.progress_ready()
            return
        f = self.__dest.get().replace('\\', '/').replace('"', '').replace("'", '')
        if f == '':
            self.request(tkmb.showerror, 'ERROR', 'Please specify an output.')
            self.progress_ready()
            return
        try:
            os.stat(a)
        except Exception as e:
            self.request(tkmb.showerror, 'ERROR', 'Could not access input directory or file. Reason: {0}'.format(e))
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
        global outPath
        outPath = f + 'out'
        self.__stateVar.set('Status:\t\tSearching for msg files...')
        files, filenames, cons = getall(a)
        if self.__canceler.get():
            self.progress_ready()
            return
        if len(files) == 0:
            self.request(tkmb.showerror, 'ERROR', 'Could not find any ".msg" files.')
            self.progress_ready()
            return
        flen = len(files)
        self.__progress.stop()
        self.__progress.configure(maximum = flen, value = 0, mode = 'determinate')
        self.__stateVar.set('Status:\t\tConverting Files...')
        errors = []
        had_error = False
        for x in range(flen):
            if self.__canceler.get():
                self.progress_ready()
                return
            self.__currFile.set('Current File:\t' + filenames[x])
            self.__progress.step()
            try:
                processFile(files[x], filenames[x], _canceler = self.__canceler)
            except NodeError as e:
                had_error = True
                self.__stateVar.set('Status:\t\tHandling Error...')
                self.progress_indeterminate()
                os.chdir(self.__root + '/' + e.foldername)
                shutil.copy(error + '/node.out', './node.out')
                os.chdir(self.__root)
                shutil.copy(files[x], './' + filenames[x] + '.msg')
            except Exception as e:
                # Handle unexpected exception gracefully
                self.request(tkmb.showerror, 'EXCEPTION', '{}: {}'.format(type(e), e))
                pass #TODO#TODO#TODO#TODO#TODO
                had_error = True
                global ex
                ex = e
                self.__stateVar.set('Status:\t\tHandling Error...')
                self.progress_indeterminate()
                os.chdir(self.__root)
                shutil.copy(files[x], './' + filenames[x] + '.msg')
        os.chdir(resDir)
        if self.__canceler.get():
            self.progress_ready()
            return
        if had_error:
            p = resDir + '/Send to developer/' + time.strftime("%Y_%m_%d-%H.%M.%S") + '.zip'
            zipFolder(self.__root, p)
            self.request(tkmb.showerror, 'ERROR', 'Major error occurred during conversion. Please send the file "' + p + '" to the developer.')
            self.progress_ready()
        else:
            self.__stateVar.set('Status:\t\tMoving files to out directory...')
            self.__currFile.set('Current File:\tNone')
            self.progress_indeterminate()
            if get_short_path_name(outPath) == '':
                shutil.move('temp', outPath)
            else:
                os.chdir(resDir + '/temp')
                for fil in os.listdir(u'.'):
                    if get_short_path_name(outPath + '/' + fil) == '':
                        try:
                            shutil.move(fil, outPath + '/' + fil)
                        except: #TODO once the too long filename exception is properly raised, write the handling code
                            pass
                    else:
                        errors.append('\tDirectory ' + fil + ' already exists in out directory.\n')
            if len(errors) > 0:
                err = 'The following errors occured while copying the msg files:\n'
                for e in errors:
                    err += e
                err += 'A log file of these errors has been created at '
                err += get_long_path_name(resDir)
                err += '\\log.txt\nIf any of these are not true, please report this to the developer.'
                with open(resDir + '/log.txt', 'a+') as logfile:
                    logfile.write(err)
                self.request(tkmb.showerror, 'ERROR', err)
                self.progress_ready()

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

# If the script is called directly from the interpreter
# (as opposed to being imported by another script) the
# following script will be called.
if __name__ == '__main__':
    while True:
        os.chdir(resDir) # Change current directory to the resource directory.
        os.system('cls')
        print('Creating temp directory...')
        try:
            os.mkdir('temp') # Create temporary working directory.
        except: # If "temp" already exists in the resource directory, an exception will have been raised.
            # delFolder(resDir + '/temp') # If temp already exists, delete it and its contents.
            shutil.rmtree(resDir + '/temp')
            time.sleep(3)
            os.mkdir('temp')
        os.chdir('temp') # Enter temp directory.
        root = os.getcwd()
        os.system('cls')
        try:
            os.mkdir('C:\Msg')
        except:
            pass
        if sys.argv[1] == '':
            a = raw_input('Enter folder path you want to search, or the file you want to convert (You may just drag the folder onto this window):\n')
            a = a.replace('"', '').replace("'", '')
            argum = a
            f = raw_input('Enter destination folder (Again, you can drag and drop):\n')
            f = f.replace('\\', '/').replace('"', '').replace("'", '')
            try:
                os.stat(f)
            except Exception as e:
                raise Exception('Could not access destination directory. Reason: {0}'.format(e))
            if f[-1] != '/':
                f += '/'
        else:
            argum = sys.argv[1]
        print('Searching for msg files...')
        files, filenames, outPath = getall(argum) # Search for msg files.
        if sys.argv[1] == '':
            outPath = f
        else:
            outPath += '/out'
        if len(files) == 0: # If no msg file found, raise exception.
            raise BaseException('No msg files in directory')
        os.system('cls')
        print('Converting files...')
        proBar = progressbar.ProgressBar()
        fLen = len(files)
        errors = []
        for x in proBar(range(fLen)):
            print(' Converting file: ' + filenames[x])
            try:
                processFile(files[x], filenames[x])
            except NodeError as e:
                errors.append(e.message)
                shutil.copy('C:/Msg/node.out', './node.out')
                shutil.copy(files[x], './' + filenames[x] + '.msg')
                os.chdir(root + '/' + e.foldername)
                t = os.listdir(u'.')
                os.chdir(root)
                zipFolder(root + '/' + e.foldername, root + '/Send to developer/' + e.foldername + '.zip')
            except IOError as e:
                if e.message == 'not an OLE2 structured storage file':
                    errors.append('File "{}" is not recognised as a proper msg file.'.format(files[x]))
                else:
                    raise
            os.system('cls')
            print('Converting files...')
        os.chdir('..')
        os.system('cls')
        print('Moving files to out directory...')
        if get_short_path_name(outPath) == '':
            shutil.move('temp', outPath)
        else:
            os.chdir(resDir + '/temp')
            for fil in os.listdir('.'):
                if get_short_path_name(outPath + '/' + fil) == '':
                    shutil.move(fil, outPath + '/' + fil)
                else:
                    errors.append('\tDirectory ' + fil + ' already exists in out directory.')
        if len(errors) > 0:
            logfile = open(resDir + '/log.txt', 'a+')
            print('The following errors occured while converting the msg files:')
            logfile.write('The following errors occured while converting the msg files:')
            for e in errors:
                print(e)
                logfile.write(e)
            print('A log file of these errors has been created at ' + get_long_path_name(resDir) + '\\log.txt')
            print('If any of these are not true, please report this to the developer.')
            logfile.write('If any of these are not true, please report this to the developer.')
            logfile.write('\n\n')
            logfile.close()
            if sys.argv[1] == '':
                pause()
        if sys.argv[1] != '':
            break
