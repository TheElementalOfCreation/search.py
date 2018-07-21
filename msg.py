#Note: In python, a "#" indicates the start of a comment and text enclosed in triple quotes
#is a comment that can be retrieved from the python command line with help(command).
#For example, if you are trying to get help for the command "test()" in the script "pro.py"
#you would first import the script with "import pro" and then use the command
#"help(pro.test)".

#----Start importing necessary libraries----#

import sys;
import os;
import re;
import base64;
from subprocess import call;
import compressed_rtf as LZFu;
import ExtractMsg;
import shutil;
import progressbar;
import time;
import BeautifulSoup;
import exceptions;
import zipfile;
import httplib;
from creatorUtils.path import *; #"from path import *" means to import all functionas and varibles from "path.py"

#----End importing necessary libraries----#

#----Define Attachment class----#
class Attachment(ExtractMsg.Attachment):
	def __init__(self, msg, dir_):
		ExtractMsg.Attachment.__init__(self, msg, dir_);
	def saveEmbededMessage(self, contentId = False, json = False, useFileName = False, raw = False):
		processFile(self.msg.path, ''.join(i for i in self.msg.filename.split('/').pop().split('.')[0] if i not in r'\/:*?"<>|'), self.__prefix)

#----Define NodeError class----#
class NodeError(exceptions.Exception):
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
	def __init__(self, filename, filepath, foldername, out = 'C:/Msg/node.out'):
		o  = open(out, 'r');
		self.__nodeErrorLog = o.read();
		o.close();
		self.__filename = filename;
		self.__foldername = foldername;
		self.__filepath = filepath;
		self.__message = ('\n\n' + '-'*40 + '\n!CRITICAL!: Node.js ran into an error.\nFile: {}\nPath: {}\n\n The log of this error is as follows:\n{}\n\n' + '-'*20 + '\n').format(self.filename, self.filepath, self.nodeErrorLog);
		Exception.__init__(self, self.__message);

	@property
	def nodeErrorLog(self):
		return self.__nodeErrorLog;

	@property
	def filepath(self):
		return self.__filepath;

	@property
	def filename(self):
		return self.__filename;

	@property
	def foldername(self):
		return self.__foldername;

	@property
	def message(self):
		return self.__message;

#----Start defining necessary functions----#

def pause():
	raw_input('Press enter to continue...');

def zipFolder(folderpath, outfilepath):
	z = zipfile.ZipFile(outfilepath, 'w');
	paths = [];
	for x in os.walk(folderpath):
		for y in x[1]:
			paths.append(x[0] + '/' + y);
		for y in x[2]:
			paths.append(x[0] + '/' + y);
	paths.sort();
	for x in paths:
		z.write(x, x);
	z.close();

def callNode(filename, filepath, foldername, outfile = 'C:/Msg/node.out'):
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
	fil1 = open(outfile, 'w');
	a = call([nodePath, rtf2htmlPath], stdout = fil1, stderr = fil1);
	fil1.close();
	if a != 0:
		raise NodeError(filename, filepath, foldername, outfile);

def getHeader(messageObject):
	"""
	Gets the necessary header information from the
	Message object
	"""
	#Dictionary that will contain the header of the output file.
	header = {'From:':'', 'To:':'', 'Sent:':'', 'Cc:':'', 'Subject:':''};
	header['From:'] = messageObject.sender;
	header['Sent:'] = messageObject.date;
	header['To:'] = messageObject.to;
	header['Cc:'] = messageObject.cc;
	header['Subject:'] = messageObject.subject;
	return header;

def getFormattedHeader(header):
	"""
	Formats the header and saves the values to the dictionary "formatted".
	"""
	#Dictionary that will contain the formatted header strings.
	formatted = {'From:':'', 'To:':'', 'Sent:':'', 'Cc:':'', 'Subject:':''};
	for x in headerVals:
		a = header[x];
		if x == 'Cc:' or x == 'To:':
			if a != None:
				while True:
					if a.find('<') == -1:
						break;
					a = a.replace(a[a.find('<'):a.find('>')+1], '');
				a = a.replace('\r\n\t', '');
				a = a.replace(' , ', ', ');
				a = a.replace(',', ';');
				a = a.replace(' ;', ';');
		if x == 'Sent:':
			try:
				b = a.split(' ');
				b[0] = dayOfWeek[(b[0])];
				b[2] = month[b[2]];
				if int(b[4][0:2]) <= 12:
					g = 'AM';
				else:
					g = 'PM';
				b[4] = str(int(b[4][0:2])%12) + b[4][2:5];
				c = [b[0], b[2], b[1] + ',', b[3], b[4], g, 'GMT', b[5]];
				a = ' '.join(c);
			except:
				print('Message date error! Setting sent date as "Null"...');
				a = 'Null';
		if a == None:
			a = '<b>' + x + '</b>&nbsp;<br>';
		else:
			a = '<b>' + x + '</b>&nbsp;' + a + '<br>';
		formatted[x] = a;
	return formatted;

def callToPdf(name):
	"""
	Calls to the program to convert the html file to a pdf file.
	"""
	call([toPdf, '-q', '-O', 'Landscape', 'output.html', name + ' message.pdf']);
	#wkhtmltopdf.wkhtmltopdf(*['output.html', os.getcwd() + '\\' + name + ' message.pdf'], **{'Orientation' : 'Landscape'});

def mkdir(name):
	"""
	Tries to make a directory called "name". If it fails,
	it will add " (#)", where "#" is a number, to the name
	and try again up to 20. If it succedes, it returns the
	name of the directory it created.
	"""
	try:
		os.mkdir(name);
		return name;
	except:
		name += ' ({})';
		for x in range(2, 20):
			try:
				os.mkdir(name.format(x));
				return name.format(x);
			except Exception:
				if x == 19:
					raise;

def addHeader(headerString):
	"""
	This function adds the formatted header
	to the html file before it is converted.
	"""
	htmlFileIn = open('output.html','rb');
	text = htmlFileIn.read();
	htmlFileIn.close();
	text = text.replace('vlink="#954F72"><div class=WordSection1>', 'vlink="#954F72"><div class=WordSection1>' + headerString.encode('iso-8859-1'));
	htmlFileOut = open('output.html','wb');
	htmlFileOut.write(text);
	htmlFileOut.close();

def embedImages():
	shutil.copy('output.html', 'outputOLD.html');
	file1 = open('output.html','rb');
	con = file1.read();
	file1.close();
	bs = BeautifulSoup.BeautifulSoup();
	try:
		bs.feed(con.decode('utf8'));
	except UnicodeDecodeError:
		bs = BeautifulSoup.BeautifulSoup(con, fromEncoding = 'windows-1252');
		#bs.feed(con.decode('iso-8859-1'));
	bs.goahead(True);
	if len(bs.fetch('meta', {'http-equiv': 'Content-Type'})) == 0:
		bs.fetch('head').insert(1, BeautifulSoup.Tag(bs, 'meta', {'http-equiv':'Content-Type', 'content': 'text/html; charset=utf-8'}));
	tagsTemp = bs.findAllPrevious('img');
	tags = [];
	for x in tagsTemp:
		h = x.get('src')
		if h != 'None':
			if len(h) > 3:
				if h[0:4] == 'cid:':
					tags.append(x);
	for x in tags:
		src = x['src'];
		emb = open(src[4:], 'rb');
		stream = emb.read();
		emb.close();
		data = 'data:image;base64,' + base64.b64encode(stream);
		x['src'] = data;
	con = bs.prettify();
	file1 = open('output.html','wb');
	file1.write(con);
	file1.close();

def processFile(filepath, filename, prefix = ''):
	messa = ExtractMsg.Message(filepath, prefix, Attachment); #Create Message object.
	header = getHeader(messa); #Gets the header.
	formatted = getFormattedHeader(header); #The next few lines format the header
	headerString = '<p class=MsoNormal>';
	for u in headerVals:
		headerString = headerString + formatted[u];
	headerString = headerString + '</p><br>';
	ofilename = filename;
	filename = mkdir(filename);
	os.chdir(filename);
	messa.save_attachments(True); #Saves the attatchments
	if messa.htmlBody != None: #Has html body?
		o = open("output.html", 'wb');
		o.write(messa.htmlBody);
		o.close();
	else:
		rtfContents = LZFu.decompress(messa.compressedRtf); #Read contents, decompress them, and store them in a variable.
		rtfFile = open('out.rtf','w');
		rtfFile.write(rtfContents);
		rtfFile.close();
		callNode(ofilename, filepath, filename);
		os.remove('out.rtf');
	addHeader(headerString);
	embedImages();
	callToPdf(filename);
	os.remove('output.html');
	os.chdir('..'); #Move back to the parent directory
	#Note, if the current directory is "C:\path\to\current\directory\" then ".."
	#refers to the parent directory, "C:\path\to\current\" in this case.

#----End Defining functions----#

#----Start defining global variables----#

# Determine the directory for the other resources.
# This MUST be the first global variable in this list.
# "__file__" is a variable in every python script that
# contains the path of that file.
resDir = get_short_path_name('/'.join(__file__.split('\\')[0:-1]));
# Define the path for the node.js interpreter.
nodePath = get_short_path_name(resDir + '/nodejs/node.exe');
# Define the path for the rtf2html script.
rtf2htmlPath = get_short_path_name(resDir + '/nodejs/rtf2html.js');
# Define the path for the output directory.
outPath = '';
# Tuple (unchangable array) that contains the index names for header values
# in the order they will be retrieved.
headerVals = ('From:','Sent:','To:','Cc:','Subject:')
# Dictionary of short day names to long day names.
dayOfWeek = {'Sun,': 'Sunday,', 'Mon,': 'Monday,', 'Tue,': 'Tuesday,', 'Wed,': 'Wednesday,', 'Thu,': 'Thursday,', 'Fri,': 'Friday,', 'Sat,': 'Saturday,'}
# Dictionary of short month names to long month names.
month = {'Jan': 'January', 'Feb': 'February', 'Mar': 'March', 'Apr': 'April', 'May': 'May', 'Jun': 'June', 'June': 'June', 'Jul': 'July', 'July': 'July', 'Aug': 'August', 'Sep': 'September', 'Oct': 'October', 'Nov': 'November', 'Dec': 'December'};
# Defines the path to program to convert html to pdf.
toPdf = get_short_path_name(resDir + '/wkhtmltopdf/wkhtmltopdf.exe');
# The next few lines are defining specific exceptions so they can be used later.
winExceptions = [
	WindowsError(2, 'Cannot find file.'),
	WindowsError(3, 'Cannot find path.'),
	WindowsError(5, 'Permission denied.'),
	WindowsError(15, 'Cannot find drive.'),
	WindowsError(17, 'Cannot copy file across drives.'),
	WindowsError(19, 'Location is write protected.'),
	WindowsError(32, 'File or directory busy.'),
	WindowsError(183, 'File already Exists.')
];

#----End defining global variables----#

# If the script is called directly from the interpreter
# (as opposed to being imported by another script) the
# following script will be called.
if __name__ == '__main__':
	while True:
		os.chdir(resDir); # Change current directory to the resource directory.
		os.system('cls');
		print('Creating temp directory...');
		try:
			os.mkdir('temp'); # Create temporary working directory.
		except: # If "temp" already exists in the resource directory, an exception will have been raised.
			# delFolder(resDir + '/temp'); # If temp already exists, delete it and its contents.
			shutil.rmtree(resDir + '/temp');
			time.sleep(3);
			os.mkdir('temp');
		os.chdir('temp'); # Enter temp directory.
		root = os.getcwd();
		os.system('cls');
		try:
			os.mkdir('C:\Msg');
		except:
			pass;
		if sys.argv[1] == '':
			a = raw_input('Enter folder path you want to search, or the file you want to convert (You may just drag the folder onto this window):\n');
			a = a.replace('"', '').replace("'", '');
			argum = a;
			f = raw_input('Enter destination folder (Again, you can drag and drop):\n');
			f = f.replace('\\', '/').replace('"', '').replace("'", '');
			try:
				os.stat(f);
			except Exception as e:
				raise Exception('Could not access destination directory. Reason: {0}'.format(e));
			if f[-1] != '/':
				f += '/';
		else:
			argum = sys.argv[1];
		print('Searching for msg files...');
		files, filenames, outPath = getall(argum); # Search for msg files.
		if sys.argv[1] == '':
			outPath = f;
		else:
			outPath += '/out';
		if len(files) == 0: # If no msg file found, rasie exception.
			raise BaseException('No msg files in directory');
		os.system('cls');
		print('Converting files...');
		proBar = progressbar.ProgressBar();
		fLen = len(files);
		errors = [];
		for x in proBar(range(fLen)):
			print(' Converting file: ' + filenames[x]);
			try:
				processFile(files[x], filenames[x]);
			except NodeError as e:
				errors.append(e.message);
				shutil.copy('C:/Msg/node.out', './node.out');
				shutil.copy(files[x], './' + filenames[x] + '.msg');
				os.chdir(root + '/' + e.foldername);
				t = os.listdir('.');
				os.chdir(root);
				zipFolder(root + '/' + e.foldername, root + '/Send to developer/' + e.foldername + '.zip');
			except IOError as e:
				if e.message == 'not an OLE2 structured storage file':
					errors.append('File "{}" is not recognised as a proper msg file.'.format(files[x]));
				else:
					raise;
			os.system('cls');
			print('Converting files...');
		os.chdir('..');
		os.system('cls');
		print('Moving files to out directory...');
		if get_short_path_name(outPath) == '':
			shutil.move('temp', outPath);
		else:
			os.chdir(resDir + '/temp');
			for fil in os.listdir('.'):
				if get_short_path_name(outPath + '/' + fil) == '':
					shutil.move(fil, outPath + '/' + fil);
				else:
					errors.append('\tDirectory ' + fil + ' already exists in out directory.');
		if len(errors) > 0:
			logfile = open(resDir + '/log.txt', 'a+');
			print('The following errors occured while converting the msg files:');
			logfile.write('The following errors occured while converting the msg files:');
			for e in errors:
				print(e);
				logfile.write(e);
			print('A log file of these errors has been created at ' + get_long_path_name(resDir) + '\\log.txt');
			print('If any of these are not true, please report this to the developer.');
			logfile.write('If any of these are not true, please report this to the developer.');
			logfile.write('\n\n');
			logfile.close();
			if sys.argv[1] == '':
				pause();
		if sys.argv[1] != '':
			break;
