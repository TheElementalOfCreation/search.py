import httplib;
import os;
import subprocess;

print('Initializing update program...');

def get(conn, x):
	conn.request('GET', url.format(x));
	data = conn.getresponse().read();
	file1 = open(x, 'wb');
	file1.write(data);
	file1.close();

files = [
	'msg.py',
	'toPdf.bat',
	'search.py',
	'SEARCH.bat',
	'Version',
];

url = '/TheElementalOfCreation/search.py/master/{}';
data = {};
version = (None, None);

try:
	a = open('Version', 'r');
	version = a.read().split('.');
	a.close();
except:
	pass;

print('Checking for new version...');

a = httplib.HTTPSConnection('raw.githubusercontent.com');
a.request('GET', url.format('Version'));
v = a.getresponse().read().split('.');
if v[0] != version[0]:
	# update.py and/or UPDATE.bat update
	print('Updating update program...');
	f = open('Version', 'w');
	f.write('.'.join((v[0], '-update-')));
	f.close();
	get(a, 'update.py');
	get(a, 'UPDATE.bat');
	print('Restarting...');
	exit(102);

elif v[1] != version[1]:
	# Update of other file(s)
	print('Updating files...');
	f = open('Version', 'w');
	f.write('.'.join(v));
	f.close();
	for x in files:
		get(a, x);

subprocess.call(['pypy2-v5.9.0-win32\\pypy.exe', '-m', 'pip', 'install', 'git+https://github.com/TheElementalOfCreation/creatorUtils']);

print('Done.');