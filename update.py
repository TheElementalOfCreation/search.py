import httplib;
import os;
import subprocess;

def get(conn, x):
	conn.request('GET', url.format(x));
	data = conn.getresponse().read();
	file1 = open(x, 'wb');
	file1.write(data);
	file1.close();

__git__ = '/'.join(__file__.replace('\\', '/').split('/')[:-1] + ['git', 'bin']);

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
	version = a.split('.');
	a.close();
except:
	pass;

a = httplib.HTTPSConnection('raw.githubusercontent.com');
a.request('GET', url.format('Version'));
v = a.getresponse().read().split('.');
if v[0] != version[0]:
	# update.py and/or UPDATE.bat update
	f = open('Version');
	f.write('.'.join((v[0], '-update-')));
	get(a, 'update.py');
	get(a, 'UPDATE.bat');
	exit(102);

elif v[1] != version[1]:
	# Update of other file(s)
	for x in files:
		get(a, x);

os.envrion['PATH'] += __git__ + ';';
os.envrion.update();
subprocess.call(['pypy2-v5.9.0-win32\\pypy.exe', '-m', 'pip', 'install', 'git+https://github.com/TheElementalOfCreation/creatorUtils']);