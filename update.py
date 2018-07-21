import httplib;
import os;

files = [
	'msg.py',
	'toPdf.bat',
	'search.py',
	'SEARCH.bat',
	'update.py',
	'UPDATE.bat',
	'Version',
	'pypy2-v5.9.0-win32/lib-python/2.7/creatorUtils/path.py',
	'pypy2-v5.9.0-win32/lib-python/2.7/creatorUtils/__init__.py',
	'pypy2-v5.9.0-win32/lib-python/2.7/creatorUtils/compat/__init__.py',
	'pypy2-v5.9.0-win32/lib-python/2.7/creatorUtils/compat/progress_bar.py',
	'pypy2-v5.9.0-win32/lib-python/2.7/creatorUtils/compat/types.py',
];

url = '/TheElementalOfCreation/search.py/master/{}';
data = {};
version = None;

try:
	
except:
	

a = httplib.HTTPSConnection('raw.githubusercontent.com');
for x in files:
	a.request('GET', url.format(x));
	data[x] = a.getresponse().read();
	