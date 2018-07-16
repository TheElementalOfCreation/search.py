import httplib;
import os;
import _md5;

files = [
	'msg.py',
	'toPdf.bat',
	'search.py',
	'SEARCH.bat',
	'update.py',
	'UPDATE.bat'
];

url = '/TheElementalOfCreation/search.py/master/{}';
md5sCurrent = {};
md5sNew = {};
data = {};

try:
	file1 = open('md5Hashes/md5.txt');
	for line in file1:
		md5sCurrent[line.split('\t')[0]] = line.split('\t')[1];
	file1.close();
except:
	for x in files:
		file1 = open(x, 'rb');
		md5sCurrent[x] = _md5.new(file1.read()).hexdigest();
		file1.close();

a = httplib.HTTPSConnection('raw.githubusercontent.com');
for x in files:
	a.request('GET', url.format(x));
	data[x] = a.getresponse().read();
	md5sNew[x] = _md5.new(data[x]).hexdigest();
	if md5sNew[x] != md5sCurrent[x]:
		file1 = open(x, 'wb');
		file1.write(data[x]);
		file1.close();


file2 = open('md5Hashes/md5.txt', 'w');
for x in files:
	file2.write(x + '\t' + md5sNew[x] + '\n');
file2.close();