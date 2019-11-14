import http.client as httplib
import os
import shutil
import subprocess
import time

from creatorUtils.path import get_short_path_name

# Determine the directory for the other resources.
# This MUST be the first global variable in this list.
# "__file__" is a variable in every python script that
# contains the path of that file.
resDir = get_short_path_name('/'.join(os.path.abspath(__file__).split('\\')[0:-1]))
# Define path to error folder
error = get_short_path_name(resDir + '/ERROR')

date = int(time.time())

back_init = resDir + '/BACKUP/' + str(date) + '/'

try:
	os.makedirs(back_init)
except:
	pass

back = get_short_path_name(back_init)


print('Initializing update program...')

def get(conn, x):
	conn.request('GET', url + x)
	data = conn.getresponse().read()
	shutil.move(x, back + x + '.old')
	file1 = open(x, 'wb')
	file1.write(data)
	file1.close()

files = [
	'msg.py',
	'toPdf.bat',
	'search.py',
	'SEARCH.bat',
	'Version',
]

url = '/TheElementalOfCreation/search.py/master/'
data = {}
version = (None, None)

try:
	with open('Version', 'r') as a:
		version = a.read().split('.')
except:
	pass;

print('Checking for new version...')

a = httplib.HTTPSConnection('raw.githubusercontent.com')
a.request('GET', url + 'Version')
v = a.getresponse().read().split('.')
if v[0] != version[0]:
	# update.py and/or UPDATE.bat update
	print('Updating update program...')
	f = open('Version', 'w')
	f.write('.'.join((v[0], '-update-')))
	f.close()
	get(a, 'update.py')
	get(a, 'UPDATE.bat')
	print('Restarting...')
	exit(102)

elif v[1] != version[1]:
	# Update of other file(s)
	print('Updating files...')
	f = open('Version', 'w')
	f.write('.'.join(v))
	f.close()
	for x in files:
		get(a, x)

subprocess.call(['Python36_64\\python.exe', '-m', 'pip', 'install', '--disable-pip-version-check', 'git+https://github.com/TheElementalOfCreation/creatorUtils'])
subprocess.call(['Python36_64\\python.exe', '-m', 'pip', 'install', '--disable-pip-version-check', '--upgrade', 'extract-msg'])

print('Done.')
