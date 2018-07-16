import path;
import shutil;
import os;
import progressbar;
import json;

def tryToCopy(source, output, num = 0):
	if num > 5:
		raise Exception('Could not copy file as "{}"'.format(source));
	try:
		shutil.copyfile(source, output);
		try:
			os.stat(output);
		except:
			tryToCopy(source, output, num + 1);
	except Exception as e:
		print(e);


def start():
	while True:
		try:
			a = raw_input('Enter folder path you want to search (You may just drag the folder onto this window):\n');
			a = a.replace('"', '').replace("'", '');
			f = raw_input('Enter destination folder (Again, you can drag and drop):\n');
			f = f.replace('\\', '/').replace('"', '').replace("'", '');
			try:
				os.stat(f);
			except Exception as e:
				raise Exception('Could not access destination directory. Reason: {0}'.format(e));
			if f[-1] != '/':
				f += '/';
			probar = progressbar.ProgressBar();
			try:
				print('Starting Search. Depending on the amount of subdirectories, this may take a while...');
				b, c, d = path.getall(a, True, ['pdf'], progressBar = progressbar.ProgressBar());
				file1 = open('Session paths.out', 'w');
				file2 = open('Session filenames.out', 'w');
				file1.write(json.dumps(b));
				file1.close();
				file2.write(json.dumps(c));
				file2.close();
			except Exception as e:
				raise Exception('Could not search directory. Reason: {0}'.format(e));
			for x in probar(b):
				tryToCopy(x, f + x.split('/')[-1]);
			os.remove('Session filenames.out')
			os.remove('Session paths.out')
		except:
			raise;


if __name__ == '__main__':
	start()
