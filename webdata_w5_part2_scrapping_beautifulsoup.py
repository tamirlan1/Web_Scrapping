# Note - this code must run in Python 2.x and you must download
# http://www.pythonlearn.com/code/BeautifulSoup.py
# Into the same folder as this program

import urllib
from BeautifulSoup import *

url = raw_input('Enter url: ')
count = int(raw_input('Enter count: '))
position = int(raw_input('Enter position: '))
print url
for i in range(count):
	counter = 0
	html = urllib.urlopen(url).read()
	soup = BeautifulSoup(html)

	# Retrieve all of the anchor tags
	tags = soup('a')
	for tag in tags:
		counter += 1
		url = tag.get('href', None)
		if counter == position:
			print url
			break
