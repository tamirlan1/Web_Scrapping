# Note - this code must run in Python 2.x and you must download
# http://www.pythonlearn.com/code/BeautifulSoup.py
# Into the same folder as this program

import urllib
from BeautifulSoup import *
import re

url = 'https://www.sec.gov/cgi-bin/current?q1=0&q2=6&q3=4' #raw_input('Enter - ')
html = urllib.urlopen(url).read()

soup = BeautifulSoup(html)
# print soup
# Retrieve all of the anchor tags
total = 0
tags = soup('a')
# print str(tags[1])
print len(tags)
for tag in tags:
	if not str(tag).startswith('<a href="/Archives'):
		tags.remove(tag)
# 		# 
#     total += int(tag.contents[0])
# print total
print len(tags)
