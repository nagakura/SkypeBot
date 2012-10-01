#!/usr/bin/python
#-*-coding: utf-8-*-
from __future__ import with_statement
from daemon import DaemonContext
from lockfile.pidlockfile import PIDLockFile

import Skype4Py
import time
import re

import xlrd
import random
import daemon

#read excel file
wb = xlrd.open_workbook("hyro2.xls")
sheets = wb.sheets()

s = sheets[2]

#leguler
p = re.compile(u'ヒイロ')

#daemon settings
dc = DaemonContext()

def selectPost():
	rand = random.randint(0, s.nrows)
	c = s.cell(rand, 0)
	return c.value

def handler(msg, event):
	if event == u"RECEIVED":
		#out = msg.Body.encode('utf-8')
		if msg.Body == u"ヒイロ":
			msg.Chat.SendMessage(u"ヒイロ・ユイ:「%s」" % selectPost())

def main():
	skype = Skype4Py.Skype(Transport='x11')
	skype.OnMessageStatus = handler
	skype.Attach()

	while True:
		time.sleep(10)

with dc:
	if __name__ == "__main__":
		main()
