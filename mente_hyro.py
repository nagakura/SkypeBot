#encoding: utf-8

import Skype4Py
import time
import re

import xlrd
import random

wb = xlrd.open_workbook("hyro2.xls")
sheets = wb.sheets()

s = sheets[2]

def selectPost():
	rand = random.randint(0, s.nrows)
	c = s.cell(rand, 0)
	return c.value

def handler(msg, event):
	if event == u"RECEIVED":
		if msg.Body == u"ヒイロ":
			msg.Chat.SendMessage(u"ヒイロ・ユイ:「メンテ中だぁぁ！！！」")

def main():
	skype = Skype4Py.Skype(Transport='x11')
	skype.OnMessageStatus = handler
	skype.Attach()

	while True:
		time.sleep(1)

if __name__ == "__main__":
	main()
