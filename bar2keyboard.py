import msvcrt as m
import os
import win32gui
import win32con
import win32api
from ctypes import windll, byref
from ctypes.wintypes import SMALL_RECT
import time
import keyboard
from datetime import datetime
import csv
import sys
import json

# Script was originally made to work with old non-web version of Sage accounting software which wasn't working with barcode reader
# This script/pyinstaller final executable supports only Windows since it rely on windows api methods to keep it on top of other apps and detecting keypresses using msvcrt. Can be altered by removing the two methods below

def fixToTop(w):
	hwnd = win32gui.GetForegroundWindow()
	win32gui.SetWindowPos(hwnd,win32con.HWND_TOPMOST,win32api.GetSystemMetrics(0)-int(win32api.GetSystemMetrics(0)/w),0,int(win32api.GetSystemMetrics(0)/w),win32api.GetSystemMetrics(1)-100,0)

def getCode():
	global active
	global l
	global pBar
	global pDesc
	global uniqueL
	inp=""
	while True:
		r = m.getch()
		if r==b'\r':
			break
		elif (r==b'\b'):
			inp=inp[:-1]
			print('\b \b',end='',flush=True)
			continue
		elif (r==b'\xe0'):
			next = m.getch()
			if next==b'H':
				# print("UP")
				if(active>0):
					active-=1
					reDraw(l,pBar,pDesc)
			elif next==b'P':
				# print("DOWN")
				if(active<len(uniqueL)-1):
					active+=1
					reDraw(l,pBar,pDesc)
			elif next==b'S':
				# print("DELETE")
				l.reverse()
				l.remove(uniqueL[active])
				l.reverse()
				uniqueL = list(dict.fromkeys(l))
				reDraw(l,pBar,pDesc)
			continue

		print(r.decode("utf-8"),end='',flush=True)
		inp += str(r.decode("utf-8"))
	return inp

#redraw whole list. Called after any changes to lists or to move pointer
def reDraw(l,pBar,pDesc):
	global uniqueL
	global active
	os.system('cls')
	pos = 0
	for i in uniqueL:
		try:
			print(("[ "if pos==active else "  ") + i +" x "+ str(l.count(i)) + " " + pDesc[i] + " " + pBar[i] + (" ]"if pos==active else "  "))
		except KeyError:
			print(("[ "if pos==active else "  ") + i +" x "+ str(l.count(i)) + " " + pDesc[i] + (" ]"if pos==active else "  "))
		pos+=1
			
def clickToPaste(l):
	global uniqueL
	print("")
	print("Click 1st field")

	state_left = win32api.GetKeyState(0x01)  # Left button down = 0 or 1. Button up = -127 or -128
	state_right = win32api.GetKeyState(0x02)  # Right button down = 0 or 1. Button up = -127 or -128

	while True:
	    a = win32api.GetKeyState(0x01)

	    if a != state_right:  # Button state changed
	        state_right = a
	        if a < 0:
	            print('Starting input')
	            time.sleep(1)
	            for i in uniqueL:
		            keyboard.write(i,delay=0.05)
		            time.sleep(0.2)
		            keyboard.send('enter')
		            time.sleep(0.2)
		            keyboard.send('enter')
		            time.sleep(0.2)
		            keyboard.write(str(l.count(i)))
		            time.sleep(0.2)
		            keyboard.send('enter')
		            time.sleep(0.2)
		            keyboard.send('enter')
		            time.sleep(0.2)
	            x = input("Press R to restart or P to paste again\n")
	            if(x.upper()=='R'):
	            	return
	            elif(x.upper()=='P'):
	            	time.sleep(0.001)
	            	continue
	            else:
	            	sys.exit()	            	


	    time.sleep(0.001)

def setActive(ins):
	global active
	global uniqueL
	pos=0
	for i in uniqueL:
		if i==ins:
			active=pos
			return
		pos+=1


while True:
	# Config json/text file to import settings (end user cannot modify pyinstaller exe)
	try:
		with open("config.txt") as f:
			config = f.read()
		jsonData = json.loads("{"+config+"}")

		fileName = jsonData["fileName"]
		nthArea = jsonData["nthArea"]
	except:
		print("Invalid/missing config.txt")
		os.system('pause')
		sys.exit()

	
	os.system('cls')
	l = []
	uniqueL = []
	active = 0

	# Map productcodes,barcodes and descriptions into dictionaries

	reader = csv.reader(open(fileName))
	pBar = dict((rows[0],rows[2]) for rows in reader)

	reader = csv.reader(open(fileName))
	pDesc = dict((rows[0],rows[1]) for rows in reader)

	reader = csv.reader(open(fileName))
	bPcode = dict((rows[2],rows[0]) for rows in reader)

	# fix screen to front of every application (WORKS ONLY IN WINDOWS)
	# widthDivider accept values more than 1. area occupied by application becomes 1/widthDivider. Example, Passing 3 will make application take 1/3 area of screen, 2 will make it half
	fixToTop(widthDivider)
	while True:
		print("---START SCANNING OR ENTER PRODUCT CODE MANUALLY---")
		inp = getCode()
		if(inp==""):
			break
		inp = inp.upper()
		if(inp in bPcode):
			l.append(bPcode[inp])
			uniqueL = list(dict.fromkeys(l))
			setActive(inp)
			reDraw(l,pBar,pDesc)
		elif(inp in pBar):
			l.append(inp)
			uniqueL = list(dict.fromkeys(l))
			setActive(inp)
			reDraw(l,pBar,pDesc)
		else:
			print("\n"+inp+" Not found")
	clickToPaste(l)