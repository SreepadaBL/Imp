import pyautogui
from time import sleep
import os
import time
import pyperclip
import threading
from threading import RLock
from threading import Thread, current_thread
l=[]
x=360
y=295
a=400
b=580
q=260
r=295
filename_type=str(input("Enter Filename:"))
n=int(input("Enter No of Files:"))
delay1=eval(input("Enter the delay:"))
dirpath_type=str(input("Enter dirpath:"))
lock = threading.RLock() 
print(threading.activeCount())
print(current_thread().name)
def fun1(lock,filename_type,n,delay1,dirpath_type):
	count=0
	for m in range(1,n+1):
		lock.acquire()
		print(current_thread().name)
		count=count+1
		if (m%100==0):
			lock.release()
			print("Current thread:",current_thread().name)
			sleep(20)
			lock.acquire()
		if(m<=30): 
			pyautogui.click(95,110+(m*30))
			sleep(delay1)
			pyautogui.typewrite(filename_type+'_'+format(m,'05d'))
			sleep(delay1)
			pyautogui.hotkey('enter')
			sleep(delay1)
			pyautogui.click(955,85)
			sleep(delay1)
			pyautogui.click(95,110+(m*30))
			sleep(delay1)
			pyautogui.typewrite(dirpath_type+'_'+format(m,'05d'))
			sleep(delay1)
			pyautogui.hotkey('enter')
			sleep(delay1)
			pyautogui.click(820,80)
			sleep(delay1)
		else:
			pyautogui.hotkey('enter')
			pyautogui.click(95,110+(31*28))
			sleep(delay1)
			pyautogui.typewrite(filename_type+'_'+format(m,'05d'))
			sleep(delay1)
			pyautogui.hotkey('enter')
			sleep(delay1)
			pyautogui.click(955,85)
			sleep(delay1)
			pyautogui.click(95,110+(31*28))
			sleep(delay1)
			pyautogui.typewrite(dirpath_type+'_'+format(m,'05d'))
			sleep(delay1)
			pyautogui.hotkey('enter')
			sleep(delay1)
			pyautogui.click(820,80)
			sleep(delay1)
	print(count)
	lock.release() 
t1 = threading.Thread(target=fun1, args=(lock,filename_type,n,delay1,dirpath_type)) 
#fun1(filename_type=filename_type,n=n,delay1=delay1,dirpath_type=dirpath_type)
t1.start()
t1.join()
print(current_thread().name)
print('thread{}'.format(t1))

	#lock.release()
	#pyautogui.doubleClick(95,110)
	#pyautogui.hotkey('ctrl','c')
	#pyautogui.click(95,110)
	##pyautogui.click(95,110+(m*20))
	#pyautogui.hotkey('enter')
	##v=pyperclip.paste()
	#pyautogui.hotkey('ctrl','v')
	#sleep(delay1)
		#if(m<=22):
		#	z=pyautogui.moveTo(x,y+(22*m))
		#	sleep(delay1)
		#	pyautogui.doubleClick(z)
		#	sleep(delay1)
		#	pyautogui.typewrite(filename_type+'_'+format(m,'05d'))
		#	pyautogui.hotkey('enter')
		#	pyautogui.click(q,r+(m*22))
		#	sleep(delay1)
		#	pyautogui.click(1070, 190)
		#	sleep(delay1)
		#	pyautogui.click(400, 255)
		#	sleep(delay1)
		#	pyautogui.typewrite(dirpath_type+'\\'+filename_type+'_'+format(m,'05d')+".mpa")
		#	sleep(delay1)
		#	pyautogui.hotkey('enter')
		#	sleep(delay1)
		#	pyautogui.click(140,420)
		#	sleep(delay1)			
		#else:
		#	z=pyautogui.moveTo(x,y+(22*22))
		#	sleep(delay1)
		#	pyautogui.doubleClick(z)
		#	sleep(delay1)
		#	pyautogui.typewrite(filename_type+'_'+format(m,'05d'))
		#	pyautogui.hotkey('enter')
		#	sleep(delay1)
		#	p=pyautogui.moveTo(q,r+(22*22))
		#	sleep(delay1)
		#	pyautogui.click(p)
		#	sleep(delay1)
		#	pyautogui.click(1070, 190)
		#	sleep(delay1)
		#	pyautogui.click(400, 255)
		#	sleep(delay1)
		#	pyautogui.typewrite(dirpath_type+'\\'+filename_type+'_'+format(m,'05d')+".mpa")
		#	sleep(delay1)
		#	pyautogui.hotkey('enter')
		#	sleep(delay1)
		#	pyautogui.click(140,420)
		#	sleep(delay1)