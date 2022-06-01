# Filename: rfd_deals.py
# Name: Jason Luu
# Description: Checks RedFlagDeals Deals on Computers and Electronics
#              every 5 minutes for new deals
#              Toast Windows 10 to show new deals every 2 minutes
#

#import libraries
from urllib.request import urlopen
from bs4 import BeautifulSoup
from win10toast import ToastNotifier
import time

deals_page = 'https://forums.redflagdeals.com/hot-deals-f9/?c=9' #RedFlagDeals Deals page for "Computers and Electronics"
dealList = [] #holds all the deals
dealExist = 0 #checks for existing deals

#Sleep Time
convertMinstoSecs = 60
loopSleep = 2* convertMinstoSecs
showToastTime = 4

toaster = ToastNotifier() #Setup Windows 10 Toast
infLoop = 0

print('Running script')

while infLoop < 10:
	page = urlopen(deals_page) #Open page
	soup = BeautifulSoup(page, 'html.parser') #Parse the page
	#PART 1: Retrieve deals from RedFlagDeals
	for i in range (0,20): #Get 20 topics top to bottom
		newDeal = soup.findAll('h3', attrs={'class': 'topictitle topictitle_has_retailer'})[i].text

		#Iterate through the whole list and compares newDeal to each item
		for deal in dealList:
			if newDeal == deal:
				dealExist = 1

		#If duplicate isn't found, add to list
		if dealExist == 0:
			dealList.append(newDeal)
		else:
			dealExist = 0

	#PART 2: Show Toast Notifications in Windows 10
	numberOfDeals = len(dealList) #Counts the number of deals
	#loop through each deal in list and toast it every 2 minutes
	i = 0
	while i < numberOfDeals:
		toaster.show_toast('Deal', dealList[i],icon_path='.\\cart.ico',duration=5)
		time.sleep(showToastTime)
		i += 1

	#Start over the script by deleting entire dealList and sleep for a specified time
	dealList.clear()
	time.sleep(loopSleep) #Every X minutes check for new deals
