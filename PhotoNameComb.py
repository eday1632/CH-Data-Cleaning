import os, sys
from GetPictureNames import getPictureNames

def combPhotoNames():
	mypath = "U:\Merchandising\CompressedPhotos"

	os.chdir(mypath)
	print 'combing photo names...'
	for item in os.listdir(mypath):
		tempName = item

		if '.png' in tempName:
			tempName = tempName.replace('.png', '.jpg')
		
		try:
			os.rename(item, tempName)
		except Exception, e:
			print "Error! --> " + item + " and " + tempName

	print "All done combing!"
	getPictureNames()