from os import listdir, chdir
from os.path import isfile, join

def getPictureNames():
	print 'updating picture name file...'
	mypath ='U:\Merchandising\CompressedPhotos'
	pictureNames = open("U:\Retail Reporting\pictureNames.txt", "w")
	onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

	for item in onlyfiles:
		item = mypath + '\\' + item
	  	pictureNames.write("%s\n" % item)

	pictureNames.close()

	print "Got all the picture names!"