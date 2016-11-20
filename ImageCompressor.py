from os import listdir, chdir
from os.path import isfile, join
from PIL import Image
from PhotoNameComb import combPhotoNames


directory = 'U:\Merchandising\CompressedPhotos'
chdir(directory)

onlyPics = [f for f in listdir(directory) if isfile(join(directory, f))]
print 'compressing photos...'
for pic in onlyPics:
	try:
		image = Image.open(pic)
		
		if image.size[1] > 400:
			ratio = float(image.size[0])/image.size[1]
			height = 400
			width = int(400 * ratio)
		else:
			height = image.size[1]
			width = image.size[0]

		image = image.resize((width, height), Image.ANTIALIAS)
		newLocation = 'U:\Merchandising\CompressedPhotos\\' + pic
		image.save(newLocation, optimize= True, quality=95)
	except IOError:
		"Whoops! Couldn't open this one, boss! --> " + pic
	print pic
print 'done compressing photos'
combPhotoNames()



