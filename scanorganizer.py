import PIL
import PIL.Image
import os, shutil


# res = {i: [] for i in reversed(os.listdir("scans"))}
scans = {}

for idx, i in enumerate(os.listdir("scans")):
	image = PIL.Image.open("scans/"+i)

	b1 = sum([i[0] for i in image.crop(( 35, 35,  45, 45)).getcolors() if not (i[1][0] > 120 and i[1][1] > 120 and i[1][2] > 120)]) > 50
	b2 = sum([i[0] for i in image.crop(( 60, 35,  70, 45)).getcolors() if not (i[1][0] > 120 and i[1][1] > 120 and i[1][2] > 120)]) > 50
	b3 = sum([i[0] for i in image.crop(( 83, 35,  93, 45)).getcolors() if not (i[1][0] > 120 and i[1][1] > 120 and i[1][2] > 120)]) > 50
	b4 = sum([i[0] for i in image.crop((105, 35, 115, 45)).getcolors() if not (i[1][0] > 120 and i[1][1] > 120 and i[1][2] > 120)]) > 50
	pageindex = (b4*1)+(b3*2)+(b2*4)+(b1*8)
	print(i+": "+str(pageindex))
	try:
		scans[pageindex] += 1
	except:
		scans[pageindex]  = 1
	shutil.copy2("scans/"+i, f"sortedscans/scan_{scans[pageindex]}_page_{pageindex}.png")

