############################
### Works only on linux  ###
### Install sane package ###
############################


import sane

sane.init()
printer = sane.open("hpaio:/net/DeskJet_3700_series?ip=192.168.1.51")
printer.mode = 'color'
try:
	i = 0
	while True:
		i += 1
		while True:
			try:
				printer.scan().save("scans/scan_"+str(i)+".png")
				print("scanned page "+str(i))
				break
			except sane._sane.error as e:
				if not e.args[0] == "Document feeder out of documents":
					print(e)
					break
except Exception as e:
	print(e)
	printer.close()
