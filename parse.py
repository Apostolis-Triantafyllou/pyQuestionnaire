import PIL
import PIL.Image
import sys, os, openpyxl, re

doc = openpyxl.open("doc.xlsx")
sheet = doc['data']

pagenscan = re.compile(r"scan_([0-9]+)_page_([0-9]+)\.png").match
scans = {pagenscan(i)[2]: [] for i in os.listdir("sortedscans")}

# print(scans)
# exit()

for i in os.listdir("sortedscans"):

	scan, page = (lambda a: (int(a[1]), int(a[2])))(pagenscan(i))
	print(f"{i}: {scan}: {page}")


	image = PIL.Image.open("sortedscans/"+i)


	match sheet[f'D{((page-1)*5)+2}'].value:
		case '1to5':
			aa = {i[1]: i[0] for i in image.crop(( 80, 185,  95, 200)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ab = {i[1]: i[0] for i in image.crop((170, 185, 185, 200)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ac = {i[1]: i[0] for i in image.crop((265, 185, 280, 200)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ad = {i[1]: i[0] for i in image.crop((360, 185, 375, 200)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ae = {i[1]: i[0] for i in image.crop((455, 185, 470, 200)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			a = {'a': sum(aa.values()), 'b': sum(ab.values()), 'c': sum(ac.values()), 'd': sum(ad.values()), 'e': sum(ae.values())}
			avalprev = 0
			ares = 'blank'
			for key,val in a.items():
				if val > avalprev:
					ares = key
			match ares:
				case 'blank':
					sheet[f'E{((page-1)*5)+2}'].value += 1
				case 'e':
					sheet[f'F{((page-1)*5)+2}'].value += 1
				case 'd':
					sheet[f'G{((page-1)*5)+2}'].value += 1
				case 'c':
					sheet[f'H{((page-1)*5)+2}'].value += 1
				case 'b':
					sheet[f'I{((page-1)*5)+2}'].value += 1
				case 'a':
					sheet[f'J{((page-1)*5)+2}'].value += 1

		case 'boolean':
			at = {i[1]: i[0] for i in image.crop(( 80, 190,   95, 205)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			af = {i[1]: i[0] for i in image.crop((230, 190,  245, 205)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			a = {True: sum(at.values()), False: sum(af.values())}
			avalprev = 0
			ares = 'blank'
			for key,val in a.items():
				if val > avalprev:
					ares = key
			match ares:
				case 'blank':
					print("a: blank")
					sheet[f'E{((page-1)*5)+2}'].value += 1
				case True:
					print("a: True")
					sheet[f'F{((page-1)*5)+2}'].value += 1
				case False:
					print("a: False")
					sheet[f'G{((page-1)*5)+2}'].value += 1

	match sheet[f'D{((page-1)*5)+3}'].value:
		case '1to5':
			aa = {i[1]: i[0] for i in image.crop(( 80, 340,  95, 355)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ab = {i[1]: i[0] for i in image.crop((170, 340, 185, 355)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ac = {i[1]: i[0] for i in image.crop((265, 340, 280, 355)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ad = {i[1]: i[0] for i in image.crop((360, 340, 375, 355)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ae = {i[1]: i[0] for i in image.crop((455, 340, 470, 355)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			a = {'a': sum(aa.values()), 'b': sum(ab.values()), 'c': sum(ac.values()), 'd': sum(ad.values()), 'e': sum(ae.values())}
			avalprev = 0
			ares = 'blank'
			for key,val in a.items():
				if val > avalprev:
					ares = key
			match ares:
				case 'blank':
					sheet[f'E{((page-1)*5)+3}'].value += 1
				case 'e':
					sheet[f'F{((page-1)*5)+3}'].value += 1
				case 'd':
					sheet[f'G{((page-1)*5)+3}'].value += 1
				case 'c':
					sheet[f'H{((page-1)*5)+3}'].value += 1
				case 'b':
					sheet[f'I{((page-1)*5)+3}'].value += 1
				case 'a':
					sheet[f'J{((page-1)*5)+3}'].value += 1

		case 'boolean':
			bt = {i[1]: i[0] for i in image.crop(( 80, 340,   95, 355)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			bf = {i[1]: i[0] for i in image.crop((230, 340,  245, 355)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			b = {True: sum(bt.values()), False: sum(bf.values())}
			bvalprev = 0
			bres = 'blank'
			for key,val in b.items():
				if val > bvalprev:
					bres = key
			match bres:
				case 'blank':
					sheet[f'E{((page-1)*5)+3}'].value += 1
				case True:
					sheet[f'F{((page-1)*5)+3}'].value += 1
				case False:
					sheet[f'G{((page-1)*5)+3}'].value += 1

	match sheet[f'D{((page-1)*5)+4}'].value:
		case '1to5':
			aa = {i[1]: i[0] for i in image.crop(( 80, 485,  95, 500)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ab = {i[1]: i[0] for i in image.crop((170, 485, 185, 500)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ac = {i[1]: i[0] for i in image.crop((265, 485, 280, 500)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ad = {i[1]: i[0] for i in image.crop((360, 485, 375, 500)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ae = {i[1]: i[0] for i in image.crop((455, 485, 470, 500)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			a = {'a': sum(aa.values()), 'b': sum(ab.values()), 'c': sum(ac.values()), 'd': sum(ad.values()), 'e': sum(ae.values())}
			avalprev = 0
			ares = 'blank'
			for key,val in a.items():
				if val > avalprev:
					ares = key
			match ares:
				case 'blank':
					sheet[f'E{((page-1)*5)+4}'].value += 1
				case 'e':
					sheet[f'F{((page-1)*5)+4}'].value += 1
				case 'd':
					sheet[f'G{((page-1)*5)+4}'].value += 1
				case 'c':
					sheet[f'H{((page-1)*5)+4}'].value += 1
				case 'b':
					sheet[f'I{((page-1)*5)+4}'].value += 1
				case 'a':
					sheet[f'J{((page-1)*5)+4}'].value += 1

		case 'boolean':
			ct = {i[1]: i[0] for i in image.crop(( 80, 465,   95, 480)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			cf = {i[1]: i[0] for i in image.crop((230, 465,  245, 480)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			c = {True: sum(ct.values()), False: sum(cf.values())}
			cvalprev = 0
			cres = 'blank'
			for key,val in c.items():
				if val > cvalprev:
					cres = key
			match cres:
				case 'blank':
					sheet[f'E{((page-1)*5)+4}'].value += 1
				case True:
					sheet[f'F{((page-1)*5)+4}'].value += 1
				case False:
					sheet[f'G{((page-1)*5)+4}'].value += 1

	match sheet[f'D{((page-1)*5)+5}'].value:
		case '1to5':
			aa = {i[1]: i[0] for i in image.crop(( 80, 625,  95, 640)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ab = {i[1]: i[0] for i in image.crop((170, 625, 185, 640)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ac = {i[1]: i[0] for i in image.crop((265, 625, 280, 640)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ad = {i[1]: i[0] for i in image.crop((360, 625, 375, 640)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ae = {i[1]: i[0] for i in image.crop((455, 625, 470, 640)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			a = {'a': sum(aa.values()), 'b': sum(ab.values()), 'c': sum(ac.values()), 'd': sum(ad.values()), 'e': sum(ae.values())}
			avalprev = 0
			ares = 'blank'
			for key,val in a.items():
				if val > avalprev:
					ares = key
			match ares:
				case 'blank':
					sheet[f'E{((page-1)*5)+5}'].value += 1
				case 'e':
					sheet[f'F{((page-1)*5)+5}'].value += 1
				case 'd':
					sheet[f'G{((page-1)*5)+5}'].value += 1
				case 'c':
					sheet[f'H{((page-1)*5)+5}'].value += 1
				case 'b':
					sheet[f'I{((page-1)*5)+5}'].value += 1
				case 'a':
					sheet[f'J{((page-1)*5)+5}'].value += 1

		case 'boolean':
			dt = {i[1]: i[0] for i in image.crop(( 80, 620,   95, 635)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			df = {i[1]: i[0] for i in image.crop((230, 620,  245, 635)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			d = {True: sum(dt.values()), False: sum(df.values())}
			dvalprev = 0
			dres = 'blank'
			for key,val in d.items():
				if val > dvalprev:
					dres = key
			match dres:
				case 'blank':
					sheet[f'E{((page-1)*5)+5}'].value += 1
				case True:
					sheet[f'F{((page-1)*5)+5}'].value += 1
				case False:
					sheet[f'G{((page-1)*5)+5}'].value += 1

	match sheet[f'D{((page-1)*5)+6}'].value:
		case '1to5':
			aa = {i[1]: i[0] for i in image.crop(( 80, 770,  95, 785)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ab = {i[1]: i[0] for i in image.crop((170, 770, 185, 785)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ac = {i[1]: i[0] for i in image.crop((265, 770, 280, 785)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ad = {i[1]: i[0] for i in image.crop((360, 770, 375, 785)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ae = {i[1]: i[0] for i in image.crop((455, 770, 470, 785)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			a = {'a': sum(aa.values()), 'b': sum(ab.values()), 'c': sum(ac.values()), 'd': sum(ad.values()), 'e': sum(ae.values())}
			avalprev = 0
			ares = 'blank'
			for key,val in a.items():
				if val > avalprev:
					ares = key
			match ares:
				case 'blank':
					sheet[f'E{((page-1)*5)+6}'].value += 1
				case 'e':
					sheet[f'F{((page-1)*5)+6}'].value += 1
				case 'd':
					sheet[f'G{((page-1)*5)+6}'].value += 1
				case 'c':
					sheet[f'H{((page-1)*5)+6}'].value += 1
				case 'b':
					sheet[f'I{((page-1)*5)+6}'].value += 1
				case 'a':
					sheet[f'J{((page-1)*5)+6}'].value += 1

		case 'boolean':
			et = {i[1]: i[0] for i in image.crop(( 80, 760,   95, 775)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			ef = {i[1]: i[0] for i in image.crop((230, 760,  245, 775)).getcolors() if not (i[1][0] > 160 and i[1][1] > 160 and i[1][2] > 160)}
			e = {True: sum(et.values()), False: sum(ef.values())}
			evalprev = 0
			eres = 'blank'
			for key,val in e.items():
				if val > evalprev:
					eres = key
			match eres:
				case 'blank':
					sheet[f'E{((page-1)*5)+6}'].value += 1
				case True:
					sheet[f'F{((page-1)*5)+6}'].value += 1
				case False:
					sheet[f'G{((page-1)*5)+6}'].value += 1

	doc.save("out.xlsx")
