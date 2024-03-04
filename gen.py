import PIL
import PIL.Image
import PIL.ImageDraw
import PIL.ImageFont
import math
import openpyxl

lefont1  = PIL.ImageFont.truetype('Cantarell-Regular.otf', 512)
lefont2  = PIL.ImageFont.truetype('Cantarell-Regular.otf', 256)
lefont25 = PIL.ImageFont.truetype('Cantarell-Regular.otf', 192)
lefont3  = PIL.ImageFont.truetype('Cantarell-Regular.otf', 128)


questions = [[str(i[0]), str(i[1]), str(i[3])] for i in list(openpyxl.open("doc.xlsx", data_only=True)['data'].values)[1:]]

print(questions)

for i in range(1, math.ceil(len(questions)/5)+1):
	page = PIL.Image.new(mode="RGB", size=(4961, 7016), color=(255,255,255))
	draw = PIL.ImageDraw.Draw(page)
	_, _, w, _ = draw.textbbox((0, 0), "Ερωτηματολόγιο", font=lefont1)
	pagecode = [True if k=="1" else False for k in f"{i:04b}"]
	draw.rectangle([(100, 300), (300, 500)], fill="#000000" if pagecode[0] else '#FFFFFF', outline="#000000", width=10)
	draw.rectangle([(300, 300), (500, 500)], fill="#000000" if pagecode[1] else '#FFFFFF', outline="#000000", width=10)
	draw.rectangle([(500, 300), (700, 500)], fill="#000000" if pagecode[2] else '#FFFFFF', outline="#000000", width=10)
	draw.rectangle([(700, 300), (900, 500)], fill="#000000" if pagecode[3] else '#FFFFFF', outline="#000000", width=10)
	draw.text(((4961-w)/2+400, 200), "Ερωτηματολόγιο", font=lefont1, fill=(0,0,0))
	for j in range(0, 5):
		if len(questions) <= 5*(i-1)+j:
			break
		print(i, end=" ")
		print(j, end=" ")
		print(5*(i-1)+j)
		draw.text((400, 1200*(j+1)-400), questions[5*(i-1)+j][0], font=lefont25, fill=(0,0,0))
		draw.text((400, 1200*(j+1)-225), questions[5*(i-1)+j][1], font=lefont25, fill=(0,0,0))
		match questions[5*(i-1)+j][2]:
			case '1to5':
				draw.text(( 400, (1200*(j+1))+75), "Απο 1 έως 5:", fill="#000000", font=lefont3)
				draw.rectangle([( 400, 1200*(j+1)+250), ( 800, (1200*(j+1))+650)], fill="#FFFFFF", outline="#000000", width=20)
				draw.rectangle([(1200, 1200*(j+1)+250), (1600, (1200*(j+1))+650)], fill="#FFFFFF", outline="#000000", width=20)
				draw.rectangle([(2000, 1200*(j+1)+250), (2400, (1200*(j+1))+650)], fill="#FFFFFF", outline="#000000", width=20)
				draw.rectangle([(2800, 1200*(j+1)+250), (3200, (1200*(j+1))+650)], fill="#FFFFFF", outline="#000000", width=20)
				draw.rectangle([(3600, 1200*(j+1)+250), (4000, (1200*(j+1))+650)], fill="#FFFFFF", outline="#000000", width=20)
				draw.text(( 850, (1200*(j+1))+300), "1", fill="#000000", font=lefont2)
				draw.text((1650, (1200*(j+1))+300), "2", fill="#000000", font=lefont2)
				draw.text((2450, (1200*(j+1))+300), "3", fill="#000000", font=lefont2)
				draw.text((3250, (1200*(j+1))+300), "4", fill="#000000", font=lefont2)
				draw.text((4050, (1200*(j+1))+300), "5", fill="#000000", font=lefont2)
			case 'boolean':
				draw.rectangle([(400, 1200*(j+1)+225), (800, (1200*(j+1))+625)], fill="#FFFFFF", outline="#000000", width=20)
				draw.text(( 850, 1200*(j+1)+275), "Ναι", fill="#000000", font=lefont2)
				draw.rectangle([(1600, 1200*(j+1)+225), (2000, (1200*(j+1))+625)], fill="#FFFFFF", outline="#000000", width=20)
				draw.text((2050, 1200*(j+1)+275), "Οχι", fill="#000000", font=lefont2)

	page.save("out/page_"+str(i)+".png")
