# pyQuestionnaire

[![forthebadge made-with-python](http://ForTheBadge.com/images/badges/made-with-python.svg)](https://www.python.org/) [![forthebadge made-with-python](http://ForTheBadge.com/images/badges/made-with-python.svg)](https://www.python.org/)
![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white) ![Linux](https://img.shields.io/badge/Linux-FCC624?style=for-the-badge&logo=linux&logoColor=black) ![poetry-badge](https://img.shields.io/badge/packaging-poetry-cyan.svg?style=for-the-badge)

## A better way to make questionnaires

### ./gen.py
generates the pages on ./out/ from doc.xlsx

### ./scan.py
constantly queries the printer for a scan and ignores "no paper" 
error, saves output to ./scans/scan_{count}.png (only works on linux)

### ./scanorganizer.py
reads the binary in the topleft corner from the photos in ./scans/ and puts them in ./sortedscans/scan_{count}_page_{binary}.png

### ./parse.py
reads the images from ./sortedscans/ and question types(1to5/binary) from doc.xlsx, adds all the values, puts them back into the excel document and saves it as ./out.xlsx
