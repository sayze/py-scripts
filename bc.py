from pyPdf import PdfFileReader
from collections import OrderedDict
from pyPdf import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.graphics.barcode.common import I2of5
import re
import os
import StringIO
import sys


#BELOW VARIABLES MAY BE UPDATED
jnum4dig =  "3426" # Job number
src = "BAT1_TESTFILE.pdf" #source pdf
hopper = "000000" # hopper information

# PROGRAM STARTS HERE
sheets = OrderedDict()
infile = PdfFileReader(file(src, "rb"))
	
if infile.getNumPages() > 1000:
	print "\nMaximum 1000 pages at a time(split document)\n"
	exit()

#FIRST -- read pdf contents	and store details in dictionary	
for idx in xrange(infile.getNumPages()):
	cont = infile.getPage(idx).extractText()
	
	cont = "\n".join(cont.replace(u"\xa0", " ").strip().split())
	open("log.txt","a").write(cont.encode('utf-8')+"\n*************************\n")
	cont = cont.split("\n")

	#get employee number position in data
	empidx = [i for i, item in enumerate(cont) if re.search(r"Number:\d+", item)]
	#get employee number
	empid = re.findall(r"\d+", cont[empidx[0]])
	
	if empid[0] in sheets:
		sheets[empid[0]] += 1		
	else:
		sheets[empid[0]] =1
	print "Reading : " + str(idx) + "\n"	

		
#SECOND -- Generate barcodes for pages and store in array
ctr  = 0 # Sequence number
barcodes = [] # store barcodes
for emp, sheet in sheets.iteritems():
	ctr += 1
	for sidx in xrange(0,sheet):
		kern = "*"
		kern  += jnum4dig
		kern += '%06d' % ctr
		kern += str(sidx+1) + str(sheet)
		kern += hopper
		kern += "*"
		barcodes.append(kern)
		

#THIRD -- Add the stored barcodes to the PDF and make new file
cleaned = "Barcoded" # barcode output dir
output = PdfFileWriter()
base = os.path.basename(src) # get file name of input
base = base[0:base.find('.')]+"_BC.pdf" # Append BC to output
spool = os.path.join(cleaned,base) # Final spool location
parts = 0
debug = open("log.txt","a")

if not os.path.exists(cleaned):
    os.makedirs(cleaned)	

for bidx in xrange(len(barcodes)):
	packet = StringIO.StringIO()
	print "writing : " + str(bidx) + "\n"	
	# create a new PDF with Reportlab
	can = canvas.Canvas(packet, pagesize=letter)
	can.rotate(90)
	barcode =  I2of5(barcodes[bidx], barWidth=0.3*mm,barHeight=8*mm, bearers=0,checksum=0)
	barcode.drawOn(can,568, -30)
	can.save()		
	packet.seek(0)
	
	#add barcode to new page	
	new_pdf = PdfFileReader(packet)
	page = infile.getPage(bidx)
	page.mergePage(new_pdf.getPage(0))
	output.addPage(page)
	
	
	

print "Creating file..."
# finally, write "output" to a real file
outputStream = file(spool, "wb")
output.write(outputStream)
outputStream.close()