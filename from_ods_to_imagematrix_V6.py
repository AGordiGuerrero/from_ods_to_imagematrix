
################################################
# Extract contents from a Open Office sheet's columns and convert it to text in image arrays.
# Can rescale the input image with command line arguments. 

# Created by A.Gordillo-Guerrero (anto@unex.es), September , 2023.
# University of Extremadura. Smart Open Lab.
# Released into the public domain.

################################################

#!/usr/bin/env python3


# To use libre office and excel sheets
import pyexcel
from pyexcel._compact import OrderedDict

# To modify and create images with texts
from PIL import Image, ImageFont, ImageDraw

# To read arguments from command line
import argparse

# To make good logs with info instead of only print()
import logging

text = 'This program read data from a Libre Office, open document, sheet and convert it to centered text, with a given font, in a matrix from a base image file.\n\nRun with -h option to see command line arguments.\n\n'


# Show info to user
print(text)
# Configuring the log file
logging.basicConfig(filename='output.log', encoding='utf-8', level=logging.DEBUG)
logging.info(text)


################################################
# Reading arguments from the command line
################################################

# Initiate the parser with a description
parser = argparse.ArgumentParser(description=text)

parser.add_argument("--outWidth", "-ow", help="Width in mm of the output image", type=int, default=210, nargs='?') # default A4 size
parser.add_argument("--outHeight", "-oh", help="Height in mm of the output image", type=int, default=297, nargs='?') # default A4 size
parser.add_argument("--inWidth", "-iw", help="Width in mm for the scaled input image", type=int, default=40, nargs='?')
parser.add_argument("--inHeight", "-ih", help="Height in mm of the scaled input image", type=int, default=40, nargs='?')
parser.add_argument("--datasheet", "-d", help="Input data sheet file name", nargs='?') #Required argument
parser.add_argument("--imagefile", "-i", help="Input image file name", nargs='?') #Required argument
parser.add_argument("--fontfile", "-f", help="Truetype font file name", nargs='?')  #Required argument
parser.add_argument("--fontsize", "-s", help="Truetype font size (pt)", type=int, default=18, nargs='?') # default 18 pt

#parser.parse_args()

args = parser.parse_args()

if args.outWidth:
    print("Width in mm of the output image: %s mm", args.outWidth)
    logging.info("Width in mm of the output image: %s mm", args.outWidth)
if args.outHeight:
    print("Height in mm of the output image: %s mm", args.outHeight)
    logging.info("Height in mm of the input image: %s mm", args.outHeight)
if args.inWidth:
    print("Width in mm of rescaled base image: %s mm", args.inWidth)
    logging.info("Width in mm of rescaled base image: %s mm", args.inWidth)
if args.inHeight:
    print("Height in mm of the rescaled base image: %s mm", args.inHeight)
    logging.info("Height in mm of the rescaled base image: %s mm", args.inHeight)
if args.imagefile:
    print("Input data sheet file name: %s", args.datasheet)
    logging.info("Input data sheet file name: %s", args.datasheet)
if args.fontfile:
    print("Input image file name: %s", args.imagefile)
    logging.info("Input image file name: %s", args.imagefile)
if args.fontfile:
    print("Truetype font file name: %s", args.fontfile)
    logging.info("Truetype font file name: %s", args.fontfile)
if args.fontsize:
    print("Truetype font size (pt): %s\n", args.fontsize)
    logging.info("Truetype font size (pt): %s\n", args.fontsize)

    
################################################
# Obtaining data from sheet's first column
################################################

# Reading the table as an ordered dictionary
my_dict = pyexcel.get_dict(file_name=args.datasheet, name_columns_by_row=0)

# Some test...
# Printing in the log file the readed lines
for i in range (1, len(my_dict["Nombre"])):
    logging.debug(my_dict["Nombre"][i])



################################################
# Defining new canvas with the desired output size
################################################

# Conversion from mm to pixels (1 mm = 3.7795275591 pixel at 96 dpi)
canvasPixwidth  = int(3.7795*int(args.outWidth))
canvasPixheight = int(3.7795*int(args.outHeight))
canvas=Image.new('RGBA', (canvasPixwidth, canvasPixheight), 'white')
#canvas.show()

# Loading base image file
imBase = Image.open(args.imagefile)
#print(imOriginal.format, imOriginal.size, imOriginal.mode)
#imBase.show()

# Putting image base in white background.
imBaseWhiteBack = Image.new("RGBA", imBase.size, "WHITE") # Create a white rgba background
imBaseWhiteBack.paste(imBase, (0, 0), imBase)              # Paste the image on the background. Go to the links given below for details.
imBaseWhiteBack.convert('RGB')
#imBaseWhiteBack.show()

# Rescaling base image according to input parameters
baseRescaledWidth = int(3.7795*int(args.inWidth))
baseRescaledHeight = int(3.7795*int(args.inHeight))
imBaseResized = imBaseWhiteBack.resize((baseRescaledWidth ,baseRescaledHeight))

##print(imBase.format, imBase.size, imBase.mode)
#baseWidth, baseHeight = imBase.size
##print(imBase.format, imBase.size, imBase.mode)

# To add 2D graphics in an image call draw Method
drawBase = ImageDraw.Draw(imBase)

# Using input arguments for font definition
TTfont = ImageFont.truetype(args.fontfile, args.fontsize)


    
#########################################################    
# Making a 2D array of images. Filling the matrix by rows
#########################################################    

#Defining small offset in pixels between images to ease laser cuts
offset=4 #pixels

#How many images can I put in a row with an offset?
nImByRow=int(canvasPixwidth/(baseRescaledWidth+offset))
#How many images can I put in a column with an offset?
nImByCol=int(canvasPixheight/(baseRescaledHeight+offset))
print ("With these sizes we can put", nImByRow, "images per row, and", nImByRow, "images per column, rounding down")
logging.info("With these sizes we can put", nImByRow, "images per row, and", nImByRow, "images per column, rounding down")


print("Making a total of %s possible elements in total,\n" , int(nImByCol*nImByRow))
logging.info("Making a total of %s possible elements in total,\n" , int(nImByCol*nImByRow))
print("Our data sheet has %s elements. \n" % len(my_dict["Nombre"]))
logging.info("Our data sheet has %s elements. \n" , len(my_dict["Nombre"]))
# Do we have enough room for the elements in the table?

if (len(my_dict["Nombre"])>int(nImByCol*nImByRow)):
    print("We do not have enough room for all the table elements. Exiting...\n")
    logging.error("We do not have enough room for all the table elements. Exiting...\n")
    exit()
else:
    print("We have %s remaining positions in our canvas.\n" % int( int(nImByCol*nImByRow)- len(my_dict["Nombre"]) ) )
    logging.info("We have %s remaining positions in our canvas.\n" % int( int(nImByCol*nImByRow)- len(my_dict["Nombre"]) ) )

#Defining array of images    
for i in range (len(my_dict["Nombre"])):
    #row identification
    irow= i % nImByRow
    #column identification
    icol= i // nImByRow
    # putting one base image in the correct location
    canvas.paste( imBaseResized, ( int( ( baseRescaledWidth + offset )*irow ) , int( ( baseRescaledHeight +offset) *icol ) ) ) 
    #putting the corresponding text in the center, depending on its size
    draw = ImageDraw.Draw(canvas)
    # obtaining this text size
    textWidth, textHeight = draw.textsize(my_dict["Nombre"][i], font=TTfont)
    thisTextCenterX=int((baseRescaledWidth  + offset)*irow+(baseRescaledWidth-textWidth)/2)
    thisTextCenterY=int((baseRescaledHeight + offset)*icol+(baseRescaledHeight-textHeight)/2)
    draw.text((thisTextCenterX, thisTextCenterY), my_dict["Nombre"][i], font=TTfont,fill=(0,0,0,255) ) #fill letters in black

# saving to .png to retain transparent background
canvas.save('labelarray_ouput.png')

# saving to .jpg
canvas=canvas.convert('RGB')
# saving the label created label array
canvas.save('labelarray_ouput.jpg')

print("Array done, showing result after file save. Exiting...")
logging.info("Array done, showing result after file save. Exiting...")

canvas.show()
