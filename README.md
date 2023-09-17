# from_ods_to_imagematrix

Extract contents from a Open Office sheet's columns and convert it to text in image arrays.
Can rescale the input image with command line arguments. 

Created by A.Gordillo-Guerrero (anto@unex.es), September , 2023.
University of Extremadura. Smart Open Lab.
Released into the public domain.

## Requires

- Python 3

- Libraries:
 - pyexcel: to read and write libre office and excel sheets
 - PIL: to modify and create images with texts
 - argparse: to read arguments from command line

## Usage

Arguments:

  -h, --help            show this help message and exit

--outWidth [OUTWIDTH], -ow [OUTWIDTH]
                        Width in mm of the output image

--outHeight [OUTHEIGHT], -oh [OUTHEIGHT]
                        Height in mm of the output image

--inWidth [INWIDTH], -iw [INWIDTH]
                        Width in mm for the scaled input image

--inHeight [INHEIGHT], -ih [INHEIGHT]
                        Height in mm of the scaled input image

--datasheet [DATASHEET], -d [DATASHEET]
                        Input data sheet file name

--imagefile [IMAGEFILE], -i [IMAGEFILE]
                        Input image file name

--fontfile [FONTFILE], -f [FONTFILE]
                        Truetype font file name

--fontsize [FONTSIZE], -s [FONTSIZE]
                        Truetype font size (pt)

Example:

`python3 ./from_ods_to_imagematrix_V5.py -ow 1000 -oh 800 -iw 43 -ih 43 -d datasheets/table_names_1.ods -i images/base_llavero_V0.1.png -f  fonts/Misyalli-dafont.ttf` -s 24

Will generate an image output file with 1000x800mm, composed of a matrix of 43x43mm images with 24 point of letter size. Using:

- `images/base_llavero_V0.1.png`:. as input image.

- `fonts/Misyalli-dafont.ttf`: as font for the new text.

- `datasheets/table_names_1.ods`: as data input file.

