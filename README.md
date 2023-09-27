# from_ods_to_imagematrix

Extract contents from a Open Office sheet's columns and convert it to text in image arrays.
Can rescale the input image with command line arguments. 

## Requires

- Python 3

- Libraries:
  - [pyexcel](http://docs.pyexcel.org/en/latest/): to read and write libre office and excel sheets
  - [PIL](https://pypi.org/project/Pillow/): to modify and create images with texts
  - [argparse](https://docs.python.org/3/library/argparse.html): to read arguments from command line

Can all be installed in one step using:

`pip install -r requirements.txt`

## Usage

From command line:

`python3 ./from_ods_to_imagematrix_V5.py [ARGUMENTS]`

Arguments:

-  -h, --help            show this help message and exit

- --outWidth [OUTWIDTH], -ow [OUTWIDTH]
                        Width in mm of the output image

- --outHeight [OUTHEIGHT], -oh [OUTHEIGHT]
                        Height in mm of the output image

- --inWidth [INWIDTH], -iw [INWIDTH]
                        Width in mm for the scaled input image

- --inHeight [INHEIGHT], -ih [INHEIGHT]
                        Height in mm of the scaled input image

- --datasheet [DATASHEET], -d [DATASHEET]
                        Input data sheet file name

- --imagefile [IMAGEFILE], -i [IMAGEFILE]
                        Input image file name

- --fontfile [FONTFILE], -f [FONTFILE]
                        Truetype font file name

- --fontsize [FONTSIZE], -s [FONTSIZE]
                        Truetype font size (pt)

Example for Linux command line for one field matrix generation:

`python3 ./from_ods_to_imagematrix_V6.py -ow 1000 -oh 800 -iw 43 -ih 43 -d datasheets/table_names_1.ods -i images/base_llavero_V0.1.png -f  fonts/Misyalli-dafont.ttf -s 24`

Will generate an output image file with 1000x800mm, composed of a matrix of 43x43mm images with 24 point of letter size.

Using:

- `images/base_llavero_V0.1.png`:. as input image.

- `fonts/Misyalli-dafont.ttf`: as font for the new text.

- `datasheets/table_names_1.ods`: as data input file.

<p align="center">
<img width="800" src="https://github.com/AGordiGuerrero/from_ods_to_imagematrix/blob/master/labelarray_output.jpg">
</p>

Example for Windows command line for two field matrix generation:

`python ./from_ods_to_imagematrix_twofieldsV1.py -ow 500 -oh 600 -iw 140 -ih 70 -d datasheets/table_name_surname_1.ods -i images/cartel_ponente_140x90mm_EDD_01.png -f  fonts/Orbitron-VariableFont_wght.ttf -s 42`

Using:

- `images/cartel_ponente_140x90mm_EDD_01.png` :. as input image.

- `fonts/Orbitron-VariableFont_wght.ttf`: as font for the new text.

- `datasheets/table_name_surname_1.ods`: as data input file.

<p align="center">
<img width="800" src="https://github.com/AGordiGuerrero/from_ods_to_imagematrix/blob/master/labelarray_output_twofields.jpg">
</p>

----------------

Created by A.Gordillo-Guerrero (anto@unex.es), September , 2023.

[Smart Open Lab](www.smartopenlab.com). University of Extremadura. Spain.

Released into the public domain.
