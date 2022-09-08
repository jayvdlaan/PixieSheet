# PixieSheet

![GitHub last commit](https://img.shields.io/github/last-commit/Teitoku42/PixieSheet)
![GitHub top language](https://img.shields.io/github/languages/top/Teitoku42/PixieSheet)
![GitHub](https://img.shields.io/github/license/Teitoku42/PixieSheet)

This script was created to easily convert pixel arts to spreadsheets as to easily coordinate pixeling efforts for r/PlaceNL. Later functionality was also implemented that allows for the conversion from pixelart spreadsheets back to images so that these could easily be imported in the r/PlaceNL command & control center in order to be streamed to the botnet.

## Features
- Conversion from spreadsheet to image
  - Supports different spreadsheet color types (RGB, theme, indexed)

- Conversion from image to spreadsheet
  - Single pixel mode
  - RGB header for easy determination of pixel art start point and pixel size
  - Specify absolute coordinates of to-be pixel art on r/Place
  - Specify cell width and height

## Usage
- -m - **(Required) <Mode>** - Specify whether you wanna convert image to spreadsheet or spreadsheet to image (values: tosheet, fromsheet)
- -i - **(Required) <Input>** - Specify the path to the input file (mode: fromsheet = xlsx, tosheet = png)
- -o - **(Required) <Output>** - Specify the name of the generated video (mode: fromsheet = png, tosheet = xlsx)
- -x <First pixel X> - Specify the X coordinate of the first pixel to place on r/place
- -y <First pixel Y> - Specify the Y coordinate of the first pixel to place on r/place
- -l <Cell width> - Specify the desired cell width (default is 2.57).
- -b <Cell height> - Specify the desired cell height (default is 15.75).
- -s <Single pixel mode> - Run the script in single pixel mode, this will assume that your image already has 1x1 pixels, this removes the need for a header.

## Limitations
- Conversion from spreadsheet to image
  - Spreadsheet must be of type `.xlsx`
  - Spreadsheet must have numbered axis (cells do not have to be numbered, just need to have text in them spanning the art's dimensions)
  - Convertible spreadsheet: 
    
    <img src="https://i.imgur.com/KJ5VRR5.png" width="500" height="400">

- Conversion from image to spreadsheet
  - The input image must be a non-transparent PNG image
  - Single pixel mode
  
    - The pixels portrayed in the input image must span only a single pixel (1x1)
  - Multi pixel mode (default)
    - Image must contain a header (details found below)
    - Pixels must be of matching width and height across entire image, example: if one pixel has the dimensions 12x12, then all pixels must bear those same dimensions
 
## RGB Header
The RGB header was introduced as a means to both determine the pixel size of an image while simultaneously marking the beginning of the pixel art to extract.
This header can be inserted into an image through various means, but the easiest is by using an image manipulation utility like Photoshop or Gimp. A header must meet the following requirements:
- Must consist of three similarly sized pixels
- Going from left to right, the pixels must have the following color values

  - Red (#FF0000)
  - Green (#00FF00)
  - Blue (#0000FF)
- Pixels must be placed be placed immediately to the right of the previous pixel
- The red pixel of the header must be placed on the first top-left pixel of the desired pixel art

### Convertible image
<img src="https://i.imgur.com/oCBfJxt.png" width="500" height="500">

### Converted image
<img src="https://i.imgur.com/a3le0uL.png" width="500" height="500">
