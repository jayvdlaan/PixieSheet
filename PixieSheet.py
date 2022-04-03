#  Copyright (c) 2022 Teitoku42. All Rights Reserved.
#
#  Permission is hereby granted, free of charge, to any person obtaining a copy
#  of this software and associated documentation files (the "Software"), to deal
#  in the Software without restriction, including without limitation the rights
#  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#  copies of the Software, and to permit persons to whom the Software is
#  furnished to do so, subject to the following conditions:
#
#  The above copyright notice and this permission notice shall be included in all
#  copies or substantial portions of the Software.
#
#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#  SOFTWARE.

import array
import sys

import xlsxwriter
import xlsxwriter.utility
import numpy
import getopt
import math
from PIL import Image


# TODO: Organize functions into classes

# Loads the passed image into a PIL image object
def load_image(path: str) -> Image:
    return Image.open(path)


# Obtains image starting coordinates and size of pixels
def get_image_info(image: Image) -> dict:
    out_info = {"begin": [], "pixel-size": []}
    header_colors = [[255, 0, 0], [0, 255, 0], [0, 0, 255]]

    # Search image for header pixels (red green and blue)
    header_found = False
    header_check_index = 0
    previous_size = 0
    size_counter = 0
    for y in range(image.height):
        if header_found:
            break
        for x in range(image.width):
            current_pixel = image.getpixel((x, y))
            if header_found:
                break
            # Found initial header pixel
            if numpy.array_equal(current_pixel, header_colors[header_check_index]):
                out_info["begin"] = [x, y]
                for inner_x in range(x, image.width):
                    current_pixel = image.getpixel((inner_x, y))
                    next_pixel = [-1, -1, -1]
                    if inner_x + 1 < image.width:
                        next_pixel = image.getpixel((inner_x + 1, y))
                    # Check if we are still counting the same pixel
                    if numpy.array_equal(current_pixel, header_colors[header_check_index]):
                        size_counter += 1
                        # Check if we're entering a pixel that belongs to the header and is not the final one
                        if header_check_index + 1 < len(header_colors) and \
                                numpy.array_equal(next_pixel, header_colors[header_check_index + 1]):
                            if size_counter != previous_size and header_check_index > 0:
                                raise RuntimeError("Header pixels differ in size")
                            else:
                                previous_size = size_counter
                                size_counter = 0
                                header_check_index += 1
                        # Check if we're entering a pixel that is not part of header colors anymore indicating finish
                        elif header_check_index + 1 == len(header_colors) and not \
                                numpy.array_equal(next_pixel, header_colors[header_check_index]):
                            if size_counter != previous_size:
                                raise RuntimeError("Header pixels differ in size")
                            else:
                                out_info["pixel-size"] = size_counter
                                header_found = True
                                break
                    # Behaviour if pixel is not of the expected header color
                    else:
                        raise RuntimeError("Unexpected pixel color encountered when interpreting header")

    return out_info


# Trims passed image to make starting coordinates 0, 0
def trim_image(image: Image, image_info: dict) -> Image:
    left = image_info["begin"][0]
    top = image_info["begin"][1]
    right = image.width
    bottom = image.height
    return image.crop((left, top, right, bottom))


# Calculate pixel hop count
def calculate_pixel_hops(axis_size: int, pixel_size: int) -> int:
    hops = math.floor(axis_size / pixel_size)
    if axis_size % pixel_size > 0:
        hops += 1

    return int(hops)


# Generate a pixel art matrix from the provided image where a pixel of x size is shrunk to a 1x1 pixel
def generate_pixel_map(image: Image, pixel_size: int) -> numpy.ndarray:
    out_array = []
    width_hops = calculate_pixel_hops(image.width, pixel_size)
    height_hops = calculate_pixel_hops(image.height, pixel_size)
    for y in range(height_hops):
        actual_y = y * pixel_size
        y_row = []
        for x in range(width_hops):
            actual_x = x * pixel_size
            y_row.append(image.getpixel((actual_x, actual_y)))
        out_array.append(y_row)

    return numpy.asarray(out_array, dtype="uint8")


def rgb_to_hex(rgb: tuple):
    return "%02x%02x%02x" % rgb


# TODO: Split functionality into different classes
# Converts a pixel map to a spreadsheet
def map_to_sheet(pixel_map: numpy.ndarray,
                 output_path: str,
                 x_begin: int,
                 y_begin: int,
                 column_width: float,
                 row_height: float):
    book = xlsxwriter.Workbook(output_path)
    sheet = book.add_worksheet("Art")

    width = pixel_map.shape[1]
    height = pixel_map.shape[0]
    corner_offset = 1

    sheet.set_column(0, width, column_width)
    sheet.set_default_row(row_height)

    for x in range(width):
        sheet.write(0, x + corner_offset, x_begin + x)

    for y in range(height):
        sheet.write(y + corner_offset, 0, y_begin + y)

    for y in range(height):
        for x in range(width):
            rgb = pixel_map[y][x]
            fmt = book.add_format({"bg_color": "#" + rgb_to_hex((rgb[0], rgb[1], rgb[2]))})
            sheet.write(y + corner_offset, x + corner_offset, "", fmt)

    book.close()


# Performs preprocessing on image, then converts it to a spreadsheet
def image_to_sheet(options: dict):
    image = load_image(options["input"])
    image_info = {"begin": (0, 0), "pixel-size": 1}
    if not options["spm"]:
        image_info = get_image_info(image)
        image = trim_image(image, image_info)
    pixel_map = generate_pixel_map(image, image_info["pixel-size"])
    map_to_sheet(pixel_map,
                 options["output"],
                 options["begin"]["x"],
                 options["begin"]["y"],
                 options["dimensions"]["w"],
                 options["dimensions"]["h"])


def main(cmdline):
    try:
        options, args = getopt.getopt(cmdline, "hi:o:x:y:l:b:s")
    except getopt.GetoptError as err:
        print("PixieSheet.py -h for instructions")
        return

    if len(options) == 0:
        print("PixieSheet.py -h for instructions")
        return

    set_values = {"input": None, "output": None, "begin": {"x": 0, "y": 0}, "dimensions": {"w": 2.57, "h": 15.75}, "spm": False}
    for option, argument in options:
        if option == "-i":
            set_values["input"] = argument
        elif option == "-o":
            set_values["output"] = argument
        elif option == "-x":
            set_values["begin"]["x"] = int(argument)
        elif option == "-y":
            set_values["begin"]["y"] = int(argument)
        elif option == "-l":
            set_values["dimensions"]["w"] = int(argument)
        elif option == "-b":
            set_values["dimensions"]["h"] = int(argument)
        elif option == "-s":
            set_values["spm"] = True
        else:
            print("PixieSheet - Fast and easy pixel art to spreadsheet\n"
                  "\n"
                  "-i (Required) <Image input path> - Specify the image to turn into a spreadsheet.\n"
                  "-o (Required) <Output name> - Specify the path of the output spreadsheet.\n"
                  "-x <First pixel X> - Specify the X coordinate of the first pixel to place on r/place\n"
                  "-y <First pixel Y> - Specify the Y coordinate of the first pixel to place on r/place\n"
                  "-l <Cell width> - Specify the desired cell width (default is 2.57).\n"
                  "-b <Cell height> - Specify the desired cell height (default is 15.75).\n"
                  "-s <Single pixel mode> - Run the script in single pixel mode, this will assume that your image "
                  "already has 1x1 pixels, this removes the need for a header.\n")
            return

    if set_values["input"] is None or set_values["output"] is None:
        print("PixieSheet.py -h for instructions")

    image_to_sheet(set_values)


if __name__ == '__main__':
    main(sys.argv[1:])
