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
import numpy
import getopt
from PIL import Image


def load_image(path: str) -> Image:
    return Image.open(path)


def get_image_info(image: Image) -> dict:
    out_info = {"begin": [], "pixel-size": []}
    header_colors = [[255, 0, 0], [0, 255, 0], [0, 0, 255]]

    # Search image for header pixels (red green and blue)
    header_found = False;
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
                        size_counter = size_counter + 1
                        # Check if we're entering a pixel that belongs to the header and is not the final one
                        if header_check_index + 1 < len(header_colors) and \
                                numpy.array_equal(next_pixel, header_colors[header_check_index + 1]):
                            if size_counter != previous_size and header_check_index > 0:
                                raise RuntimeError("Header pixels differ in size")
                            else:
                                previous_size = size_counter
                                size_counter = 0
                                header_check_index = header_check_index + 1
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


def image_to_sheet(image_path: str):
    input_image = load_image(image_path)
    image_info = get_image_info(input_image)


def main(cmdline):
    try:
        options, args = getopt.getopt(cmdline, "hi:")
    except getopt.GetoptError as err:
        print("PixieSheet.py -h for instructions")
        return

    if len(options) == 0:
        print("PixieSheet.py -h for instructions")
        return

    image_path = ""
    for option, argument in options:
        if option == "-i":
            image_path = argument
        else:
            print("Ain't nothin' ere yet boy")
            return

    if image_path != "":
        image_to_sheet(image_path)


if __name__ == '__main__':
    main(sys.argv[1:])
