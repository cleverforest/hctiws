#!/usr/bin/env python3

import sys
import os
import csv
from typing import Tuple
import toml
import xlrd
from PIL import Image, ImageFont, ImageDraw
import math

COMPRESS_BASE_IMAGE_MULTIPLE = 5

def find_file(filename: str, style_dir="") -> str:
    """Finding file in possible directories"""
    tmp_filename = os.path.join(INPUT_DIR, filename)
    if os.path.exists(tmp_filename):
        return tmp_filename
    tmp_filename = os.path.join(style_dir, filename)
    if os.path.exists(tmp_filename):
        return tmp_filename
    tmp_filename = os.path.join(os.path.dirname(
        os.path.dirname(os.getcwd())), filename)
    if os.path.exists(tmp_filename):
        return tmp_filename
    raise NameError("File not found")


def get_style(style_name: str) -> dict[str,]:
    """Getting the style information"""
    style_filename = os.path.join("config", style_name, "layout.")
    if os.path.exists(style_filename + "toml"):
        style_filename += "toml"
    elif os.path.exists(style_filename + "ini"):
        style_filename += "ini"
    else:
        raise NameError("Layout file not found")
    
    # some compatibility stuff
    ret = toml.load(style_filename)
    if "_meta" not in ret:
        ret["_meta"] = { "layout_version": 0 }
    elif "version" not in ret["_meta"]:
        ret["_meta"]["layout_version"] = 1
    layout_ver = ret["_meta"]["layout_version"]

    if layout_ver < 2 and "figure_alias" in ret:
        ret["_meta"]["figure_alias"] = ret.pop("figure_alias")
    return ret


def get_text_size(text: str, font: str, size: list[int], offset=0,
                  enable_compress = False) -> Tuple[ImageFont.ImageFont]:
    """Getting the proper size in point of the text"""
    tmp_size1 = size[1]
    tmp_font = ImageFont.truetype(font, tmp_size1)
    # while tmp_font.getsize(text)[0] <= size[0] and \
    #        tmp_font.getsize(tmp_text)[1] <= size[1] / (1 - offset):
    #    tmp_size += 1
    #    tmp_font = ImageFont.truetype(font, tmp_size)
    #tmp_size -= 1
    #tmp_font = ImageFont.truetype(font, tmp_size)
    #print (text, " ", tmp_size, " ", size)
    if (enable_compress and allow_compress_text):
        if (tmp_font.getsize(text)[0] > size[0]):
            tmp_size1 *= COMPRESS_BASE_IMAGE_MULTIPLE
            tmp_font = ImageFont.truetype(font, tmp_size1)
            return (tmp_font, tmp_size1, True)
        else:
            return (tmp_font, tmp_size1, False)
    else:
        while tmp_font.getsize(text)[0] > size[0]:
            if tmp_size1 == 1:
                break
            tmp_size1 -= 1
            tmp_font = ImageFont.truetype(font, tmp_size1)
        return (tmp_font, tmp_size1, False)
    # the last boolean value indicates that the text needs to be compressed

def draw_text(s_img: Image.Image, text: str, font: str, color: str, pos: list[int], size: list[int],
              h_align="left", v_align="top", offset=0) -> Image.Image:
    """Drawing the text to the image"""
    new_img = s_img
    try:
        font_loc = find_file(font)
    except NameError:
        font_loc = font
    get_text_size_ret = get_text_size(text, font_loc, size, offset, True)
    tmp_font = get_text_size_ret[0]
    need_compress = get_text_size_ret[2]

    if (need_compress):
        large_text_size = tmp_font.getsize(text)
        large_text_size = (math.ceil(large_text_size[0]), math.ceil(large_text_size[1]))
        tmp_canvas = Image.new("RGBA", large_text_size)
        text += "　" #for Source Han Sans
        if (color == None or color == ""):
            color = default_color
        ImageDraw.Draw(tmp_canvas).text((0, 0), text, fill=color, font=tmp_font)
        #size1_add_offset = size[1] / (1 - offset)
        # Please not the offset support in COMPRESSED TEXT still has issues in this version
        if (large_text_size[0] * compress_text_ratio / large_text_size[1] > size[0] / size[1]):
            new_canvas_size = (size[0],
                         math.floor(size[0] / compress_text_ratio * (large_text_size[1] / large_text_size[0])))
            text_size = [new_canvas_size[0], new_canvas_size[1] * (1 - offset)]
        else:
            new_canvas_size = (min(size[0], math.ceil(size[1] / large_text_size[1] * large_text_size[0])), size[1])
            text_size = [new_canvas_size[0], size[1]]
        tmp_canvas = tmp_canvas.resize(new_canvas_size)

        if (pos[0] < 0 and max_width != 0):
            pos[0] = max_width + pos[0] - size[0]
        tmp_pos = [pos[0], pos[1] - tmp_font.getsize(text)[1] * offset]

        if h_align == "center":
            tmp_pos[0] += (size[0] - text_size[0]) / 2
        elif h_align == "right":
            tmp_pos[0] += size[0] - text_size[0]
        if v_align == "center":
            tmp_pos[1] += (size[1] - text_size[1]) / 2
        elif v_align == "bottom":
            tmp_pos[1] += size[1] - text_size[1]
        new_img.paste(tmp_canvas, (int(tmp_pos[0]),
                                   int(tmp_pos[1])), tmp_canvas)
    else:
        text_size = [tmp_font.getsize(text)[0],
                    tmp_font.getsize(text)[1] * (1 - offset)]

        if (pos[0] < 0 and max_width != 0):
            pos[0] = max_width + pos[0] - size[0]
        tmp_pos = [pos[0], pos[1] - tmp_font.getsize(text)[1] * offset]

        if h_align == "center":
            tmp_pos[0] += (size[0] - text_size[0]) / 2
        elif h_align == "right":
            tmp_pos[0] += size[0] - text_size[0]
        if v_align == "center":
            tmp_pos[1] += (size[1] - text_size[1]) / 2
        elif v_align == "bottom":
            tmp_pos[1] += size[1] - text_size[1]
        text += "　" #for Source Han Sans
        #if all(ord(c) < 128 for c in text): #for Source Han Sans
        #    tmp_pos[1] -= 2
        if (color == None or color == ""):
            color = default_color
        ImageDraw.Draw(new_img).text(tmp_pos, text, fill=color, font=tmp_font)
        #print (tmp_size," ",text_size," ", size[0],"*",size[1])
    return new_img


def generate_text(s_img: Image.Image,
                  item_list: list[str],
                  index: int,
                  para_list: dict[str,],
                  meta_conf: dict[str,]) -> Tuple[Image.Image, int]:
    """Generating image of 'text' type"""
    #print(index, ",", para_list["position"])
    new_img = draw_text(s_img, item_list[index],
                        para_list["font"], para_list["color"], para_list["position"],
                        [para_list["width"], para_list["height"]],
                        para_list["horizontal_align"],
                        para_list["vertical_align"], para_list.get("offset", 0))
    return (new_img, index + 1)


def generate_colortext(s_img: Image.Image,
                       item_list: list[str],
                       index: int,
                       para_list: dict[str,],
                       meta_conf: dict[str,]) -> Tuple[Image.Image, int]:
    """Generating image of 'colortext' type"""
    new_img = draw_text(s_img, item_list[index + 1],
                        para_list["font"], item_list[index], para_list["position"],
                        [para_list["width"], para_list["height"]],
                        para_list["horizontal_align"],
                        para_list["vertical_align"], para_list.get("offset", 0))
    return (new_img, index + 2)


def generate_vertitext(s_img: Image.Image,
                       item_list: list[str],
                       index: int,
                       para_list: dict[str,],
                       meta_conf: dict[str,]) -> Tuple[Image.Image, int]:
    """Generating image of 'vertitext' type"""
    text = item_list[index]
    text_len = len(text)
    [tmp_font, tmp_size] = get_text_size(text[0], para_list["font"],
                                         [para_list["width"], int(para_list["height"] /
                                                                  text_len)],
                                         para_list.get("offset", 0))
    for i in text[1:]:
        tmp_fontset = get_text_size(i, para_list["font"],
                                    [para_list["width"], int(para_list["height"] /
                                                             text_len)],
                                    para_list.get("offset", 0))
        if tmp_fontset[1] >= tmp_size:
            [tmp_font, tmp_size] = tmp_fontset
    text_size = [0, 0]
    v_step = []
    for i in text:
        single_size = [tmp_font.getsize(i)[0],
                       tmp_font.getsize(i)[1] * (1 - para_list.get("offset", 0))]
        text_size[0] = max(text_size[0], single_size[0])
        text_size[1] += single_size[1]
        v_step.append(single_size[1])
        if i != text[-1]:
            text_size[1] += para_list["space"]
            v_step[-1] += para_list["space"]
    if para_list["vertical_align"] == "center":
        cur_v_pos = int((para_list["height"] - text_size[1]) / 2)
    elif para_list["vertical_align"] == "bottom":
        cur_v_pos = para_list["height"] - text_size[1]
    else:
        cur_v_pos = 0
    cur_v_pos += para_list["position"][1]
    for i in range(text_len):
        new_img = draw_text(s_img, text[i], para_list["font"], para_list["color"],
                            [para_list["position"][0], int(cur_v_pos)],
                            [para_list["width"], int(v_step[i])],
                            para_list["horizontal_align"], "center",
                            para_list.get("offset", 0))
        cur_v_pos += v_step[i]
    return (new_img, index + 1)


def generate_doubletext_nl(s_img: Image.Image,
                           item_list: list[str],
                           index: int,
                           para_list: dict[str,],
                           meta_conf: dict[str,]) -> Tuple[Image.Image, int]:
    """Generating image of 'doubletext_nl' type"""
    single_height = int((para_list["height"] - para_list["space"])/2)
    new_img = draw_text(s_img, item_list[index],
                        para_list["font"], para_list["color"], para_list["position"],
                        [para_list["width"],
                         single_height], para_list["horizontal_align"],
                        para_list["vertical_align"], para_list.get("offset", 0))
    new_img = draw_text(s_img, item_list[index + 1],
                        para_list["font"], para_list["color"],
                        [para_list["position"][0], para_list["position"]
                         [1] + single_height + para_list["space"]],
                        [para_list["width"],
                         single_height], para_list["horizontal_align"],
                        para_list["vertical_align"], para_list.get("offset", 0))
    return (new_img, index + 2)


def generate_doubletext(s_img: Image.Image,
                        item_list: list[str],
                        index: int,
                        para_list: dict[str,],
                        meta_conf: dict[str,]) -> Tuple[Image.Image, int]:
    """Generating image of 'doubletext' type"""
    if item_list[index + 1] == "":
        return [generate_text(s_img, item_list, index, para_list)[0], index + 2]
    i_size = [para_list["width"], para_list["height"]]
    [major_font, major_size] = get_text_size(item_list[index], para_list["font"],
                                             i_size, para_list.get("offset", 0))
    major_text_size = major_font.getsize(item_list[index])
    [minor_font, minor_size] = get_text_size(item_list[index + 1], para_list["font"],
                                             [i_size[0] - major_text_size[0] - para_list["space"],
                                              i_size[1]], para_list.get("offset", 0))
    while minor_size < para_list["minimum_point"] or \
            major_size - minor_size > para_list["maximum_diff"]:
        major_size -= 1
        major_font = ImageFont.truetype(para_list["font"], major_size)
        major_text_size = major_font.getsize(item_list[index])
        [minor_font, minor_size] = get_text_size(item_list[index + 1], para_list["font"],
                                                 [i_size[0] - major_text_size[0] -
                                                  para_list["space"], i_size[1]],
                                                 para_list.get("offset", 0))
    if major_size - minor_size <= para_list["minimum_diff"]:
        minor_size = max(para_list["minimum_point"],
                         major_size - para_list["minimum_diff"])
    minor_font = major_font = ImageFont.truetype(para_list["font"], minor_size)
    minor_text_size = minor_font.getsize(item_list[index + 1])
    new_img = draw_text(s_img, item_list[index],
                        para_list["font"], para_list["color"], para_list["position"],
                        [major_text_size[0], para_list["height"]], "left",
                        para_list["vertical_align"], para_list.get("offset", 0))
    minor_pos = para_list["position"].copy()
    minor_pos[0] += major_text_size[0] + para_list["space"]
    minor_textsize = [minor_text_size[0], para_list["height"]].copy()
    new_img2 = draw_text(new_img, item_list[index + 1],
                         para_list["font"], para_list["color"], minor_pos,
                         minor_textsize, para_list["horizontal_align_right"],
                         para_list["vertical_align"], para_list.get("offset", 0))
    #print (major_size,", ",minor_size,", ",item_list[index])
    return (new_img2, index + 2)


def generate_figure(s_img: Image.Image,
                    item_list: list[str],
                    index: int,
                    para_list: dict[str,],
                    meta_conf: dict[str,]) -> Tuple[Image.Image, int]:
    """Generating image of 'figure' type"""
    try:
        if item_list[index] == "":
            return [s_img, index + 1]
    except KeyError:
        return [s_img, index + 1]
    #print(para_list, ", ", item_list[index])
    new_img = s_img
    try:
        imgfile = find_file(item_list[index])
    except NameError:
        imgfile = meta_conf["figure_alias"][item_list[index]]
    fig: Image.Image = Image.open(imgfile)
    
    if (para_list["position"][0] < 0 and max_width != 0):
        para_list["position"][0] = max_width + para_list["position"][0] - para_list["width"]
    
    if para_list["keep_aspect_ratio"] == 0:
        fig = fig.resize((para_list["width"], para_list["height"]))
        pos = para_list["position"]
    else:
        fig.thumbnail((para_list["width"], para_list["height"]))
        pos = para_list["position"].copy()
        if para_list["horizontal_align"] == "center":
            pos[0] += int((para_list["width"] - fig.size[0]) / 2)
        elif para_list["horizontal_align"] == "right":
            pos[0] += para_list["width"] - fig.size[0]
        if para_list["vertical_align"] == "center":
            pos[1] += int((para_list["height"] - fig.size[1]) / 2)
        elif para_list["vertical_align"] == "bottom":
            pos[1] += para_list["height"] - fig.size[1]
    new_img.paste(fig, pos, fig.convert("RGBA"))
    return (new_img, index + 1)


# def generate_figuregroup(s_img, item_list, index, para_list):
#    """Generate image of 'figuregroup' type"""
#    new_img = s_img
#    return [new_img, index+1]


def process_section(style: dict[str,],
                    layout_type: str,
                    item_list: list[str]) -> Image.Image:
    """Processing a single section"""
    switchfunc = {
        "text": generate_text,
        "colortext": generate_colortext,
        "vertitext": generate_vertitext,
        "doubletext_nl": generate_doubletext_nl,
        "doubletext": generate_doubletext,
        "figure": generate_figure
        #"figuregroup": generate_figuregroup
    }
    if layout_type.startswith('_'):
        raise NameError("Invalid type name")
    layout = style[layout_type]
    s_img = Image.open(find_file(layout["background"]))
    index = 0
    for i in layout["item"]:
        try:
            i_conf = layout["default_" + i["type"]].copy()
            i_conf.update(i)
        except KeyError:
            i_conf = i.copy()
        [s_img, index] = switchfunc[i["type"]](
            s_img, item_list, index, i_conf, style["_meta"])
    return s_img


def process_sheet(s_sheet: list[list]):
    """Processing a single sheet"""
    s_index = 0
    s_len = len(s_sheet)
    if (len(s_sheet[s_index]) == 0):
        return
    while s_sheet[s_index][0] == "//":
        s_index += 1
        if s_index == s_len:
            raise NameError("No valid style")
    style_name = s_sheet[s_index][0]
    style = get_style(style_name)

    # loading meta configs
    global max_width
    global default_color
    global allow_compress_text
    global compress_text_ratio
    max_width = 0
    default_color = "black"
    allow_compress_text = False
    compress_text_ratio = 1
    if ("_meta" in style):
        if ("max_width" in style["_meta"]):
            max_width = style["_meta"]["max_width"]
        if ("default_color" in style["_meta"]):
            default_color = style["_meta"]["default_color"]
        if ("allow_compress_text" in style["_meta"] and "compress_text_ratio" in style["_meta"]):
            allow_compress_text = style["_meta"]["allow_compress_text"]
            compress_text_ratio = style["_meta"]["compress_text_ratio"]
            if ((not allow_compress_text)
                or (type(compress_text_ratio) != int and type(compress_text_ratio) != float)
                or compress_text_ratio > 1
                or compress_text_ratio < 0):
                compress_text_ratio = 1
    
    os.chdir(os.path.join("config", style_name))
    s_index += 1
    current_type = "default"
    main_img = None
    main_size = [0, 0]
    for i in s_sheet[s_index:]:
        if i[0] == "//":
            continue
        if i[0] != "":
            current_type = i[0]
        s_img = process_section(
            style, current_type, i[1:])
        s_size = list(s_img.size)
        if main_img is None:
            main_img = s_img
            main_size = s_size
        else:
            new_size = [max(main_size[0], s_size[0]), main_size[1] + s_size[1]]
            main_img = main_img.crop([0, 0] + new_size)
            new_box = [0, main_size[1], s_size[0], new_size[1]]
            main_img.paste(s_img, new_box)
            main_size = new_size
    os.chdir(os.path.dirname(os.path.dirname(os.getcwd())))
    return main_img


def read_csv_file(csv_filename: str) -> dict[str, list[list]]:
    """Reading CSV file"""
    c_file = open(csv_filename, "r")
    content = list(csv.reader(c_file))
    c_file.close()
    return {"csvsheet": content}


def read_excel_file(excel_filename: str) -> dict[str, list[list]]:
    """Reading .xlsx/.xls file"""
    x_book = xlrd.open_workbook(excel_filename)
    content = {}
    tmp_sheetname = x_book.sheet_names()
    for i in range(x_book.nsheets):
        tmp_content = []
        tmp_sheet = x_book.sheet_by_index(i)
        for j in range(tmp_sheet.nrows):
            tmp_content += [tmp_sheet.row_values(j)]
        content[tmp_sheetname[i]] = tmp_content
    return content


def valid_filename(input_filename: str) -> str:
    """Making filename valid"""
    return input_filename.translate(str.maketrans("*/\\<>:\"|", "--------"))


def main(argv=None):
    """Main function of HCTIWS"""
    global INPUT_DIR
    if argv is None:
        argv = sys.argv
    # display the version info
    print("HCTIWS Creates the Image with Sheets")
    print("       (C) ZMSOFT 2018-2025")
    print("version 2.80\n")
    # get input filename
    if len(argv) == 1:
        print("Usage: hctiws [INPUT_FILE [OUTPUT_DIRECTORY]]\n")
        input_filename = input("Input file (.csv, .xlsx, .xls): ")
        INPUT_DIR = os.path.dirname(input_filename)
    elif len(argv) == 2:
        input_filename = argv[1]
        INPUT_DIR = os.path.dirname(input_filename)
    else:
        input_filename = argv[1]
        INPUT_DIR = argv[2]
    # open worksheet/sheet file
    if input_filename[-4:] == ".csv":
        content = read_csv_file(input_filename)
    elif input_filename[-5:] == ".xlsx" or input_filename[-4:] == ".xls":
        content = read_excel_file(input_filename)
    # process and save the result of each sheet
    for i in content:
        tmp_img = process_sheet(content[i])
        # tmp_img.show()  # show the image before saving for debug
        tmp_name = os.path.join(
            INPUT_DIR, valid_filename(
                os.path.basename(input_filename)).
            rsplit(".", 1)[0] + "_" + valid_filename(i))
        if os.path.exists(tmp_name + ".png"):
            for j in range(1, 100):
                if not os.path.exists(tmp_name + "-" + str(j) + ".png"):
                    tmp_name += "-" + str(j) + ".png"
                    break
            else:
                raise NameError("Too many duplicated file names")
        else:
            tmp_name += ".png"
        tmp_img.save(tmp_name, format="png") #, optimize=True) #compress_level=0)
        print(tmp_name, "DONE")
    return 0


INPUT_DIR = ""
if __name__ == "__main__":
    sys.exit(main())
