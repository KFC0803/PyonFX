# -*- coding: utf-8 -*-
# PyonFX: An easy way to do KFX and complex typesetting based on subtitle format ASS (Advanced Substation Alpha).
# Copyright (C) 2019 Antonio Strippoli (CoffeeStraw/YellowFlash)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# PyonFX is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this program. If not, see http://www.gnu.org/licenses/.

import os
import sys
import time
import re
import copy
import subprocess
from typing import List
from .font_utility import Font
from .convert import Convert


def pretty_print(obj, indent=0, name=""):
    # Utility function to print object Meta, Style, Line, Word, Syllable and Char (this is a dirty solution probably)
    if type(obj) == Line:
        out = " " * indent + f"lines[{obj.i}] ({type(obj).__name__}):\n"
    elif type(obj) == Word:
        out = " " * indent + f"words[{obj.i}] ({type(obj).__name__}):\n"
    elif type(obj) == Syllable:
        out = " " * indent + f"syls[{obj.i}] ({type(obj).__name__}):\n"
    elif type(obj) == Char:
        out = " " * indent + f"chars[{obj.i}] ({type(obj).__name__}):\n"
    else:
        out = " " * indent + f"{name}({type(obj).__name__}):\n"

    # Let's print all this object fields
    indent += 4
    for k, v in obj.__dict__.items():
        if "__dict__" in dir(v):
            # Work recursively to print another object
            out += pretty_print(v, indent, k + " ")
        elif type(v) == list:
            for i, el in enumerate(v):
                # Work recursively to print other objects inside a list
                out += pretty_print(el, indent, f"{k}[{i}] ")
        else:
            # Just print a field of this object
            out += " " * indent + f"{k}: {str(v)}\n"

    return out


class Meta:
    """Meta对象包含了Ass的信息。

    可以在此获取它们的更多信息: https://aegi.vmoe.info/docs/3.2/Styles/ 。

    Attributes:
        wrap_style (int): 决定字幕行如何换行。
        scaled_border_and_shadow (bool): 确定是否使用脚本分辨率(*True*)或视频分辨率(*False*)来缩放边框和阴影。
        play_res_x (int): 视频宽度。
        play_res_y (int): 视频高度。
        audio (str): 加载的音频的绝对路径。
        video (str): 加载的视频的绝对路径。
    """

    wrap_style: int
    scaled_border_and_shadow: bool
    play_res_x: int
    play_res_y: int
    audio: str
    video: str

    def __repr__(self):
        return pretty_print(self)


class Style:
    """Style对象包含一组应用于对话行的排版格式规则。

    可以在此获取样式的更多信息: https://aegi.vmoe.info/docs/3.2/ASS_Tags/ 。

    Attributes:
        fontname (str): 字体名
        fontsize (float): 字体大小（点数）
        color1 (str): 主要颜色
        alpha1 (str): 主要颜色的透明度
        color2 (str): 次要颜色（用于kalaoke效果）
        alpha2 (str): 次要颜色的透明度
        color3 (str): 边框颜色
        alpha3 (str): 边框颜色的透明度
        color4 (str): 阴影颜色
        alpha4 (str): 阴影颜色的透明度
        bold (bool): 字体是否加粗
        italic (bool): 字体是否为斜体
        underline (bool): 字体是否有下划线
        strikeout (bool): 字体是否有删除线
        scale_x (float): 文本水平缩放
        scale_y (float): 文本垂直缩放
        spacing (float): 字间距
        angle (float): 旋转角度
        border_style (bool): 是否是不透明背景
        outline (float): 边框厚度
        shadow (float): 阴影距离
        alignment (int): 对齐方式
        margin_l (int): 左边框距离
        margin_r (int): 右边框距离
        margin_v (int): 垂直边框距离
        encoding (int): 字符编码
    """

    fontname: str
    fontsize: float
    color1: str
    alpha1: str
    color2: str
    alpha2: str
    color3: str
    alpha3: str
    color4: str
    alpha4: str
    bold: bool
    italic: bool
    underline: bool
    strikeout: bool
    scale_x: float
    scale_y: float
    spacing: float
    angle: float
    border_style: bool
    outline: float
    shadow: float
    alignment: int
    margin_l: int
    margin_r: int
    margin_v: int
    encoding: int

    def __repr__(self):
        return pretty_print(self)


class Char:
    """Char对象包含了Ass中一行的单个字符的信息。

    一个char可以是karaoke标签 (k, ko, kf) 间的一些文本。

    Attributes:
        i (int): 字符的索引值
        word_i (int): 字符所在的单词的索引值 (e.g.: 在文本 ``Hello PyonFX users!`` 中，字母 "u" 的 word_i=2)。
        syl_i (int): 字符所在的音节的索引值 (e.g.: 在文本 ``{\\k0}Hel{\\k0}lo {\\k0}Pyon{\\k0}FX {\\k0}users!`` 中，字母 "F" 的 syl_i=3)。
        syl_char_i (int): 字符在所在音节中的索引值 (e.g.: 在文本 ``{\\k0}Hel{\\k0}lo {\\k0}Pyon{\\k0}FX {\\k0}users!`` 中，"users" 中的字母 "e" 的 syl_char_i=2)。
        start_time (int): 字符开始时间（毫秒）
        end_time (int): 字符结束时间（毫秒）
        duration (int): 字符持续时间（毫秒）
        styleref (obj): 此对象原始行的Style对象的引用。
        text (str): 字符文本。
        inline_fx (str): 字符内联特效 (在K值后使用 \\-特效名)。
        prespace (int): 文本之前的字符空白空间。
        postspace (int): 文本之后的字符空白空间。
        width (float): 字符文本宽度。
        height (float): 字符文本高度。
        x (float): 字符水平坐标（取决于对齐方式）。
        y (float): 字符垂直坐标（取决于对齐方式）。
        left (float): 字符左侧横坐标。
        center (float): 字符中心横坐标。
        right (float): 字符右侧横坐标。
        top (float): 字符顶部纵坐标。
        middle (float): 字符中心纵坐标。
        bottom (float): 字符底部纵坐标。
    """

    i: int
    word_i: int
    syl_i: int
    syl_char_i: int
    start_time: int
    end_time: int
    duration: int
    styleref: Style
    text: str
    inline_fx: str
    prespace: int
    postspace: int
    width: float
    height: float
    x: float
    y: float
    left: float
    center: float
    right: float
    top: float
    middle: float
    bottom: float

    def __repr__(self):
        return pretty_print(self)


class Syllable:
    """Syllable对象包含了Ass中一行的单个音节的信息。

    一个syl可以是karaoke标签 (k, ko, kf) 后的一些文本。
    (e.g.: 在 ``{\\k0}Hel{\\k0}lo {\\k0}Pyon{\\k0}FX {\\k0}users!`` 中，"Pyon" 和 "FX" 是两个不同的音节)，

    Attributes:
        i (int): 音节索引值。
        word_i (int): 音节单词索引值 (e.g.: 在行文本 ``{\\k0}Hel{\\k0}lo {\\k0}Pyon{\\k0}FX {\\k0}users!`` , 音节 "Pyon" 的 word_i=1)。
        start_time (int): 音节开始时间（毫秒）。
        end_time (int): 音节结束时间（毫秒）。
        duration (int): 音节持续时间（毫秒）。
        styleref (obj): 此对象原始行的Style对象的引用。
        text (str): 音节文本。
        tags (str): 除\\k之外音节前面的所有标签。
        inline_fx (str): 音节内联特效 (在K值后使用 \\-特效名)。
        prespace (int): 文本之前的音节空白空间。
        postspace (int): 文本之后的音节空白空间。
        width (float): 音节文本宽度。
        height (float): 音节文本高度。
        x (float): 音节水平坐标（取决于对齐方式）。
        y (float): 音节垂直坐标（取决于对齐方式）。
        left (float): 音节左侧横坐标。
        center (float): 音节中心横坐标。
        right (float): 音节右侧横坐标。
        top (float): 音节顶部纵坐标。
        middle (float): 音节中心纵坐标。
        bottom (float): 音节底部纵坐标。
    """

    i: int
    word_i: int
    start_time: int
    end_time: int
    duration: int
    styleref: Style
    text: str
    tags: str
    inline_fx: str
    prespace: int
    postspace: int
    width: float
    height: float
    x: float
    y: float
    left: float
    center: float
    right: float
    top: float
    middle: float
    bottom: float

    def __repr__(self):
        return pretty_print(self)


class Word:
    """Word对象包含了Ass中一行的单个单词的信息。

    一个word可以是一些文本，在其前后都有一些可选的空格
    (e.g.: 在字符串“What a beautiful world!”中，“beautiful” 和 “world”都是不同的单词)。

    Attributes:
        i (int): 单词索引值。
        start_time (int): 单词开始时间（与行开始时间相同）（毫秒）。
        end_time (int): 单词结束时间（与行结束时间相同）（毫秒）。
        duration (int): 单词持续时间（与行持续时间相同）（毫秒）。
        styleref (obj): 此对象原始行的Style对象的引用。
        text (str): 单词文本。
        prespace (int): 文本之前的字空白空间。
        postspace (int): 文本之后的字空白空间。
        width (float): 单词文本宽度。
        height (float): 单词文本高度。
        x (float): 单词水平坐标（取决于对齐方式）。
        y (float): 单词垂直坐标（取决于对齐方式）。
        left (float): 单词左横坐标。
        center (float): 单词中心横坐标。
        right (float): 单词右横坐标。
        top (float): 单词上部纵坐标。
        middle (float): 单词中心纵坐标。
        bottom (float): 单词底部纵坐标。
    """

    i: int
    start_time: int
    end_time: int
    duration: int
    styleref: Style
    text: str
    prespace: int
    postspace: int
    width: float
    height: float
    x: float
    y: float
    left: float
    center: float
    right: float
    top: float
    middle: float
    bottom: float

    def __repr__(self):
        return pretty_print(self)


class Line:
    """Line对象包括了Ass中每一行的信息。

    Note:
        (*) = 此项仅在 :class:`extended<Ass>` = True 时有效

    Attributes:
        i (int): 行索引值
        comment (bool): 是否为注释行。如果为 *True* ，这行将不会在屏幕上显示。
        layer (int): 行的层号。层号较高的行将显示在层号较低的行上方。
        start_time (int): 行开始时间（毫秒）。
        end_time (int): 行结束时间（毫秒）。
        duration (int): 行持续时间（毫秒） (*)。
        leadin (float): 行前时间（毫秒，首行为 1000.1） (*)。
        leadout (float): 行后时间（毫秒，首行为 1000.1） (*)。
        style (str): 此行使用的样式名称。
        styleref (obj): 此行的Style对象的引用 (*)。
        actor (str): 说话人。
        margin_l (int): 该行的左边距。
        margin_r (int): 该行的右边距。
        margin_v (int): 该行的垂直边距。
        effect (str): 特效。
        raw_text (str): 该行的原始文本。
        text (str): 该行的文本（无标签）。
        width (float): 行文本宽度 (*)。
        height (float): 行文本高度 (*)。
        ascent (float): 行上方距离 (*)。
        descent (float): 行下方距离 (*)。
        internal_leading (float): Line font internal lead (*).
        external_leading (float): Line font external lead (*).
        x (float): 行水平坐标（取决于对齐方式）(*)。
        y (float): 行垂直坐标（取决于对齐方式）(*)。
        left (float): 行左侧横坐标 (*)。
        center (float): 行中心横坐标 (*)。
        right (float): 行右侧横坐标 (*)。
        top (float): 行顶部纵坐标 (*)。
        middle (float): 行中心纵坐标 (*)。
        bottom (float): 行底部纵坐标 (*)。
        words (list): 包含此行内 :class:`Word` 对象的列表 (*)。
        syls (list): 包含此行内 :class:`Syllable` 对象的列表(如果有) (*)。
        chars (list): 包含此行内 :class:`Char` 对象的列表 (*)。
    """

    i: int
    comment: bool
    layer: int
    start_time: int
    end_time: int
    duration: int
    leadin: float
    leadout: float
    style: str
    styleref: Style
    actor: str
    margin_l: int
    margin_r: int
    margin_v: int
    effect: str
    raw_text: str
    text: str
    width: float
    height: float
    ascent: float
    descent: float
    internal_leading: float
    external_leading: float
    x: float
    y: float
    left: float
    center: float
    right: float
    top: float
    middle: float
    bottom: float
    words: List[Word]
    syls: List[Syllable]
    chars: List[Char]

    def __repr__(self):
        return pretty_print(self)

    def copy(self):
        """
        Returns:
            此对象 (line) 的深度副本
        """
        return copy.deepcopy(self)


class Ass:
    """包含关于 ASS 格式文件的所有信息，以及用于输入和输出的方法。

    | 通常，您将创建一个 Ass 对象，并将其用于输入和输出(参见 example _ section)。
    | PyonFX 会自动把输出文件中的所有信息设置为绝对路径，这样无论你
    把生成的文件放在哪里，它都会正确地加载视频和音频。

    Args:
        path_input (str): 输入文件路径 (可以为与 .py 文件的相对路径，也可以是绝对路径)。
        path_output (str): 输出文件路径 (可以为与 .py 文件的相对路径，也可以是绝对路径) (默认: "Output.ass")。
        keep_original (bool): 如果为True，输入文件的所有行将注释后放在生成的新行之前。
        extended (bool): 计算更多来自行的信息(通常用不到)。
        vertical_kanji (bool): 如果为True，对齐方式为 4, 5, 6 的行会被竖直放置。

    Attributes:
        path_input (str): 输入文件绝对路径。
        path_output (str): 输出文件绝对路径。
        meta (:class:`Meta`): 包含给出的 ASS 的信息。
        styles (list of :class:`Style`): 包含给出的 ASS 的所有样式。
        lines (list of :class:`Line`): 包含给出的 ASS 的所有行（事件）。

    .. _example:
    Example:
        ..  code-block:: python3

            io = Ass("in.ass")
            meta, styles, lines = io.get_data()
    """

    def __init__(
        self,
        path_input="",
        path_output="Output.ass",
        keep_original=True,
        extended=True,
        vertical_kanji=True,
    ):
        # Starting to take process time
        self.__saved = False
        self.__plines = 0
        self.__ptime = time.time()

        self.meta, self.styles, self.lines = Meta(), {}, []
        # Getting absolute sub file path
        dirname = os.path.dirname(os.path.abspath(sys.argv[0]))
        if not os.path.isabs(path_input):
            path_input = os.path.join(dirname, path_input)

        # Checking sub file validity (does it exists?)
        if not os.path.isfile(path_input):
            raise FileNotFoundError(
                "Invalid path for the Subtitle file: %s" % path_input
            )

        # Getting absolute output file path
        if path_output == "Output.ass":
            path_output = os.path.join(dirname, path_output)
        elif not os.path.isabs(path_output):
            path_output = os.path.join(dirname, path_output)

        self.path_input = path_input
        self.path_output = path_output
        self.__output = []

        section = ""
        li = 0
        for line in open(self.path_input, "r", encoding="utf-8-sig"):
            # Getting section
            section_pattern = re.compile(r"^\[([^\]]*)")
            if section_pattern.match(line):
                # Updating section
                section = section_pattern.match(line)[1]
                # Appending line to output
                self.__output.append(line)

            # Parsing Meta data
            elif section == "Script Info" or section == "Aegisub Project Garbage":
                # Internal function that tries to get the absolute path for media files in meta
                def get_media_abs_path(mediafile):
                    # If this is not a dummy video, let's try to get the absolute path for the video
                    if not mediafile.startswith("?dummy"):
                        tmp = mediafile
                        media_dir = os.path.dirname(self.path_input)

                        while mediafile.startswith("../"):
                            media_dir = os.path.dirname(media_dir)
                            mediafile = mediafile[3:]

                        mediafile = os.path.normpath(
                            "%s%s%s" % (media_dir, os.sep, mediafile)
                        )

                        if not os.path.isfile(mediafile):
                            mediafile = tmp

                    return mediafile

                # Switch
                if re.match(r"WrapStyle: *?(\d+)$", line):
                    self.meta.wrap_style = int(line[11:].strip())
                elif re.match(r"ScaledBorderAndShadow: *?(.+)$", line):
                    self.meta.scaled_border_and_shadow = line[23:].strip() == "yes"
                elif re.match(r"PlayResX: *?(\d+)$", line):
                    self.meta.play_res_x = int(line[10:].strip())
                elif re.match(r"PlayResY: *?(\d+)$", line):
                    self.meta.play_res_y = int(line[10:].strip())
                elif re.match(r"Audio File: *?(.*)$", line):
                    self.meta.audio = get_media_abs_path(line[11:].strip())
                    line = "Audio File: %s\n" % self.meta.audio
                elif re.match(r"Video File: *?(.*)$", line):
                    self.meta.video = get_media_abs_path(line[11:].strip())
                    line = "Video File: %s\n" % self.meta.video

                # Appending line to output
                self.__output.append(line)

            # Parsing Styles
            elif section == "V4+ Styles":
                # Appending line to output
                self.__output.append(line)
                style = re.match(r"Style: (.+?)$", line)

                if style:
                    # Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour,
                    # Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle,
                    # BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
                    style = [el for el in style[1].split(",")]
                    tmp = Style()

                    tmp.fontname = style[1]
                    tmp.fontsize = float(style[2])

                    r, g, b, a = Convert.coloralpha(style[3])
                    tmp.color1 = Convert.coloralpha(r, g, b)
                    tmp.alpha1 = Convert.coloralpha(a)

                    r, g, b, a = Convert.coloralpha(style[4])
                    tmp.color2 = Convert.coloralpha(r, g, b)
                    tmp.alpha2 = Convert.coloralpha(a)

                    r, g, b, a = Convert.coloralpha(style[5])
                    tmp.color3 = Convert.coloralpha(r, g, b)
                    tmp.alpha3 = Convert.coloralpha(a)

                    r, g, b, a = Convert.coloralpha(style[6])
                    tmp.color4 = Convert.coloralpha(r, g, b)
                    tmp.alpha4 = Convert.coloralpha(a)

                    tmp.bold = style[7] == "-1"
                    tmp.italic = style[8] == "-1"
                    tmp.underline = style[9] == "-1"
                    tmp.strikeout = style[10] == "-1"

                    tmp.scale_x = float(style[11])
                    tmp.scale_y = float(style[12])

                    tmp.spacing = float(style[13])
                    tmp.angle = float(style[14])

                    tmp.border_style = style[15] == "3"
                    tmp.outline = float(style[16])
                    tmp.shadow = float(style[17])

                    tmp.alignment = int(style[18])
                    tmp.margin_l = int(style[19])
                    tmp.margin_r = int(style[20])
                    tmp.margin_v = int(style[21])

                    tmp.encoding = int(style[22])

                    self.styles[style[0]] = tmp
            # Parsing Dialogues
            elif section == "Events":
                # Appending line to output (commented) if keep_original is True
                if keep_original:
                    self.__output.append(
                        re.sub(r"^(Dialogue|Comment):", "Comment:", line)
                    )

                # Analyzing line
                line = re.match(r"(Dialogue|Comment): (.+?)$", line)

                if line:
                    # Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
                    tmp = Line()

                    tmp.i = li
                    li += 1

                    tmp.comment = line[1] == "Comment"
                    line = [el for el in line[2].split(",")]

                    tmp.layer = int(line[0])

                    tmp.start_time = Convert.time(line[1])
                    tmp.end_time = Convert.time(line[2])

                    tmp.style = line[3]
                    tmp.actor = line[4]

                    tmp.margin_l = int(line[5])
                    tmp.margin_r = int(line[6])
                    tmp.margin_v = int(line[7])

                    tmp.effect = line[8]

                    tmp.raw_text = ",".join(line[9:])

                    self.lines.append(tmp)

        # Adding informations to lines and meta?
        if not extended:
            return None

        lines_by_styles = {}
        # Let the fun begin (Pyon!)
        for li, line in enumerate(self.lines):
            try:
                line.styleref = self.styles[line.style]
            except KeyError:
                line.styleref = None

            # Append dialog to styles (for leadin and leadout later)
            if line.style not in lines_by_styles:
                lines_by_styles[line.style] = []
            lines_by_styles[line.style].append(line)

            line.duration = line.end_time - line.start_time
            line.text = re.sub(r"\{.*?\}", "", line.raw_text)

            # Add dialog text sizes and positions (if possible)
            if line.styleref:
                # Creating a Font object and saving return values of font.get_metrics() for the future
                font = Font(line.styleref)
                font_metrics = font.get_metrics()

                line.width, line.height = font.get_text_extents(line.text)
                (
                    line.ascent,
                    line.descent,
                    line.internal_leading,
                    line.external_leading,
                ) = font_metrics

                if self.meta.play_res_x > 0 and self.meta.play_res_y > 0:
                    # Horizontal position
                    tmp_margin_l = (
                        line.margin_l if line.margin_l != 0 else line.styleref.margin_l
                    )
                    tmp_margin_r = (
                        line.margin_r if line.margin_r != 0 else line.styleref.margin_r
                    )

                    if (line.styleref.alignment - 1) % 3 == 0:
                        line.left = tmp_margin_l
                        line.center = line.left + line.width / 2
                        line.right = line.left + line.width
                        line.x = line.left
                    elif (line.styleref.alignment - 2) % 3 == 0:
                        line.left = (
                            self.meta.play_res_x / 2
                            - line.width / 2
                            + tmp_margin_l / 2
                            - tmp_margin_r / 2
                        )
                        line.center = line.left + line.width / 2
                        line.right = line.left + line.width
                        line.x = line.center
                    else:
                        line.left = self.meta.play_res_x - tmp_margin_r - line.width
                        line.center = line.left + line.width / 2
                        line.right = line.left + line.width
                        line.x = line.right

                    # Vertical position
                    if line.styleref.alignment > 6:
                        line.top = (
                            line.margin_v
                            if line.margin_v != 0
                            else line.styleref.margin_v
                        )
                        line.middle = line.top + line.height / 2
                        line.bottom = line.top + line.height
                        line.y = line.top
                    elif line.styleref.alignment > 3:
                        line.top = self.meta.play_res_y / 2 - line.height / 2
                        line.middle = line.top + line.height / 2
                        line.bottom = line.top + line.height
                        line.y = line.middle
                    else:
                        line.top = (
                            self.meta.play_res_y
                            - (
                                line.margin_v
                                if line.margin_v != 0
                                else line.styleref.margin_v
                            )
                            - line.height
                        )
                        line.middle = line.top + line.height / 2
                        line.bottom = line.top + line.height
                        line.y = line.bottom

                # Calculating space width and saving spacing
                space_width = font.get_text_extents(" ")[0]
                style_spacing = line.styleref.spacing

                # Adding words
                line.words = []

                wi = 0
                for prespace, word_text, postspace in re.findall(
                    r"(\s*)([^\s]+)(\s*)", line.text
                ):
                    word = Word()

                    word.i = wi
                    wi += 1

                    word.start_time = line.start_time
                    word.end_time = line.end_time
                    word.duration = line.duration

                    word.styleref = line.styleref
                    word.text = word_text

                    word.prespace = len(prespace)
                    word.postspace = len(postspace)

                    word.width, word.height = font.get_text_extents(word.text)
                    (
                        word.ascent,
                        word.descent,
                        word.internal_leading,
                        word.external_leading,
                    ) = font_metrics

                    line.words.append(word)

                # Calculate word positions with all words data already available
                if line.words and self.meta.play_res_x > 0 and self.meta.play_res_y > 0:
                    if line.styleref.alignment > 6 or line.styleref.alignment < 4:
                        cur_x = line.left
                        for word in line.words:
                            # Horizontal position
                            cur_x = cur_x + word.prespace * (
                                space_width + style_spacing
                            )

                            word.left = cur_x
                            word.center = word.left + word.width / 2
                            word.right = word.left + word.width

                            if (line.styleref.alignment - 1) % 3 == 0:
                                word.x = word.left
                            elif (line.styleref.alignment - 2) % 3 == 0:
                                word.x = word.center
                            else:
                                word.x = word.right

                            # Vertical position
                            word.top = line.top
                            word.middle = line.middle
                            word.bottom = line.bottom
                            word.y = line.y

                            # Updating cur_x
                            cur_x = (
                                cur_x
                                + word.width
                                + word.postspace * (space_width + style_spacing)
                                + style_spacing
                            )
                    else:
                        max_width, sum_height = 0, 0
                        for word in line.words:
                            max_width = max(max_width, word.width)
                            sum_height = sum_height + word.height

                        cur_y = x_fix = self.meta.play_res_y / 2 - sum_height / 2
                        for word in line.words:
                            # Horizontal position
                            x_fix = (max_width - word.width) / 2

                            if line.styleref.alignment == 4:
                                word.left = line.left + x_fix
                                word.center = word.left + word.width / 2
                                word.right = word.left + word.width
                                word.x = word.left
                            elif line.styleref.alignment == 5:
                                word.left = self.meta.play_res_x / 2 - word.width / 2
                                word.center = word.left + word.width / 2
                                word.right = word.left + word.width
                                word.x = word.center
                            else:
                                word.left = line.right - word.width - x_fix
                                word.center = word.left + word.width / 2
                                word.right = word.left + word.width
                                word.x = word.right

                            # Vertical position
                            word.top = cur_y
                            word.middle = word.top + word.height / 2
                            word.bottom = word.top + word.height
                            word.y = word.middle
                            cur_y = cur_y + word.height

                # Search for dialog's text chunks, to later create syllables
                # A text chunk is a text with one or more {tags} preceding it
                # Tags can be some text or empty string
                text_chunks = []
                tag_pattern = re.compile(r"(\{.*?\})+")
                tag = tag_pattern.search(line.raw_text)
                word_i = 0

                if not tag:
                    # No tags found
                    text_chunks.append({"tags": "", "text": line.raw_text})
                else:
                    # First chunk without tags?
                    if tag.start() != 0:
                        text_chunks.append(
                            {"tags": "", "text": line.raw_text[0 : tag.start()]}
                        )

                    # Searching for other tags
                    while True:
                        next_tag = tag_pattern.search(line.raw_text, tag.end())
                        tmp = {
                            # Note that we're removing possibles '}{' caused by consecutive tags
                            "tags": line.raw_text[
                                tag.start() + 1 : tag.end() - 1
                            ].replace("}{", ""),
                            "text": line.raw_text[
                                tag.end() : (next_tag.start() if next_tag else None)
                            ],
                            "word_i": word_i,
                        }
                        text_chunks.append(tmp)

                        # If there are some spaces after text, then we're at the end of the current word
                        if re.match(r"(.*?)(\s+)$", tmp["text"]):
                            word_i = word_i + 1

                        if not next_tag:
                            break
                        tag = next_tag

                # Adding syls
                si = 0
                last_time = 0
                inline_fx = ""
                syl_tags_pattern = re.compile(r"(.*?)\\[kK][of]?(\d+)(.*)")

                line.syls = []
                for tc in text_chunks:
                    # If we don't have at least one \k tag, everything is invalid
                    if not syl_tags_pattern.match(tc["tags"]):
                        line.syls.clear()
                        break

                    posttags = tc["tags"]
                    syls_in_text_chunk = []
                    while True:
                        # Are there \k in posttags?
                        tags_syl = syl_tags_pattern.match(posttags)

                        if not tags_syl:
                            # Append all the temporary syls, except last one
                            for syl in syls_in_text_chunk[:-1]:
                                curr_inline_fx = re.search(r"\\\-([^\\]+)", syl.tags)
                                if curr_inline_fx:
                                    inline_fx = curr_inline_fx[1]
                                syl.inline_fx = inline_fx

                                # Hidden syls are treated like empty syls
                                syl.prespace, syl.text, syl.postspace = 0, "", 0

                                syl.width, syl.height = font.get_text_extents("")
                                (
                                    syl.ascent,
                                    syl.descent,
                                    syl.internal_leading,
                                    syl.external_leading,
                                ) = font_metrics

                                line.syls.append(syl)

                            # Append last syl
                            syl = syls_in_text_chunk[-1]
                            syl.tags += posttags

                            curr_inline_fx = re.search(r"\\\-([^\\]+)", syl.tags)
                            if curr_inline_fx:
                                inline_fx = curr_inline_fx[1]
                            syl.inline_fx = inline_fx

                            if tc["text"].isspace():
                                syl.prespace, syl.text, syl.postspace = 0, tc["text"], 0
                            else:
                                syl.prespace, syl.text, syl.postspace = re.match(
                                    r"(\s*)(.*?)(\s*)$", tc["text"]
                                ).groups()
                                syl.prespace, syl.postspace = (
                                    len(syl.prespace),
                                    len(syl.postspace),
                                )

                            syl.width, syl.height = font.get_text_extents(syl.text)
                            (
                                syl.ascent,
                                syl.descent,
                                syl.internal_leading,
                                syl.external_leading,
                            ) = font_metrics

                            line.syls.append(syl)
                            break

                        pretags, kdur, posttags = tags_syl.groups()

                        # Create a Syllable object
                        syl = Syllable()

                        syl.start_time = last_time
                        syl.end_time = last_time + int(kdur) * 10
                        syl.duration = int(kdur) * 10

                        syl.styleref = line.styleref
                        syl.tags = pretags

                        syl.i = si
                        syl.word_i = tc["word_i"]

                        syls_in_text_chunk.append(syl)

                        # Update working variable
                        si += 1
                        last_time = syl.end_time

                # Calculate syllables positions with all syllables data already available
                if line.syls and self.meta.play_res_x > 0 and self.meta.play_res_y > 0:
                    if (
                        line.styleref.alignment > 6
                        or line.styleref.alignment < 4
                        or not vertical_kanji
                    ):
                        cur_x = line.left
                        for syl in line.syls:
                            cur_x = cur_x + syl.prespace * (space_width + style_spacing)
                            # Horizontal position
                            syl.left = cur_x
                            syl.center = syl.left + syl.width / 2
                            syl.right = syl.left + syl.width

                            if (line.styleref.alignment - 1) % 3 == 0:
                                syl.x = syl.left
                            elif (line.styleref.alignment - 2) % 3 == 0:
                                syl.x = syl.center
                            else:
                                syl.x = syl.right

                            cur_x = (
                                cur_x
                                + syl.width
                                + syl.postspace * (space_width + style_spacing)
                                + style_spacing
                            )

                            # Vertical position
                            syl.top = line.top
                            syl.middle = line.middle
                            syl.bottom = line.bottom
                            syl.y = line.y

                    else:  # Kanji vertical position
                        max_width, sum_height = 0, 0
                        for syl in line.syls:
                            max_width = max(max_width, syl.width)
                            sum_height = sum_height + syl.height

                        cur_y = self.meta.play_res_y / 2 - sum_height / 2

                        # Fixing line positions
                        line.top = cur_y
                        line.middle = self.meta.play_res_y / 2
                        line.bottom = line.top + sum_height
                        line.width = max_width
                        line.height = sum_height
                        if line.styleref.alignment == 4:
                            line.center = line.left + max_width / 2
                            line.right = line.left + max_width
                        elif line.styleref.alignment == 5:
                            line.left = line.center - max_width / 2
                            line.right = line.left + max_width
                        else:
                            line.left = line.right - max_width
                            line.center = line.left + max_width / 2

                        for syl in line.syls:
                            # Horizontal position
                            x_fix = (max_width - syl.width) / 2
                            if line.styleref.alignment == 4:
                                syl.left = line.left + x_fix
                                syl.center = syl.left + syl.width / 2
                                syl.right = syl.left + syl.width
                                syl.x = syl.left
                            elif line.styleref.alignment == 5:
                                syl.left = line.center - syl.width / 2
                                syl.center = syl.left + syl.width / 2
                                syl.right = syl.left + syl.width
                                syl.x = syl.center
                            else:
                                syl.left = line.right - syl.width - x_fix
                                syl.center = syl.left + syl.width / 2
                                syl.right = syl.left + syl.width
                                syl.x = syl.right

                            # Vertical position
                            syl.top = cur_y
                            syl.middle = syl.top + syl.height / 2
                            syl.bottom = syl.top + syl.height
                            syl.y = syl.middle
                            cur_y = cur_y + syl.height

                # Adding chars
                line.chars = []

                # If we have syls in line, we prefert to work with them to provide more informations
                if line.syls:
                    words_or_syls = line.syls
                else:
                    words_or_syls = line.words

                # Getting chars
                char_index = 0
                for el in words_or_syls:
                    el_text = "{}{}{}".format(
                        " " * el.prespace, el.text, " " * el.postspace
                    )
                    for ci, char_text in enumerate(list(el_text)):
                        char = Char()
                        char.i = ci

                        # If we're working with syls, we can add some indexes
                        char.i = char_index
                        char_index += 1
                        if line.syls:
                            char.word_i = el.word_i
                            char.syl_i = el.i
                            char.syl_char_i = ci
                        else:
                            char.word_i = el.i

                        # Adding last fields based on the existance of syls or not
                        char.start_time = el.start_time
                        char.end_time = el.end_time
                        char.duration = el.duration

                        char.styleref = line.styleref
                        char.text = char_text

                        char.width, char.height = font.get_text_extents(char.text)
                        (
                            char.ascent,
                            char.descent,
                            char.internal_leading,
                            char.external_leading,
                        ) = font_metrics

                        line.chars.append(char)

                # Calculate character positions with all characters data already available
                if line.chars and self.meta.play_res_x > 0 and self.meta.play_res_y > 0:
                    if line.styleref.alignment > 6 or line.styleref.alignment < 4:
                        cur_x = line.left
                        for char in line.chars:
                            # Horizontal position
                            char.left = cur_x
                            char.center = char.left + char.width / 2
                            char.right = char.left + char.width

                            if (line.styleref.alignment - 1) % 3 == 0:
                                char.x = char.left
                            elif (line.styleref.alignment - 2) % 3 == 0:
                                char.x = char.center
                            else:
                                char.x = char.right

                            cur_x = cur_x + char.width + style_spacing

                            # Vertical position
                            char.top = line.top
                            char.middle = line.middle
                            char.bottom = line.bottom
                            char.y = line.y
                    else:
                        max_width, sum_height = 0, 0
                        for char in line.chars:
                            max_width = max(max_width, char.width)
                            sum_height = sum_height + char.height

                        cur_y = x_fix = self.meta.play_res_y / 2 - sum_height / 2
                        for char in line.chars:
                            # Horizontal position
                            x_fix = (max_width - char.width) / 2
                            if line.styleref.alignment == 4:
                                char.left = line.left + x_fix
                                char.center = char.left + char.width / 2
                                char.right = char.left + char.width
                                char.x = char.left
                            elif line.styleref.alignment == 5:
                                char.left = self.meta.play_res_x / 2 - char.width / 2
                                char.center = char.left + char.width / 2
                                char.right = char.left + char.width
                                char.x = char.center
                            else:
                                char.left = line.right - char.width - x_fix
                                char.center = char.left + char.width / 2
                                char.right = char.left + char.width
                                char.x = char.right

                            # Vertical position
                            char.top = cur_y
                            char.middle = char.top + char.height / 2
                            char.bottom = char.top + char.height
                            char.y = char.middle
                            cur_y = cur_y + char.height

        # Add durations between dialogs
        for style in lines_by_styles:
            lines_by_styles[style].sort(key=lambda x: x.start_time)
            for li, line in enumerate(lines_by_styles[style]):
                line.leadin = (
                    1000.1
                    if li == 0
                    else line.start_time - lines_by_styles[style][li - 1].end_time
                )
                line.leadout = (
                    1000.1
                    if li == len(lines_by_styles[style]) - 1
                    else lines_by_styles[style][li + 1].start_time - line.end_time
                )

    def get_data(self):
        """实用功能，可轻松检索元样式和行。

        Returns:
            :attr:`meta`, :attr:`styles` and :attr:`lines`
        """
        return self.meta, self.styles, self.lines

    def write_line(self, line):
        """在输出列表上追加一行 (私有)，调用 save() 时会被写入输出文件。

        记得准备好一行后调用它，并且在调用 :func:`save` 
        之前什么也不会做，因为没有写入任何内容。

        Parameters:
            line (:class:`Line`): 一个line对象。如果无效，则抛出 TypeError 异常。
        """
        if isinstance(line, Line):
            self.__output.append(
                "\n%s: %d,%s,%s,%s,%s,%04d,%04d,%04d,%s,%s"
                % (
                    "Comment" if line.comment else "Dialogue",
                    line.layer,
                    Convert.time(max(0, int(line.start_time))),
                    Convert.time(max(0, int(line.end_time))),
                    line.style,
                    line.actor,
                    line.margin_l,
                    line.margin_r,
                    line.margin_v,
                    line.effect,
                    line.text,
                )
            )
            self.__plines += 1
        else:
            raise TypeError("Expected Line object, got %s." % type(line))

    def save(self, quiet=False):
        """将私有输出列表中的内容写入文件。

        Parameters:
            quiet (bool): 如果为True，将不会输出任何信息。
        """

        # Writing to file
        with open(self.path_output, "w", encoding="utf-8-sig") as f:
            f.writelines(self.__output)
        self.__saved = True

        if not quiet:
            print(
                "Produced lines: %d\nProcess duration (in seconds): %.3f"
                % (self.__plines, time.time() - self.__ptime)
            )

    def open_aegisub(self):
        """用Aegisub打开输出文件 (在 self.path_output 中指定)。

        如果您未安装MPV或想要详细查看输出文件，这很有用。

        Returns:
            成功返回0，如果输出文件无法打开则返回-1。
        """

        # Check if it was saved
        if not self.__saved:
            print(
                "[WARNING] You've tried to open the output with Aegisub before having saved. Check your code."
            )
            return -1

        if sys.platform == "win32":
            os.startfile(self.path_output)
        else:
            try:
                subprocess.call(["aegisub", os.path.abspath(self.path_output)])
            except FileNotFoundError:
                print("[WARNING] Aegisub not found.")
                return -1

        return 0

    def open_mpv(self, video_path="", video_start="", full_screen=False):
        """用MPV播放器以软字幕打开输出文件 (在 self.path_output 中指定)。
        要使用这个函数，需要安装MPV。并且如果您使用的是Windows的话，需要添加MPV到PATH中 (转到 https://pyonfx.readthedocs.io/en/latest/quick%20start.html#installation-extra-step).

        这是以一种舒适的方式播放输出文件的最快的方法之一。

        Parameters:
            video_path (string): 要播放的视频的绝对路径。如果未指定，将自动使用**meta.video**。
            video_start (string): 视频开始时间 (更多信息，请转到: https://mpv.io/manual/master/#options-start)。如果未指定，将自动使用0.
            full_screen (bool): 如果为True，将全屏播放输出文件。如果未指定，将自动使用False。
        """

        # Check if it was saved
        if not self.__saved:
            print(
                "[ERROR] You've tried to open the output with MPV before having saved. Check your code."
            )
            return -1

        # Check if mpv is usable
        if self.meta.video.startswith("?dummy") and not video_path:
            print(
                "[WARNING] Cannot use MPV (if you have it in your PATH) for file preview, since your .ass contains a dummy video.\n"
                "You can specify a new video source using video_path parameter, check the documentation of the function."
            )
            return -1

        # Setting up the command to execute
        cmd = ["mpv"]

        if not video_path:
            cmd.append(self.meta.video)
        else:
            cmd.append(video_path)
        if video_start:
            cmd.append("--start=" + video_start)
        if full_screen:
            cmd.append("--fs")

        cmd.append("--sub-file=" + self.path_output)

        try:
            subprocess.call(cmd)
        except FileNotFoundError:
            print(
                "[WARNING] MPV not found in your environment variables.\n"
                "Please refer to the documentation's \"Quick Start\" section if you don't know how to solve it."
            )
            return -1

        return 0
