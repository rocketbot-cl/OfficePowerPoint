# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"
    
    pip install <package> -t .

"""
import os
import sys

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + \
    'OfficePowerPoint' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)
docto = os.path.join(cur_path.replace("libs", "bin"), "docto.exe")

from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx import Presentation


def makeTmpDir(name):
    try:
        os.mkdir("tmp")
        os.mkdir("tmp" + os.sep + name)
    except:
        try:
            os.mkdir("tmp" + os.sep + name)
        except:
            pass

    return os.sep.join(["tmp", name])


module = GetParams("module")
global prs
global slide
global slide_layout

if module == "new":
    try:
        prs = Presentation()
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e


if module == "open":
    try:
        path = GetParams("path")
        prs = None
        prs = Presentation(path)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

elif module == "get_slides_types":
    res = GetParams("res")
    list_types = []
    try:
        for type_slide in prs.slide_layouts:
            list_types.append(type_slide.name)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e
    SetVar(res, list_types)


elif module == "new_slide":
    name = GetParams("type")    
    try:
        list_types = []
        for type_slide in prs.slide_layouts:
            list_types.append(type_slide.name)
        try:
            name = int(name)
            name = list_types[name]
        except ValueError: 
            pass
        print("name",name)
        slide_layout = prs.slide_layouts.get_by_name(name)
        print("slide_layout",slide_layout)
        slide = prs.slides.add_slide(slide_layout)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

elif module == "save":
    path = GetParams("path")
    try:
        if path:
            prs.save(path)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e
elif module == "write":
    text = GetParams("text").replace("\\n", "\n")
    type_ = GetParams("type")
    align = GetParams("align")
    size = GetParams("size")
    bold = GetParams("bold")
    ital = GetParams("italic")
    under = GetParams("underline")

    try:
        if align == "left":
            align = PP_ALIGN.LEFT
        elif align == "center" or None:
            align = PP_ALIGN.CENTER
        elif align == "right":
            align = PP_ALIGN.RIGHT
        elif align == "justify":
            align = PP_ALIGN.JUSTIFY

        if size:
            size = Pt(int(size))
        if bold:
            bold = eval(bold)
        if ital:
            ital = eval(ital)
        if under:
            under = eval(under)

        placeholders = slide.shapes
        for placeholder in placeholders:
            if placeholder.name.lower().startswith(type_[:-1]):
                print(placeholder.name)
                if type_[-1] == "2":
                    type_ = type_[:-1]
                    continue
                else:
                    text_frame = placeholder.text_frame
                    p = text_frame.paragraphs[0]
                    p.text = text
                    if size:
                        p.font.size = size
                    p.font.bold = bold
                    p.font.italic = ital
                    p.font.underline = under
                    p.alignment = align
                    type_ = "uuuu"

    except Exception as e:
        PrintException()
        raise e

elif module == "close":
    prs = None

elif module == "add_pic":

    img_path = GetParams("img_path")
    index = GetParams("slide")
    pos = GetParams("position")
    height = GetParams("height")
    try:
        placeholders = slide.shapes.placeholders
        list_types = []
        for type_slide in prs.slide_layouts:
            list_types.append(type_slide.name)
        if slide_layout.name == list_types[8]:
            for placeholder in placeholders:
                if "picture" in placeholder.name.lower():
                    placeholder.insert_picture(img_path)
        else:
            top, left = eval(pos)
            top, left = Pt(float(top)), Pt(float(left))
            if height:
                height = Pt(float(height))
                pic = slide.shapes.add_picture(
                    img_path, left=left, top=top, height=height)
            else:
                pic = slide.shapes.add_picture(img_path, left=left, top=top)
    except Exception as e:
        PrintException()
        raise e

elif module == "addTextbox":
    text = GetParams("text")
    position = GetParams("position")
    size = GetParams("size")

    try:
        position = position.split(",")
        size = size.split(",")
        pos = position + size
        for i in range(len(pos)):
            pos[i] = Pt(int(pos[i]))

        txt_box = slide.shapes.add_textbox(pos[0], pos[1], pos[2], pos[3])
        txt_box.text_frame.text = text

    except Exception as e:
        PrintException()
        raise e

if module == "editText":
    try:
        search_by = GetParams("searchBy")
        index = GetParams("slide")
        shape_index = GetParams("shape")
        text = GetParams("text")

        if search_by == "id":
            slide = prs.slides.get(index)
        if search_by == "position":
            slide = prs.slides[int(index)]
        shape_index = int(shape_index)
        slide.shapes[shape_index].text_frame.text = text

    except Exception as e:
        PrintException()
        raise e


if module == "listSlides":
    res = GetParams("res")
    slide_index = int(GetParams("slide_index"))
    try:
        text_runs = []
        slide = prs.slides[slide_index]
        for index, shape in enumerate(slide.shapes):
            if shape.has_text_frame:
                value = (index, shape.text)
                text_runs.append(value)
        SetVar(res, text_runs)
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
