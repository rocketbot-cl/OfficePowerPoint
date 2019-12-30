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
cur_path = base_path + 'modules' + os.sep + 'OfficePowerPoint' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

docto = os.path.join(cur_path.replace("libs", "bin"), "docto.exe")

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
    prs = Presentation()

if module == "open":
    path = GetParams("path")

    prs = Presentation(path)

elif module == "new_slide":
    name = GetParams("type")

    slide_layout = prs.slide_layouts.get_by_name(name)
    slide = prs.slides.add_slide(slide_layout)

elif module == "save":

    path = GetParams("path")

    if path:
        prs.save(path)

elif module == "write":
    text = GetParams("text").replace("\\n", "\n")
    type_ = GetParams("type")
    align = GetParams("align")
    size = GetParams("size")
    bold = GetParams("bold")
    ital = GetParams("italic")
    under = GetParams("underline")



    if align == "left":
        align = PP_ALIGN.LEFT
    elif align == "center" or None:
        align = PP_ALIGN.CENTER
    elif align == "right":
        align = PP_ALIGN.RIGHT
    elif align == "justify":
        align = PP_ALIGN.JUSTIFY

    try:

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
    pos = GetParams("position")
    height = GetParams("height")

    placeholders = slide.shapes.placeholders

    print(slide_layout.name)
    try:
        if slide_layout.name == "Picture with Caption":
            for placeholder in placeholders:
                if "picture" in placeholder.name.lower():
                    placeholder.insert_picture(img_path)
        else:
            top, left = eval(pos)
            top, left = Pt(float(top)), Pt(float(left))
            if height:
                height = Pt(float(height))
                pic = slide.shapes.add_picture(img_path, left=left, top=top, height=height)
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





