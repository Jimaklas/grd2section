# -*- coding: utf-8 -*-
import win32com.client
from pywintypes import com_error
from input import GRD_FILE_NAME

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument  # Document object


def get_num(s):
    return float("".join(char for char in s if char.isdigit() or char in ["-", "+", "."]))

f = open(GRD_FILE_NAME, "r")
line = f.readline()
while 1:
    try:
        line = f.readline()
        section = line.strip().split()[0]

        origin = doc.Utility.GetPoint(Prompt="Select origin of section %s (0,0 to skip this section):" % (section))
        if origin[0] == origin[1] == 0.0:
            while line[0] != "*":
                line = f.readline()
            continue
        originHeight, point_clicked = doc.Utility.GetEntity(None, None, Prompt="Select text that includes origin height:")
        originHeight = get_num(originHeight.TextString)

        command = "pl "
        line = f.readline()
        while line[0] != "*":
            offset, h = line.strip().split()
            offset = float(offset)
            h = float(h)

            x = origin[0] + offset
            y = origin[1] + (h - originHeight)

            command = command + "%s,%s " % (x, y)
            line = f.readline()

        print "Drawing section %s" % (section)
        doc.SendCommand(command + " ")

    except ValueError:  # raised when trying to read past EOF (why not IOError? - need to think on it)
        print "Program ended on ValueError"
        break

    except com_error:
        print "Program ended on com_error"
        break
