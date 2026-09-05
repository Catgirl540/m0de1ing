# -*- coding: utf-8 -*-
"""COM open-test + slide export via pywin32."""
import sys, os
import win32com.client

BASE = r"C:\Users\yan\Desktop\数模\ppt_studio\work_faxi"

def open_test(files):
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        for f in files:
            full = os.path.join(BASE, f)
            try:
                p = app.Presentations.Open(full, True, False, False)
                if p is None:
                    print(f, "NULL")
                else:
                    print(f, "OK", p.Slides.Count)
                    p.Close()
            except Exception as e:
                print(f, "ERR:", str(e)[:120])
    finally:
        app.Quit()

def export(pptx, prefix, width=1920, height=1080):
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        p = app.Presentations.Open(pptx, True, False, False)
        for i in range(1, p.Slides.Count + 1):
            out = os.path.join(BASE, f"{prefix}_{i}.png")
            p.Slides(i).Export(out, "PNG", width, height)
        n = p.Slides.Count
        p.Close()
        print("EXPORTED", n)
    finally:
        app.Quit()

if __name__ == "__main__":
    if sys.argv[1] == "test":
        open_test(sys.argv[2:])
    elif sys.argv[1] == "export":
        export(sys.argv[2], sys.argv[3])
