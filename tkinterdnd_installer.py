import os
import sys
import shutil
import subprocess

def resourcePath(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(filename)

try:
    shutil.copytree(resourcePath("必要なファイル\\tkdnd2.8"),sys.exec_prefix + "\\tcl\\tkdnd2.8")
except:
    import traceback
    traceback.print_exc()

os.chdir(resourcePath("必要なファイル\\tkinterdnd2-master"))
try:
    subprocess.run(["python","setup.py","install"],shell=True)
except:
    import traceback
    traceback.print_exc()
