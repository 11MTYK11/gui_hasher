#Copyright (c) 2021 mattuu

#This software is released under the MIT License.
#http://opensource.org/licenses/mit-license.php

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import hashlib
import os,sys
import subprocess
import threading
import mimetypes
import datetime
import math
from tkinter import scrolledtext
from tkinter.constants import E, NONE
import tkinterdnd2
import comtypes.client

class hashframe(tk.Frame):
    def __init__(self,master):
        super().__init__(master)
        self.master = master
        self.create_widgets() 
    def create_widgets(self):
        self.frame = tk.Frame(self)
        self.frame.pack(fill=tk.BOTH,expand=True)
        self.filepath = None
        self.filesize = None
        self.lastacsess = None
        self.maketime = None
        self.filetype = None
        self.fontsize = 15
        self.pathframe = tk.Frame(self.frame)
        self.pathframe.pack(fill=tk.BOTH)
        self.reloadbtn = tk.Button(self.pathframe,command=self.reload,text="更新")
        self.reloadbtn.pack(side=tk.LEFT,fill=tk.Y)
        self.pathlbl = tk.Label(self.pathframe,font=("",self.fontsize),text="ファイルパス")
        self.pathlbl.pack(fill=tk.BOTH,side=tk.LEFT,expand=True)
        self.dataframe = tk.LabelFrame(self.frame,text="ファイル情報")
        self.dataframe.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
        self.sizelbl = tk.Label(self.dataframe,font=("",self.fontsize),text="ファイルサイズ")
        self.sizelbl.pack(fill=tk.BOTH,expand=True)
        self.lastlbl = tk.Label(self.dataframe,font=("",self.fontsize),text="最終アクセス")
        self.lastlbl.pack(fill=tk.BOTH,expand=True)
        self.makelbl = tk.Label(self.dataframe,font=("",self.fontsize),text="作成日時")
        self.makelbl.pack(fill=tk.BOTH,expand=True)
        self.typelbl = tk.Label(self.dataframe,font=("",self.fontsize),text="ファイルタイプ")
        self.typelbl.pack(fill=tk.BOTH,expand=True)
        self.hashdata = None
        self.hashframe = tk.LabelFrame(self.frame,text="ファイルハッシュ")
        self.hashframe.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
        self.selectframe = tk.Frame(self.hashframe)
        self.selectframe.pack(fill=tk.X)
        self.hashplatorm = tk.StringVar(self)
        self.hashplatorm.set("now_platform")
        rdo1 = tk.Radiobutton(self.selectframe,value="now_platform",variable=self.hashplatorm,text='現在のプラットフォームで使えるもの',command=self.platformchange)
        rdo1.pack(fill=tk.BOTH,side=tk.LEFT)
        rdo2 = tk.Radiobutton(self.selectframe,value="all_platform",variable=self.hashplatorm,text='すべてのプラットフォームで使えるもの',command=self.platformchange)
        rdo2.pack(fill=tk.BOTH,side=tk.LEFT)
        selectframe = tk.Frame(self.hashframe)
        selectframe.pack(fill=tk.X)
        lbl = tk.Label(selectframe,text="ハッシュタイプ→")
        lbl.pack(fill=tk.X,side=tk.LEFT)
        self.hashselect = ttk.Combobox(selectframe,font=("",15),state="readonly")
        self.hashselect.pack(fill=tk.X,side=tk.LEFT,expand=True)
        self.hashselect.bind('<<ComboboxSelected>>',self.changehash)
        self.hashlbl = scrolledtext.ScrolledText(self.hashframe,font=("",15),height=5,width=1)
        self.hashlbl.pack(fill=tk.BOTH,expand=True)
        sizeframe = tk.LabelFrame(self.hashframe,text="サイズ")
        sizeframe.pack(fill=tk.BOTH,expand=True)
        self.hashsize = tk.Label(sizeframe)
        self.hashsize.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
        self.blocksize = tk.Label(sizeframe)
        self.blocksize.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
        self.progress = ttk.Progressbar(self.master)
        self.progress.pack(fill=tk.X,side=tk.BOTTOM)
        self.platformchange()
    def platformchange(self):
        if self.hashplatorm.get() == "now_platform":
            self.hashselect.config(values=list(hashlib.algorithms_available))
        else:
            self.hashselect.config(values=list(hashlib.algorithms_guaranteed))
    def changehash(self,event):
        self.hashselect.config(state="disabled")
        algo = self.hashselect.get()
        hash = hashlib.new(algo)
        Length = hashlib.new(algo).block_size * 0x800
        self.progress.config(maximum=math.ceil(os.path.getsize(self.filepath)/Length))
        self.progress.update()
        path = self.filepath
        hashcount = 0
        with open(path,'rb') as file:
            BinaryData = file.read(Length)
            while BinaryData:
                hashcount += 1
                self.progress.config(value=hashcount)
                self.progress.update()
                hash.update(BinaryData)
                BinaryData = file.read(Length)
        self.hashlbl.delete(0.0,tk.END)
        try:
            self.hashlbl.insert(tk.END,hash.hexdigest())
        except:
            self.hashlbl.insert(tk.END,"エラーが発生しました")
        self.blocksize.config(text=f"ダイジェストサイズ{hash.digest_size}Byte")
        self.hashsize.config(text=f"ハッシュサイズ{hash.block_size}Byte")
        self.hashselect.config(state="readonly")
    def readfile(self,filepath):
        self.filepath = f"{str(os.path.abspath(filepath))}"
        self.filesize = f"ファイルサイズ{str(os.path.getsize(filepath))}Byte"
        filestate = os.stat(filepath)
        self.lastacsess = f"最終アクセス:{str(datetime.datetime.fromtimestamp(filestate.st_atime))}"
        self.maketime = f"作成時間:{str(datetime.datetime.fromtimestamp(filestate.st_ctime))}"
        self.filetype = f"ファイルタイプ:{str(mimetypes.guess_type(filepath)[0])}"
        self.pathlbl.config(text=self.filepath)
        self.sizelbl.config(text=self.filesize)
        self.lastlbl.config(text=self.lastacsess)
        self.makelbl.config(text=self.maketime)
        self.typelbl.config(text=self.filetype)
        self.master.update()
    def reload(self):
        self.reloadbtn.config(state="disabled")
        try:
            self.readfile(self.filepath)
        except:
            pass
        try:
            self.changehash(None)
        except:
            pass
        self.hashselect.config(state="readonly")
        self.reloadbtn.config(state="normal")
def tabopen():
    global tabsjson
    selected_indices = selectlist.curselection()
    selecteddata = [selectlist.get(i) for i in selected_indices]
    for file in selecteddata:
        if file in tabsjson.keys():
            tabwin.select(tabsjson[file])
        else:
            frame = tk.Frame(tabwin)
            tabwin.add(frame,text=os.path.basename(file))
            hashframes.append(hashframe(frame))
            hashframes[-1].pack(fill=tk.BOTH,expand=True)
            hashframes[-1].readfile(file)
            tabsjson.setdefault(file,frame)
def removelist():
    global fileslist
    closetab()
    selected_indices = selectlist.curselection()
    selecteddata = [selectlist.get(i) for i in selected_indices]
    for file in selecteddata:
        try:
            fileslist.remove(file)
        except:
            pass
    selectvalue.set(fileslist)

def closetab():
    global tabsjson
    selected_indices = selectlist.curselection()
    selecteddata = [selectlist.get(i) for i in selected_indices]
    for file in selecteddata:
        try:
            tabwin.forget(tabsjson[file])
        except:
            pass
        try:
            tabsjson.pop(file)
        except:
            pass
root = tkinterdnd2.Tk()
root.attributes("-topmost",True)
root.title("ハッシュ計算君")
root.resizable(0,0)
tabwin = ttk.Notebook(root)
tabwin.pack(fill=tk.BOTH,expand=True)

mainframe = tk.Frame(tabwin)
tabwin.add(mainframe,text="メニュー")
tabwin.select(tabwin.tabs()[tabwin.index("end")-1])
hashframes = []
fileslist = []
tabsjson = {}
try:
    files = sys.argv[1:]
    if len(files) == 1:
        file = files[0]
        if os.path.exists(file):
            if os.path.isdir(file):
                pass
            else:
                fileslist.append(file)
                frame = tk.Frame(tabwin)
                tabwin.add(frame,text=os.path.basename(file))
                tabwin.select(tabwin.tabs()[tabwin.index("end")-1])
                hashframes.append(hashframe(frame))
                hashframes[-1].pack(fill=tk.BOTH,expand=True)
                hashframes[-1].readfile(file)
                tabsjson.setdefault(file,frame)
    else:
        for file in files:
            if os.path.exists(file):
                if os.path.isdir(file):
                    pass
                else:
                    fileslist.append(file)
except:
    import traceback
    traceback.print_exc()
def closealltab():
    global tabsjson
    keylist = list(tabsjson.keys())
    for key in keylist:
        try:
            tabwin.forget(tabsjson[key])
        except:
            pass
        try:
            tabsjson.pop(key)
        except:
            pass
def filter():
    global fileslist
    oklist = []
    if extension_entry.get() == "":
        selectvalue.set(fileslist)
    else:
        filterlist = extension_entry.get().split(",")
        for file in fileslist:
            if os.path.exists(file):
                if os.path.isfile(file):
                    filename, ext = os.path.splitext(file)
                    if ext.replace(".","") in filterlist:
                        oklist.append(file)
                        selectvalue.set(oklist)
                    else:
                        selectvalue.set(oklist)
            root.update()

selectvalue = tk.StringVar(value=fileslist)
selectlist = tk.Listbox(mainframe,listvariable=selectvalue,selectmode="extended")
selectlist.pack(fill=tk.BOTH,expand=True)
selectlist.filenames = {}
selectlist.nextcoords = [50, 20]
selectlist.dragging = False

def openfile():
    files = filedialog.askopenfilenames(filetypes=[("すべてのファイル","*")])
    if files == "":
        pass
    else:
        for file in files:
            if os.path.exists(file):
                if os.path.isdir(file):
                    pass
                else:
                    fileslist.append(file)
        selectvalue.set(fileslist)

def drag_end(event):
    selectlist.dragging = False
def drop_enter(event):
    event.widget.focus_force()
    return event.action
def drop_position(event):
    return event.action
def drop_leave(event):
    return event.action
def drop(event):
    if selectlist.dragging:
        return tkinterdnd2.REFUSE_DROP
    if event.data:
        items = selectlist.tk.splitlist(event.data)
        for item in items:
            if os.path.exists(item):
                if os.path.isdir(item):
                    pass
                else:
                    fileslist.append(item)
            selectvalue.set(fileslist)
    return event.action
def drag_init(event):
    try:
        data = ()
        sel = selectlist.select_item()
        if sel:
            data = (selectlist.filenames[sel],)
            selectlist.dragging = True
            return ((tkinterdnd2.ASK, tkinterdnd2.COPY), (tkinterdnd2.DND_FILES, tkinterdnd2.DND_TEXT), data)
        else:
            return 'break'
    except:
        return 'break'
selectlist.drop_target_register(tkinterdnd2.DND_FILES)
selectlist.dnd_bind('<<DropEnter>>', drop_enter)
selectlist.dnd_bind('<<DropPosition>>', drop_position)
selectlist.dnd_bind('<<DropLeave>>', drop_leave)
selectlist.dnd_bind('<<Drop>>', drop)
selectlist.drag_source_register(1, tkinterdnd2.DND_FILES)
selectlist.dnd_bind('<<DragInitCmd>>', drag_init)
selectlist.dnd_bind('<<DragEndCmd>>', drag_end)
btnframe = tk.Frame(mainframe)
btnframe.pack(fill=tk.BOTH,expand=True)
openbtn = tk.Button(btnframe,text="タブで開く",command=tabopen)
openbtn.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
closebtn = tk.Button(btnframe,text="タブを閉じる",command=closetab)
closebtn.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
closebtn = tk.Button(btnframe,text="すべてのタブを閉じる",command=closealltab)
closebtn.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
closebtn = tk.Button(btnframe,text="リストから削除する",command=removelist)
closebtn.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
closebtn = tk.Button(btnframe,text="参照する",command=openfile)
closebtn.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)
lbl = tk.Label(mainframe,text="拡張子フィルター↓(拡張子ごとに,で区切ってください、'.'は入れないでください)",font=("",15))
lbl.pack(fill=tk.BOTH,expand=True)
lbl = tk.Label(mainframe,text="例：exe,png,jpg,dll  フィルタリングしない場合は空白にしてください",font=("",20))
lbl.pack(fill=tk.BOTH,expand=True)
extension_entry = tk.Entry(mainframe,font=("",15))
extension_entry.pack(fill=tk.BOTH,expand=True)
btn = tk.Button(mainframe,text="フィルタリングする",font=("",15),command=filter)
btn.pack(fill=tk.BOTH,expand=True)
root.mainloop()