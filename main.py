#!/usr/bin/env python3
import sys
import os
from tkinter import filedialog,Tk,Label,Button,Checkbutton,BooleanVar
from tkinter import ttk
import tkinter as tk
from PIL import Image
from PIL import ImageTk

from image_processing import count_droplets
from plot import plot_bins
from save_excel import save_to_excel

panel_list=[]
img_original_list=[]
img_contours_list=[]
image_paths=[]
data=[]
notebook=None
btn_add_image=None
btn_save_excel=None

def add_image(image_path):
    global panel_list,showContours,img_original_list,img_contours_list,notebook,image_paths,data,btn_show_histogram

    image_paths.append(image_path)
    btn_show_histogram['state']='normal'
    btn_save_excel['state']='normal'

    droplet_diameters,img_original,img_contours=count_droplets(image_path)

    data.append(droplet_diameters)
        
    img_original = Image.fromarray(img_original)
    img_contours = Image.fromarray(img_contours)

    img_original = ImageTk.PhotoImage(img_original)
    img_contours = ImageTk.PhotoImage(img_contours)

    img_original_list.append(img_original)
    img_contours_list.append(img_contours)

    if showContours.get():
        img=img_contours
    else:
        img=img_original

    panel = Label(image=img)
    panel.image = img
    panel_list.append(panel)

    notebook.add(panel,text=image_path.split('/')[-1].split('\\')[-1])

def add_image_button():

    image_path = filedialog.askopenfilename()

    if len(image_path)>0:
        add_image(image_path)

def show_histogram():
    global image_paths,data
    image_relative_paths=[x.split('/')[-1].split('\\')[-1] for x in image_paths]
    histogram_title=",".join(image_relative_paths)
    plot_bins(data,histogram_title,image_relative_paths)

def toggleShowContours():
    for panel,img_contours,img_original in zip(panel_list,img_contours_list,img_original_list):
        if showContours.get():
            img=img_contours
        else:
            img=img_original

        panel.configure(image=img)
        panel.image = img

def save_excel():
    global image_paths,data
    image_relative_paths=[x.split('/')[-1].split('\\')[-1] for x in image_paths]
    histogram_title=",".join(image_relative_paths)
    save_to_excel(data,image_relative_paths,"excel/"+histogram_title[:249]+".xlsx")

# source: https://stackoverflow.com/a/39459376
class CustomNotebook(ttk.Notebook):
    """A ttk Notebook with close buttons on each tab"""

    __initialized = False

    def __init__(self, *args, **kwargs):
        if not self.__initialized:
            self.__initialize_custom_style()
            self.__inititialized = True

        kwargs["style"] = "CustomNotebook"
        ttk.Notebook.__init__(self, *args, **kwargs)

        self._active = None

        self.bind("<ButtonPress-1>", self.on_close_press, True)
        self.bind("<ButtonRelease-1>", self.on_close_release)

    def on_close_press(self, event):
        """Called when the button is pressed over the close button"""

        element = self.identify(event.x, event.y)

        if "close" in element:
            index = self.index("@%d,%d" % (event.x, event.y))
            self.state(['pressed'])
            self._active = index
            return "break"

    def on_close_release(self, event):
        """Called when the button is released"""

        global panel_list,img_original_list,img_contours_list,image_paths,data

        if not self.instate(['pressed']):
            return

        element =  self.identify(event.x, event.y)
        if "close" not in element:
            # user moved the mouse off of the close button
            return

        index = self.index("@%d,%d" % (event.x, event.y))

        if self._active == index:
            self.forget(index)
            panel_list.pop(index)
            img_original_list.pop(index)
            img_contours_list.pop(index)
            image_paths.pop(index)
            data.pop(index)
            if len(image_paths)==0:
                btn_show_histogram['state']='disabled'  
                btn_save_excel['state']='disabled'
            self.event_generate("<<NotebookTabClosed>>")

        self.state(["!pressed"])
        self._active = None


    def __initialize_custom_style(self):
        style = ttk.Style()
        self.images = (
            tk.PhotoImage("img_close", data='''
                R0lGODlhCAAIAMIBAAAAADs7O4+Pj9nZ2Ts7Ozs7Ozs7Ozs7OyH+EUNyZWF0ZWQg
                d2l0aCBHSU1QACH5BAEKAAQALAAAAAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU
                5kEJADs=
                '''),
            tk.PhotoImage("img_closeactive", data='''
                R0lGODlhCAAIAMIEAAAAAP/SAP/bNNnZ2cbGxsbGxsbGxsbGxiH5BAEKAAQALAAA
                AAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU5kEJADs=
                '''),
            tk.PhotoImage("img_closepressed", data='''
                R0lGODlhCAAIAMIEAAAAAOUqKv9mZtnZ2Ts7Ozs7Ozs7Ozs7OyH+EUNyZWF0ZWQg
                d2l0aCBHSU1QACH5BAEKAAQALAAAAAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU
                5kEJADs=
            ''')
        )

        style.element_create("close", "image", "img_close",
                            ("active", "pressed", "!disabled", "img_closepressed"),
                            ("active", "!disabled", "img_closeactive"), border=8, sticky='')
        style.layout("CustomNotebook", [("CustomNotebook.client", {"sticky": "nswe"})])
        style.layout("CustomNotebook.Tab", [
            ("CustomNotebook.tab", {
                "sticky": "nswe",
                "children": [
                    ("CustomNotebook.padding", {
                        "side": "top",
                        "sticky": "nswe",
                        "children": [
                            ("CustomNotebook.focus", {
                                "side": "top",
                                "sticky": "nswe",
                                "children": [
                                    ("CustomNotebook.label", {"side": "left", "sticky": ''}),
                                    ("CustomNotebook.close", {"side": "left", "sticky": ''}),
                                ]
                        })
                    ]
                })
            ]
        })
    ])

if __name__ == "__main__":

    if not os.path.exists('excel'):
        os.makedirs('excel')

    if len(sys.argv)>1:
        debug=True
        for image_path in sys.argv[1:]:
            droplet_diameters,img,img_contours=count_droplets(image_path,debug)
            print(image_path)
            print(f'Min diameter: {min(droplet_diameters)}')
            print(f'Max diameter: {max(droplet_diameters)}')
            image_paths.append(image_path.split('/')[-1].split('\\')[-1])
            data.append(droplet_diameters)
        histogram_title=",".join(image_paths)
        save_to_excel(data,image_paths,"excel/"+histogram_title[:249]+".xlsx")
        plot_bins(data,histogram_title,image_paths)
    else:

        root = Tk()
        root.title('Droplets counter')

        frame = tk.Frame(root,width=50)
        frame.pack(side="left")
        
        btn_add_image = Button(frame, text="Add image", command=add_image_button)
        btn_add_image.pack( expand="no", padx="10", pady="10")

        showContours = BooleanVar(value=True) 
        check=Checkbutton(frame,text='Show contours',variable=showContours,command=toggleShowContours)
        check.pack()

        btn_show_histogram = Button(frame, text="Show histogram", command=show_histogram)
        btn_show_histogram['state']='disabled'
        btn_show_histogram.pack( expand="no", padx="10", pady="10")

        btn_save_excel = Button(frame, text="Save data in excel", command=save_excel)
        btn_save_excel['state']='disabled'
        btn_save_excel.pack( expand="no", padx="10", pady="10")

        notebook = CustomNotebook(width=1280,height=898)
        notebook.pack(side="right", fill="both", expand=True)
        

        root.mainloop()