# Author: Leo Traussnigg @ Technikgruppe

# import enum
from logging import error
import sys
import numpy as np
# from numpy.core import numeric
import matplotlib
import matplotlib.pyplot as plt
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# from matplotlib.figure import Figure
import matplotlib.animation as animation
import tkinter as tk
from tkinter import *
from tkinter import StringVar, filedialog, Label, Button, messagebox, Spinbox
import xlrd
#import /tg.ico

matplotlib.use('TkAgg')

# Animation exportieren

def saveAnimation(ani):
    print("Wird gespeichert...")
    savePath = filedialog.askdirectory(master=root, initialdir = "/",title = "Animation Speichern",)
    try:
        ani.save(savePath + "/heatmap_animation.mp4")
        print("Animation wurde als heatmap_animation.mp4 gespeichert!")
    except:
        print("Fehler beim Speichern aufgetreten: Möglicherweise ist FFMPEG nicht installiert..")
        messagebox.showerror("Fehler beim Speichern aufgetreten", "Fehler beim Speichern aufgetreten: FFMPEG nicht installiert / wurde nicht gefunden")

# Darstellung

def visualization(sensorData, rowStart, rowCount, rowPerFrame, fps):
    # Beschriftung
    row = rowStart
    zeilen = range(1,4+1)
    spalten= range(1,2+1)

    ax.set_xticks(np.arange(len(zeilen)))
    ax.set_yticks(np.arange(len(spalten)))
    ax.set_xticklabels(zeilen)
    ax.set_yticklabels(spalten)
    ax.set_title("Wärmestrom")

    # Animation
    ims = []

    # for i in range(sensorData.nrows-40):
    for i in range(rowCount):
        row += rowPerFrame
        t1 = sensorData.cell_value(row,1)
        t2 = sensorData.cell_value(row,2)
        t3 = sensorData.cell_value(row,3)
        t4 = sensorData.cell_value(row,4)
        t5 = sensorData.cell_value(row,5)
        t6 = sensorData.cell_value(row,6)
        t7 = sensorData.cell_value(row,7)
        t8 = sensorData.cell_value(row,8)
        sensorArray = np.array([[t1,t2,t3,t4],[t5,t6,t7,t8]])
        print(i, row)
        print(sensorArray)
# spline16, spline36
# magma, inferno, YlOrBr, YlOrRd, Wistia
        im = ax.imshow(sensorArray.astype(np.float64), interpolation="spline36", cmap='gnuplot2' ,animated=True)
        if i == 0:
            im = ax.imshow(sensorArray, interpolation="spline36", cmap='gnuplot2')  
            cbar = plt.colorbar(im, fraction=0.046, pad=0.14, label="ΔT in °C") 
        
        # for i in range(len(zeilen)):
        #     for j in range(len(spalten)):
        #         text = ax.text(j, i, sensorArray[i, j],
        #             ha="center", va="center", color="w")

        ims.append([im])
        """ 
        a = fig.add_subplot(111)
        a.imshow(sensorArray.astype(np.float64), interpolation="spline16", cmap='YlOrBr' ,animated=True)
        canvas = FigureCanvasTkAgg(fig, root)
        canvas.draw() 
        """
        # Bilder einzeln exportieren:
        # fig.savefig("Zeile_" + str(row) + ".png", dpi=300)  # results in 160x120 px image
    print("Erstelle Animation...")
    ani = animation.ArtistAnimation(fig, ims, interval=(1000/fps), blit=True,
                                    repeat_delay=2000)
    plt.title("Wärmestrom Visualisierung")
    plt.show()
    if messagebox.askyesno("Speichern", "Animation als mp4 speichern?"): 
        try:
            saveAnimation(ani)
        except:
            if error == InterruptedError:
                messagebox.showwarning("FFMPEG möglicherweise nicht installiert")
        sys.exit()
    else: sys.exit()

fig = plt.Figure(figsize=(10,6), dpi=100)

fig, ax = plt.subplots()

# Pfad einlesen

def readPath():
    try:
        sheetPath = filedialog.askopenfilename(initialdir = "C:/Users/leoch/Desktop/TG",title = "Excel Datei auswählen",filetypes = ((".xls","*.xls"),("all files","*.*")))
        wb = xlrd.open_workbook(sheetPath)
        sensorData = wb.sheet_by_index(0)
    except: 
        if error == UnboundLocalError:
            messagebox.showerror("Fehler", "Bitte eine XLS Datei auswählen!")

    return sensorData

sensorData = readPath()

# UI

print("Starting...")
root = Tk()
root.wm_title("Visualisierung Wärmestrom")
#root.iconbitmap(root, default="tg.ico")
root.geometry('350x300')
root.grid_anchor("center")

fontSize = 20

lb = Label(root, text="Wärmestrom Visualisierung", font=fontSize, padx=20, pady=20) 
#lb.grid(column=1, row=2)
lb.pack()
#bt = Button(root, text="Pfad auswählen", command=readPath, font=fontSize)
#bt.grid(column=1, row=3)
#bt.pack()

root.title("Datenauswahl")
textStart = Label(root, text="Von Zeile: ")
#textStart.grid(column=3, row=3)
textStart.pack()
tvOffset = StringVar(root)
tvOffset.set=(18)
spinStart = Spinbox(root, from_=18, to=sensorData.nrows, width=7, textvariable=tvOffset)
#spinStart.grid(column=1,row=0)
spinStart.pack()
tvCount = StringVar(root)
tvCount.set(sensorData.nrows-20)
textCount = Label(root, text="Anzahl der angezeigten Zeilen: ")
# textCount.grid(column=0, row=1)
textCount.pack()
spinCount = tk.Spinbox(root, from_=1, to=sensorData.nrows, width=7, textvariable=tvCount)
# spinCount.grid(column=1,row=1)
spinCount.pack()
tvPerRow = StringVar(root)
tvPerRow.set("Zeilen pro Frame")
textPerRow = Label(root, text="Anzahl der Zeilen pro generiertem Frame: (WIP!)", fg="red")
# textPerRow.grid(column=0, row=2)
textPerRow.pack()
spinPerRow = Spinbox(root, from_=1, to=sensorData.nrows, width=7, textvariable=tvPerRow)
# spinPerRow.grid(column=1,row=2)
spinPerRow.pack()
tvFps = StringVar(root)
tvFps.set("Frames pro Sekunde")
textFps = Label(root, text="Frames pro Sekunde: ")
# textFps.grid(column=0, row=3)
textFps.pack()
spinFps = Spinbox(root, from_=1, to=sensorData.nrows, width=7, textvariable=tvFps)
# spinFps.grid(column=1,row=3)
spinFps.pack()



def confirmSelection():
    rowStart = int(spinStart.get())
    rowCount = int(spinCount.get())
    rowPerFrame = int(spinPerRow.get())
    fps = int(spinFps.get())
    root.withdraw()
    visualization(sensorData, rowStart, rowCount, rowPerFrame, fps)

confirm = Button(root, text="Bestätigen", command=confirmSelection, font=fontSize, padx=20, pady=10, border=5)
#confirm.grid(column=1, row=4)
confirm.pack()
#saveButton = Button(root, text="Animation Speichern", command=saveAnimation)
#saveButton.grid(column=1,row=4)
root.mainloop()