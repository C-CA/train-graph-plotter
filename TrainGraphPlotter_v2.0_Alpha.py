#! /csoft/epd-7.3.2/bin/python

"""
Created on Thu Oct  1 16:26:34 2015

@author: CHill16
This script was created by Catherine Hill (Network Rail Capability and Capacity Analysis team).

This script plots a graph of any compatible Excel timetable, if it comtains times in the format hh:mm:ss, and arrival/departure labels.

Parts labelled 'TGP2' are changes made during the Train Graph Plotter v2.0 upgrade in November 2020.

'TGP core' or 'core' refers to the original pre-November 2020 Train Graph Plotter v1.4 codebase.

Problems? Suggestions? If the date is before July 27,2021, send them to thariq.fahry@networkrail.co.uk.
                       If the date is after  July 27,2021, send them to alec.howe@networkrail.co.uk.
"""
#Ordered Set module
import collections

class OrderedSet(collections.MutableSet):

    def __init__(self, iterable=None):
        self.end = end = [] 
        end += [None, end, end]         # sentinel node for doubly linked list
        self.map = {}                   # key --> [key, prev, next]
        if iterable is not None:
            self |= iterable

    def __len__(self):
        return len(self.map)

    def __contains__(self, key):
        return key in self.map

    def add(self, key):
        if key not in self.map:
            end = self.end
            curr = end[1]
            curr[2] = end[1] = self.map[key] = [key, curr, end]

    def discard(self, key):
        if key in self.map:        
            key, prev, next = self.map.pop(key)
            prev[2] = next
            next[1] = prev

    def __iter__(self):
        end = self.end
        curr = end[2]
        while curr is not end:
            yield curr[0]
            curr = curr[2]

    def __reversed__(self):
        end = self.end
        curr = end[1]
        while curr is not end:
            yield curr[0]
            curr = curr[1]

    def pop(self, last=True):
        if not self:
            raise KeyError('set is empty')
        key = self.end[1][0] if last else self.end[2][0]
        self.discard(key)
        return key

    def __repr__(self):
        if not self:
            return '%s()' % (self.__class__.__name__,)
        return '%s(%r)' % (self.__class__.__name__, list(self))

    def __eq__(self, other):
        if isinstance(other, OrderedSet):
            return len(self) == len(other) and list(self) == list(other)
        return set(self) == set(other)

            
if __name__ == '__main__':
    s = OrderedSet('abracadaba')
    t = OrderedSet('simsalabim')
   # print(s | t)
    #print(s & t)
    #print(s - t)
'''
The following copyright notice relates to the ordered set module above:

Copyright (c) 2009  Raymond Hettinger

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

      
'''

#%%
#Import required packages.
import tkinter as tk
from tkinter import Label, Entry, Checkbutton, Button, messagebox, Frame, LabelFrame, DISABLED
import tkinter.ttk as ttk
import re
import xlwings as xw
import xlrd
import time
import numpy as np
import os
import datetime as dt
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.axis as ax
import matplotlib.lines as mlines
from matplotlib.dates import DateFormatter, MinuteLocator, HourLocator
from matplotlib.ticker import AutoMinorLocator
import ctypes
import sys
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import tkinter.filedialog
import base64
import pdb
#import pypdfocr.pypdfocr_gs as pdfImg
#import glob
#%%
np.set_printoptions(threshold=sys.maxsize)
#ctypes.windll.shcore.SetProcessDpiAwareness(1) #TGP2 allow DPI scaling for laptops


if getattr(sys, 'frozen', False):
  # Override dll search path.
  ctypes.windll.kernel32.SetDllDirectoryW('C:/Users/Tfarhy/Anaconda/Library/bin/')
  # Init code to load external dll
  ctypes.CDLL('mkl_avx2.dll')
  ctypes.CDLL('mkl_def.dll')
  ctypes.CDLL('mkl_vml_avx2.dll')
  ctypes.CDLL('mkl_vml_def.dll')

  # Restore dll search path.
  ctypes.windll.kernel32.SetDllDirectoryW(sys._MEIPASS)

if getattr(sys, 'frozen', False):
# we are running in a |PyInstaller| bundle
    basedir = sys._MEIPASS
else:
# we are running in a normal Python environment
    basedir = os.path.dirname(__file__)
#Set whether input box is used.
#useinputbox = True

#Variables: these are only used if useinputbox = False.
#Define main variables.
#directory = "H:\Capacity Plus\Stoke + Handsacre to Colwich\Phase 1" #Directory where input file is.
#infilename = 'Stoke_ScenarioB_graph.csv' #Input file with train times.
#outfilename = 'test2' #Filename for graph without file extenstion.
#pdf = True #Set to "True" to provide pdf output.
#png = False #Set to "True" to provide png output.
#gridlines = True
#min_time = dt.datetime.strptime("01:00:00","%H:%M:%S") #Minimum time to be shown on graph.
#max_time = dt.datetime.strptime("02:00:00","%H:%M:%S") #Maximum time to be shown on graph.
#min_dist = 63.4 #Minimum distance to be shown on graph.
#max_dist = -14.9 #Maximum distance to be shown on graph.

#Define variables to customise graph.
#l = 0.14 #Left position of graph on page (adjust these so that the location labels and legend fit. Must be between 0 and 1.)
#b = 0.08 #Bottom position of graph on page.
#t = 0.91 #Top position of graph on page.
#r = 0.93 #Right position of graph on page.
#legend_position = 1.09 #y position of the legend on the graph.
#ncols = 2 #Number of columns in the legend.
#size_x = 8 #Width of page in inches.
#size_y = 12.5 #Height of page in inches.

#icon = base64.encodestring(open("TGP3.ico", "rb").read())
icon = """
"""
icondata= base64.b64decode(icon)
tempfile= "icon.ico"
iconfile= open(tempfile,"wb")
iconfile.write(icondata)
iconfile.close()


################################debug function
def e(variable,bookname,sheetname): #"variable explorer"
    wb = xw.Book(bookname)
    sheet = wb.sheets[sheetname]
    sheet.range('A1:CC500').clear()
    sheet.range('A1:CC500').api.Font.Size = 10.0
    sheet.range('A1').value = variable
###############################################               

#readexcel(): directly reads an Excel spreadsheet. Since the xlwings library interacts with the Excel app's COM interface, and not the .xlsx file, 
#the spreadsheet that is being read needs to be open. This function will first open the sheet in Excel if it is not already.

##Since Excel internally stores times as floating-point values, e.g. 0.87364726482643 (even when displayed as hh:mm:ss on the .xlsx sheet),
##this function has to both read and convert those values to the hh:mm:ss that the TGP (Train Graph Plotter core code) takes.
         
#arguments:
#bookname = name of excel file to read from
#sheetname = name of sheet within file
#cell_range = range of cells to be read (e.g. A1:B55), 
#only_times = when True, ony Times in the format hh:mm:ss are read 
#as_string = when True, anything that's not a time is read as a String. Usually used in conjunction with the previous argument being False.

#return values: (3 seperate numpy arrays):
#cells = [cell contents],
#indices = [cell addresses in the format 'r,c'], used later to point to an exact cell when generating errors
#exceptions = [exceptions generated when reading each cell (only used when debugging)]    

def readexcel(bookname,sheetname,cell_range,only_times,as_string):                  
    wb = xw.Book(bookname)
    sheet = wb.sheets[sheetname]
    
    xcells = sheet.range(cell_range)                             #xcells just stores the range
    cells = xcells.options(np.array, ndim=2, dtype = object, empty = None).value #create data matrix as numpy array. specify ndim to be 2 to not trip up our nested iterator
                             
    indices = np.empty(xcells.shape, dtype = "<U20")             #preallocate index matrix with max length 20
    exceptions = np.empty(xcells.shape, dtype="<U20")            #preallocate empty string matrix with max length 20

    offset_r = sheet.range(cell_range).row                      #if cell_range does not start at A1, we add an offset to make sure Indices contains the right addresses
    offset_c = sheet.range(cell_range).column
                          
    for r, row in enumerate(cells):                             #iterate through rows (enumerate returns index, value for each row)
        for c, cell in enumerate(row):                          #iterate through cells in the current row
            indices[r][c] = '{},{}'.format(r+offset_r,c+offset_c) #populate address array
            if cell is None:                                    #if blank cell
                cells[r][c] = ''                                #set to empty string because TGP can't handle Nones
            else:
                try:                                            #try to convert float (from excel) to TGP-readable time string
                    value = xlrd.xldate_as_tuple(cell, 0)#note: last parameter = 'datemode', which doesn't matter in this case where we converting relative, not absolute, times
                    value = '{}:{}:{}'.format(str(value[3]).zfill(2),str(value[4]).zfill(2),str(value[5]).zfill(2))#convert tuple to hh:mm:ss format
                    cells[r][c] = value                         #place converted value in array
                except Exception as e:                          #if not a time (when xldate_as_tuple conversion fails)
                    exceptions[r][c] = (str(type(e).__name__)+':'+str(e)) #record exception
                    if only_times:                              #when reading time matrix, 'clean' it up by deleting non-times, to be TGP-compatible
                        cells[r][c] = ''                        #delete whatever non-time was in that cell                        
                    else:                                       #when we want to just read and not clean
                        if as_string:                           #when we want to read each non-time as a string (which we do, usually), particularly if any of said non-times consist of purely numbers
                            cells[r][c] = str(cells[r][c])
                        else:                                   #(usually not used): when we neither want to delete non-times nor convert to string
                            pass                                #do nothing (leave cell as is). This else:pass is not strictly necessary, but is left in for clarity.

    return cells, indices, exceptions

#stitch(locations,trains,cells,arrival/departure labels,list of labels): format and stitch together seperately-read arrays into a format TGP can parse.

def stitch(locations,trains,cells,arrdep,label_list):                                             
    #Create a colour row, auto-assigning colours to trains. Note: trains with identical names are assigned the same colour.
    for i, train in enumerate(trains[0]):
        if train == '':
            trains[0][i] = '<no label>'
    
    lightcolours = ['whitesmoke','white','snow','mistyrose','seashell','peachpuff','linen',
'antiquewhite','blanchedalmond','papayawhip','moccasin','wheat','oldlace','floralwhite',
'cornsilk','lemonchiffon','palegoldenrod','ivory','beige','lightyellow','lightgoldenrodyellow',
'honeydew','mintcream','azure','lightcyan','aliceblue','ghostwhite','lavender','thistle',
'lavenderblush','pink']
    darkcolours = [colourname for colourname in matplotlib.colors.cnames.keys() if colourname not in lightcolours]
    
    colour_dict = dict(zip(trains[0],darkcolours))
    colour_row = np.empty(trains.shape,dtype="<U35")
    for i, train in enumerate(trains[0]):
        colour_row[0][i] = colour_dict[train]
    
    #add the blank cells in the top-left corner
    colour_row = np.insert(colour_row,[0,0],'',axis=1)  
    trains = np.insert(trains,[0,0],'',axis=1)
    
    #Join Colour and Name rows vertically.
    traincolours = np.concatenate((trains,colour_row),axis=0)

    #Delete any rows not labelled Arrival, Departure or Pass.
    deletions = [] #
    for i in range (len(locations)):
        if arrdep[i][0] not in label_list or arrdep[i][0]=='':
            deletions = np.append(deletions,i)  
    locations = np.delete(locations,deletions,0) #last param is axis
    cells = np.delete(cells,deletions,0)
    
    #Auto-generate Distances column.
    distances = np.empty(locations.shape,dtype=int)
    j=0
    for i,dist in enumerate(distances):
        if locations[i] != '' and locations[i]!=locations[i-1]: #note, edge case: this line will misbehave (allocate 0 as first distance) if first and last location have the same names (a highly unlikely scenario) since on the first loop it compares location[0] to location[-1] (last element of location)
            j = j+1
        distances[i] = j
    
    max_dist = distances[-1][0]
    
    #Paste the other arrays into place, to mimic a TGP-compatible input array.
    output = np.concatenate((locations,distances),axis=1)
    output = np.concatenate((output,cells),axis=1)
    output = np.concatenate((traincolours,output),axis=0)
    
    #Delete trains with no times.
    deletions = []
    for i,tname in enumerate(output[0,:]): #note that, for simplicity, this also enumerates (and checks blank-ness of) the locations and distances columns (and doesn't delete either of them because they are, presumably, non-empty)
        if np.all(output[2:,i]==''):
            deletions = np.append(deletions,i)
    output = np.delete(output,deletions,1)
    
    #Create dummy. 
    #Note that below, the preallocation type is dtype="<U30" (string) - this is because passing a dtype = object to the core code's later functions seems to break them
    output_intermediate = np.empty((2*output.shape[0]+1,output.shape[1]),dtype="<U30") #create a blank array with (twice the height of our output + 1 blank row) to place the dummy in
    output_intermediate[:] = ''                                                        #numpy preallocates an Object array with Nones, and we want empty cells to have '' instead
    output_intermediate[0:output.shape[0],0:output.shape[1]] = output                  #'paste' our Output array into the top half of this new one that has twice the height
    output_intermediate[output.shape[0]+1:,0:3] = output[:,0:3]                        #'paste' the dummy train (which is just a copy of the first non-empty Up train) into the bottom half of the new array, along with the same locations and distances
    output = output_intermediate
    
    return output, max_dist

def exceladdressof(cell):
    return sheet.range(tuple([int(i) for i in re.findall('([0-9]+),([0-9]+)',cell)[0]])).get_address(False,False)

#Set up input box.
class GUI:
    def __init__(self, master):
        self.master = master
        master.title("Train Graph Plotter v2.0 (alpha build NOV18b)")
        master.geometry('690x530')
        master.resizable(False,False)
        master.iconbitmap(tempfile)
        os.remove(tempfile)
        #Read in saved variables if they exist.
        home = os.path.expanduser("~")
        home1 = 'H:\Train Graph Plotter\saved_variables_v2.0Alpha.txt'
        home2 = home + '\Train Graph Plotter\saved_variables_v2.0Alpha.txt'
        if os.path.isfile(home1) == True:
            path = home1
        elif os.path.isfile(home2) == True:
            path = home2
        else:
            path = None
        if path != None:
            variable_list = [line.rstrip() for line in open(path)]
            if len(variable_list) != 29:
                path = None
            if path != None:
                directory_save = variable_list[0]
                infilename_save = variable_list[1]
                outfilename_save = variable_list[2]
                pdf_save = variable_list[3]
                png_save = variable_list[4]
                gridlines_save = variable_list[5]
                min_time_save = variable_list[6][11:19]
                max_time_save = variable_list[7][11:19]
                min_dist_save = variable_list[8]
                max_dist_save = variable_list[9]
                label_freq_save = variable_list[10]
                grid_freq_save = variable_list[11]
                axes_save = variable_list[12]
                l_save = variable_list[13]
                b_save = variable_list[14]
                t_save = variable_list[15]
                r_save = variable_list[16]
                legend_position_save = variable_list[17]
                ncols_save = variable_list[18]
                size_x_save = variable_list[19]
                size_y_save= variable_list[20]
                preview_save = variable_list[21]
                
                ######TGP2
                location_column_save = variable_list[22]
                arrdep_labels_save = variable_list[23]
                arrdep_column_save = variable_list[24]
                train_row_save = variable_list[25]
                time_cell_start_save = variable_list[26]
                time_cell_end_save = variable_list[27]
                sheetname_save = variable_list[28]
                
                
        
        #Set up frames.
        self.fileframe=LabelFrame(master,text='File',width=350,height=100,bd=2,relief='groove')
        self.fileframe.grid(row=0,column=0,padx=(10,10),pady=(5,0),sticky='N')
        self.fileframe.grid_propagate(False)
        
        self.sheetframe=LabelFrame(master,text='Sheet',width=350,height=280,bd=2,relief='groove')
        self.sheetframe.grid(row=1,column=0,padx=(10,10),pady=(0,0),sticky='N')        
        self.sheetframe.grid_propagate(False)
        
        self.outputframe=LabelFrame(master,text='Output',width=350,height=100,bd=2,relief='groove')
        self.outputframe.grid(row=2,column=0,sticky='N',padx=(10,10))   
        self.outputframe.grid_propagate(False)
        
        self.optionframe=LabelFrame(master,text='Options',width=300,height=500,bd=2,relief='groove')
        self.optionframe.grid(row=0,column=1,rowspan=3,padx=(10,10),pady=(5,0))   
        self.optionframe.grid_propagate(False)
        


        #####################################sub-frames
        self.layoutframe=LabelFrame(self.optionframe,text='Graph position on page',width=250,height=85,bd=2,relief='groove')
        self.layoutframe.grid(row=29,column=0,columnspan=5,padx = 10,pady=(7,10),sticky='w')
        self.layoutframe.grid_propagate(False)
        
        self.legendframe=LabelFrame(self.optionframe,text='Legend',width=200,height=85,bd=2,relief='groove')
        self.legendframe.grid(row=30,column=0,columnspan=5,padx = 10,pady=(0,10),sticky='w')
        self.legendframe.grid_propagate(False)
        
        self.gridlineframe=LabelFrame(self.optionframe,text='Gridlines',width=200,height=85,bd=2,relief='groove')
        self.gridlineframe.grid(row=31,column=0,columnspan=5,padx = 10,pady=(0,10),sticky='w')
        self.gridlineframe.grid_propagate(False)
        
        self.labelsframe=LabelFrame(self.optionframe,text='Labels',width=200,height=90,bd=2,relief='groove')
        self.labelsframe.grid(row=32,column=0,columnspan=5,padx = 10,sticky='w')
        self.labelsframe.grid_propagate(False)
        
        self.sheetoptionframe=LabelFrame(self.sheetframe,text='Data locations',width=300,height=215,bd=2,relief='groove')
        self.sheetoptionframe.grid(row=1,column=0,columnspan=5,padx = 10,pady=(10,0),sticky='w')
        self.sheetoptionframe.grid_propagate(False)
        
        #Set up entry labels.
        self.l1=Label(master, text= "Main variables", font="Helvetica 10 bold")#.grid(row=1)
        self.l2=Label(self.fileframe, text= "Directory").grid(row=0,column=0,sticky = 'w',padx = (15,0),pady = (10,0))
        self.l3=Label(self.fileframe, text= "Input file").grid(row=1,column=0,sticky = 'w',padx = (15,0),pady = (10,0))
        self.l4=Label(self.outputframe, text= "Output file name").grid(row=0,column=0,sticky = 'w',padx = (15,0),pady = (10,0))
        self.l5=Label(master, text= "File format:", font="Helvetica 8 bold")#.grid(row=4)
        
        self.warninglabel=Label(master, text= "Warning: Alpha build. Copy any files before using!                                         Problems? E-mail thariq.fahry@networkrail.co.uk", font="Helvetica 8 bold").grid(row=3,sticky = 'NW',padx = (8,0),pady=(0,0),columnspan=3,rowspan=2)
        
        #Define checkbutton variables.
        self.check1var = tk.BooleanVar()
        self.check2var = tk.BooleanVar()
        self.check3var = tk.BooleanVar()
        self.check4var = tk.BooleanVar()
        self.check5var = tk.BooleanVar()

        #Check pdf and uncheck png and gridlines by default, unless there is a saved value.
        self.check1 = Checkbutton(self.outputframe, text= "pdf", variable=self.check1var)
        if "pdf_save" in locals():
            if pdf_save == 'True':
                self.check1.select()
        else:
            self.check1.select()
        self.check1.grid(row=1,column=0,sticky='w',padx=(110,0))
        
        self.check2=Checkbutton(self.outputframe, text= "png", variable=self.check2var)
        if "png_save" in locals():
            if png_save == 'True':
                self.check2.select()
        self.check2.grid(row=1, column=0,sticky='w',padx=(157,0))

        self.check3=Checkbutton(self.gridlineframe, text= "Show", variable=self.check3var)
        if "gridlines_save" in locals():
            if gridlines_save == 'True':
                self.check3.select()
        self.check3.grid(row=0, sticky = 'w',padx = (10,5),pady=(3,0))

        #Continue setting up entry labels.
        self.l6=Label(self.optionframe, text= "From time").grid(row=0,column=0,sticky = 'e',padx = 5,pady = (10,0))
        self.l7=Label(self.optionframe, text= "to time").grid(row=0,column=1,sticky = 'e',padx = (50,0),pady = (10,0),columnspan = 1)
        
        ######################################## removed in TGP2
        self.l8=Label(self.optionframe, text= "Minimum distance")#.grid(row=9) #
        self.l9=Label(self.optionframe, text= "Maximum distance")#.grid(row=10)
        self.l10=Label(self.optionframe, text= "Variables for customising graph", font="Helvetica 10 bold")#.grid(row=12)
        self.l22=Label(self.optionframe, text = "Labels", font="Helvetica 8 bold")#.grid(row=13)
        self.l25=Label(self.optionframe, text= "Axes", font="Helvetica 8 bold")#.grid(row=16)
        self.l11=Label(self.optionframe, text= "Graph position on page", font="Helvetica 8 bold")#.grid(row=18)
        self.l16=Label(self.optionframe, text= "Legend", font="Helvetica 8 bold")#.grid(row=23)
        self.l19=Label(self.optionframe, text= "Page size (inches)", font="Helvetica 8 bold")#.grid(row=26)
        #########################################################################
        
        self.l23=Label(self.labelsframe, text = "Label frequency").grid(row=1,sticky = 'w',padx = (23,0),pady=(3,0))
        self.l24=Label(self.gridlineframe, text = "Gridline frequency").grid(row=1,sticky = 'e',padx = (10,10))
        
        self.l12=Label(self.layoutframe, text= "Left").grid(row=1,column=0,sticky = 'e',padx = (10,5),pady=(10,1))
        self.l13=Label(self.layoutframe, text= "Right").grid(row=2,column=0,sticky = 'e',padx = (10,5))
        self.l14=Label(self.layoutframe, text= "Top").grid(row=1,column=3,sticky = 'e',padx = (25,5),pady=(10,1))
        self.l15=Label(self.layoutframe, text= "Bottom").grid(row=2,column=3,sticky = 'e',padx = (25,5))
        
        self.l17=Label(self.legendframe, text= "Legend position").grid(row=0,column=0,sticky = 'e',padx = 5,pady=(7,1))
        self.l18=Label(self.legendframe, text= "Number of columns").grid(row=1,column=0,sticky = 'e',padx = 5)
        
        self.l20=Label(self.optionframe, text= "Graph width").grid(row=27,column=0,sticky = 'e',padx = 5,pady = (10,0))
        self.l21=Label(self.optionframe, text= "Graph height").grid(row=28,column=0,sticky = 'e',padx = 5,pady = (1,10))
        
        
        ##########################################TGP2 sheet variables
        self.l21=Label(self.sheetoptionframe, text= "My locations are in column").grid(row=0,column=0,sticky = 'e',padx = 5,pady = (10,20))
        self.l21=Label(self.sheetoptionframe, text= "My arrival/departure labels are").grid(row=1,column=0,sticky = 'e',padx = 5,pady = (0,2))
        self.l21=Label(self.sheetoptionframe, text= "and they are in column").grid(row=2,column=0,sticky = 'e',padx = 5,pady = (0,20))
        self.l21=Label(self.sheetoptionframe, text= "My trains are in row").grid(row=3,column=0,sticky = 'e',padx = 5,pady = (0,20))
        self.l21=Label(self.sheetoptionframe, text= "My times are in cell").grid(row=4,column=0,sticky = 'e',padx = 5,pady = (0,0))
        self.l21=Label(self.sheetoptionframe, text= "to cell").grid(row=4,column=2,sticky = 'e',padx = 2,pady = (0,0))
        
        self.l21=Label(self.sheetframe, text= "Sheet name").grid(row=0,column=0,sticky = 'w',padx = (15,0),pady = (10,0)) #sheetname
        ######################################################
        
        #Set up each entry by defining a variable (used to extract entries later on), creating an entry field and filling it with a saved variable/default value.
        self.directory=tk.StringVar()
        self.e1=Entry(self.fileframe, textvariable=self.directory, width=38)
        if 'directory_save' in locals():
            self.e1.insert(0, directory_save)
        self.e1.grid(row=0, column=0,columnspan=3,sticky='w',padx=(75,0),pady=(10,0))
        
        self.infilename=tk.StringVar()
        self.e2=Entry(self.fileframe, textvariable=self.infilename, width=38)
        if 'infilename_save' in locals():
            self.e2.insert(0, infilename_save)
        self.e2.grid(row=1, column=0,columnspan=3,sticky='w',padx=(75,0),pady=(10,0))

        self.outfilename=tk.StringVar()
        self.e3=Entry(self.outputframe, textvariable=self.outfilename, width=21)
        if 'outfilename_save' in locals():
            self.e3.insert(0, outfilename_save)
        self.e3.grid(row=0, column=0,columnspan=3,sticky='w',padx=(115,0),pady=(10,0))


#        self.sheetname=tk.StringVar()
#        self.e24=Entry(self.sheetframe, textvariable=self.sheetname,width = 36)
#        self.e24.delete(0, "end")
#        if 'sheetname_save' in locals():
#            self.e24.insert(0, sheetname_save)
#        else:
#            self.e24.insert(0, 'UP Peak 12tph')
#        self.e24.grid(row=0, column=0,columnspan = 2,sticky = 'w',padx = (87,0),pady=(10,0))




        self.min_time=tk.StringVar()
        self.e4=Entry(self.optionframe, textvariable=self.min_time,width = 8)
        if 'min_time_save' in locals():
            self.e4.insert(0, min_time_save)
        self.e4.grid(row=0, column=1,sticky = 'w',pady = (10,0))

        self.max_time=tk.StringVar()
        self.e5=Entry(self.optionframe, textvariable=self.max_time,width = 8)
        if 'max_time_save' in locals():
            self.e5.insert(0, max_time_save)
        self.e5.grid(row=0, column=3,sticky = 'w',pady = (10,0))
        
        

        self.min_dist=tk.DoubleVar()
        self.e6=Entry(self.optionframe, textvariable=self.min_dist)
        self.e6.delete(0, "end")
        if 'min_dist_save' in locals():
            self.e6.insert(0, min_dist_save)
        #self.e6.grid(row=9, column=1)

        self.max_dist=tk.DoubleVar()
        self.e7=Entry(self.optionframe, textvariable=self.max_dist)
        self.e7.delete(0, "end")
        if 'max_dist_save' in locals():
            self.e7.insert(0, max_dist_save)
        #self.e7.grid(row=10, column=1)




        #Customising graph entries
        self.label_freq=tk.IntVar()
        self.e16=Entry(self.labelsframe, textvariable=self.label_freq,width = 8)
        self.e16.delete(0, "end")
        if 'label_freq_save' in locals():
            self.e16.insert(0, label_freq_save)
        else:
            self.e16.insert(0, 10)
        self.e16.grid(row=1,column=0,sticky='w',padx=123,pady=(3,0))
        
        
        
        
        self.grid_freq=tk.DoubleVar()
        self.e17=Entry(self.gridlineframe, textvariable=self.grid_freq,width = 8)
        self.e17.delete(0, "end")
        if 'grid_freq_save' in locals():
            self.e17.insert(0, grid_freq_save)
        else:
            self.e17.insert(0, 1)
        self.e17.grid(row=1, column=1,sticky='w')
        
        
        
        self.check5 = Checkbutton(self.labelsframe, text= "Display above graph", variable=self.check5var)
        if "axes_save" in locals():
            if axes_save == 'True':
                self.check5.select()
        self.check5.grid(row=0, sticky = 'w',padx = (10,5),pady=(3,0))
    
    
    

        self.l=tk.DoubleVar()
        self.e8=Entry(self.layoutframe, textvariable=self.l,width = 8)
        self.e8.delete(0, "end")
        if 'l_save' in locals():
            self.e8.insert(0, l_save)
        else:
            self.e8.insert(0, 0.14)
        self.e8.grid(row=1, column=1,sticky = 'w',pady=(10,1))

        self.r=tk.DoubleVar()
        self.e9=Entry(self.layoutframe, textvariable=self.r,width = 8)
        self.e9.delete(0, "end")
        if 'r_save' in locals():
            self.e9.insert(0, r_save)
        else:
            self.e9.insert(0, 0.93)
        self.e9.grid(row=2, column=1,sticky = 'w')

        self.t=tk.DoubleVar()
        self.e10=Entry(self.layoutframe, textvariable=self.t,width = 8)
        self.e10.delete(0, "end")
        if 't_save' in locals():
            self.e10.insert(0, t_save)
        else:
            self.e10.insert(0, 0.91)
        self.e10.grid(row=1, column=4,sticky = 'w',pady=(10,1))

        self.b=tk.DoubleVar()
        self.e11=Entry(self.layoutframe, textvariable=self.b,width = 8)
        self.e11.delete(0, "end")
        if 'l_save' in locals():
            self.e11.insert(0, b_save)
        else:
            self.e11.insert(0, 0.08)
        self.e11.grid(row=2, column=4,sticky = 'w')
        
        
        
        

        self.legend_position=tk.DoubleVar()
        self.e12=Entry(self.legendframe, textvariable=self.legend_position,width = 8)
        self.e12.delete(0, "end")
        if 'legend_position_save' in locals():
            self.e12.insert(0, legend_position_save)
        else:
            self.e12.insert(0, 1.09)
        self.e12.grid(row=0, column=1,pady=(7,1))

        self.ncols=tk.IntVar()
        self.e13=Entry(self.legendframe, textvariable=self.ncols,width = 8)
        self.e13.delete(0, "end")
        if 'ncols_save' in locals():
            self.e13.insert(0, ncols_save)
        else:
            self.e13.insert(0, 1)
        self.e13.grid(row=1, column=1)
        
        
        
        

        

        self.size_x=tk.DoubleVar()
        self.e14=Entry(self.optionframe, textvariable=self.size_x,width = 8)
        self.e14.delete(0, "end")
        if 'size_x_save' in locals():
            self.e14.insert(0, size_x_save)
        else:
            self.e14.insert(0, 8)
        self.e14.grid(row=27, column=1,sticky = 'w',pady = (10,0))

        self.size_y=tk.DoubleVar()
        self.e15=Entry(self.optionframe, textvariable=self.size_y,width = 8)
        self.e15.delete(0, "end")
        if 'size_y_save' in locals():
            self.e15.insert(0, size_y_save)
        else:
            self.e15.insert(0, 7)
        self.e15.grid(row=28, column=1,sticky = 'w',pady = (1,10))
        
        
###########################################################################TGP2 sheet variables        
        
       
        self.location_column=tk.StringVar()
        self.e18=Entry(self.sheetoptionframe, textvariable=self.location_column,width = 3)
        self.e18.delete(0, "end")
        if 'location_column_save' in locals():
            self.e18.insert(0, location_column_save)
        else:
            self.e18.insert(0, 'A')
        self.e18.grid(row=0, column=1,sticky = 'w',pady = (10,20))
        
        
        self.arrdep_labels=tk.StringVar()
        self.e19=Entry(self.sheetoptionframe, textvariable=self.arrdep_labels,width = 12)
        self.e19.delete(0, "end")
        if 'arrdep_labels_save' in locals():
            self.e19.insert(0, arrdep_labels_save)
        else:
            self.e19.insert(0, 'arr,dep,pass')
        self.e19.grid(row=1, column=1,sticky = 'w',pady = (0,2),columnspan =3)
        
        
        self.arrdep_column=tk.StringVar()
        self.e20=Entry(self.sheetoptionframe, textvariable=self.arrdep_column,width = 3)
        self.e20.delete(0, "end")
        if 'arrdep_column_save' in locals():
            self.e20.insert(0, arrdep_column_save)
        else:
            self.e20.insert(0, 'B')
        self.e20.grid(row=2, column=1,sticky = 'w',pady = (0,20))
        
        
        self.train_row=tk.StringVar()
        self.e21=Entry(self.sheetoptionframe, textvariable=self.train_row,width = 3)
        self.e21.delete(0, "end")
        if 'train_row_save' in locals():
            self.e21.insert(0, train_row_save)
        else:
            self.e21.insert(0, '2')
        self.e21.grid(row=3, column=1,sticky = 'w',pady = (0,20))
        
        
        self.time_cell_start=tk.StringVar()
        self.e22=Entry(self.sheetoptionframe, textvariable=self.time_cell_start,width = 4)
        self.e22.delete(0, "end")
        if 'time_cell_start_save' in locals():
            self.e22.insert(0, time_cell_start_save)
        else:
            self.e22.insert(0, 'C9')
        self.e22.grid(row=4, column=1,sticky = 'w',pady = (0,0))
        
        
        self.time_cell_end=tk.StringVar()
        self.e23=Entry(self.sheetoptionframe, textvariable=self.time_cell_end,width = 4)
        self.e23.delete(0, "end")
        if 'time_cell_end_save' in locals():
            self.e23.insert(0, time_cell_end_save)
        else:
            self.e23.insert(0, 'V58')
        self.e23.grid(row=4, column=3,sticky = 'w',pady = (0,0))
        
        
        self.sheetname=tk.StringVar()
        self.e24=Entry(self.sheetframe, textvariable=self.sheetname,width = 36)
        self.e24.delete(0, "end")
        if 'sheetname_save' in locals():
            self.e24.insert(0, sheetname_save)
        else:
            self.e24.insert(0, 'UP Scenario 10')
        self.e24.grid(row=0, column=0,columnspan = 2,sticky = 'w',padx = (87,0),pady=(10,0))
        

###########################################################################        

        self.check4 = Checkbutton(self.outputframe, text= "Preview?", variable=self.check4var)
        if "preview_save" in locals():
            if preview_save == 'True':
                self.check4.select()
        self.check4.grid(row=1,column=0,sticky='w',padx=(246,0),pady=(0,0))
        
        #Create list of entries and checkbuttons
        self.entry_list = [self.e1,self.e2,self.e3,self.e4,self.e5,self.e8,self.e9,self.e10,self.e11,self.e12,self.e13,self.e14,self.e15, self.e16, self.e17,self.e18,self.e19,self.e20,self.e21,self.e22,self.e23]
        self.entry_value_list = [self.directory, self.infilename, self.outfilename, self.min_time, self.max_time, self.l, self.r, self.t, self.b, self.legend_position, self.ncols, self.size_x, self.size_y]
        self.check_list = [self.check1, self.check2, self.check3, self.check4, self.check5]
        
        #Add buttons at the end.
        self.b1 =Button(self.outputframe, text="Generate!", command=self.runcmd,width=10, height = 1).grid(row=0,padx=(250,0),pady=(10,0),columnspan=2)
        self.b2=Button(master, text="Quit", command=self.cancelcmd)#.grid(row=30, column=1) 
        self.b3=Button(self.optionframe, text="Clear", command=self.clearcmd)#.grid(row=2,column=1,rowspan=10)
        self.b4=Button(self.optionframe, text="Reset options", command=self.defaultscmd).grid(row=28, column=1,columnspan=3,padx=(110,0))
        
        #Stop program if x button pressed.
        master.protocol('WM_DELETE_WINDOW', self.xwindow)         
    
    
        
    #Define command for OK button.    
    def runcmd(self):
        print("\n*******************************************************************************")
        print("If you're seeing this, and you have an error, please e-mail your options, timetable file and a picture of the black window to \nthariq.fahry@networkrail.co.uk")
        print("*******************************************************************************\n")
        
        errmsg1=[]
        errmsg2=[]
        #all_errmsg=[]
        
        for i in range(len(self.entry_value_list)):
            if not self.entry_list[i].get():
                errmsg1.append('error')
        if len(errmsg1) != 0:
            messagebox.showerror("Error", "All entries must be filled.")
            #all_errmsg.append('error')
            return
            
        if self.check1var.get() == False and self.check2var.get()==False:
            messagebox.showerror("Error", "No file format selected.")
            #all_errmsg.append('error')
            return
        '''No longer needed.
        if "\\" in self.directory.get():
            messagebox.showerror("Error", "Directory contains backslash. Change backslash to forwardslash.")
        '''
        '''
        float_variable_names=["Minimum distance", "Maximum distance", "Left position", "Right position", "Top position", "Bottom position", "Legend position", "Horizontal size", "Vertical size"]
        float_variables=[self.min_dist, self.max_dist, self.l, self.r, self.t, self.b, self.legend_position, self.size_x, self.size_y]
        for i in range(len(float_variables)):
            if isinstance(float_variables[i].get(), float)==False and isinstance(float_variables[i].get(),int)==False:
                messagebox.showerror("Error", float_variable_names[i]+" is not a number.")
                return

        if isinstance(self.ncols.get(), int) == False:
            messagebox.showerror("Error","Number of columns must be an integer.")       
            return
        '''    
        position_list=[self.l, self.r, self.b, self.t]
        position_name_list=["Left", "Right", "Bottom", "Top"]
        for i in range(len(position_list)):
            try:
                if position_list[i].get() < 0.0 or position_list[i].get() > 1.0:
                    errmsg2.append('error')
            except tk.TclError:
                messagebox.showerror("Error",position_name_list[i]+" position must be a number.")
                return
        #pattern = re.compile("^([0-9][0-9]:[0-9][0-9]:[0-9][0-9]*)")
        #if not pattern.match(self.e4.get()) or not pattern.match(self.e5.get()):
            #messagebox.showerror("Error", "Time must be in hh:mm:ss format.")
            #all_errmsg.append('error')
            #return
        if len(errmsg2) != 0:
            messagebox.showerror("Error", "All positions must be between 0 and 1.")
            #all_errmsg.append('error')
            return
        #if len(all_errmsg) == 0:
            #self.master.destroy()
        #Get variables from input box.
        directory=my_gui.directory.get()
        infilename=my_gui.infilename.get()
        outfilename=my_gui.outfilename.get()
        pdf=my_gui.check1var.get()
        png=my_gui.check2var.get()
        gridlines=my_gui.check3var.get()
        
        ####TGP2
        location_column =   my_gui.location_column.get() 
        arrdep_labels =     my_gui.arrdep_labels.get()
        arrdep_column =     my_gui.arrdep_column.get()
        train_row =         my_gui.train_row.get()
        time_cell_start =   my_gui.time_cell_start.get()
        time_cell_end =     my_gui.time_cell_end.get()
        sheetname =         my_gui.sheetname.get()
        
        
        try:
            min_time=dt.datetime.strptime(my_gui.min_time.get(),"%H:%M:%S")
        except ValueError:
            try:
                min_time=dt.datetime.strptime(my_gui.min_time.get(),"%d %H:%M:%S")
            except ValueError:
                messagebox.showerror("Error", "Time must be in hh:mm:ss or d hh:mm:ss format.")
                return
        try:
            max_time=dt.datetime.strptime(my_gui.max_time.get(),"%H:%M:%S")
        except ValueError:
            try:
                max_time=dt.datetime.strptime(my_gui.max_time.get(),"%d %H:%M:%S")
            except ValueError:
                messagebox.showerror("Error", "Time must be in hh:mm:ss or d hh:mm:ss format.")
                return
#        try: #TGP2 removed and replaced with auto dist inside stitch()
#            min_dist=my_gui.min_dist.get()
#        except (ValueError, tk.TclError):
#            messagebox.showerror("Error","Minimum distance must be a number.")
#            return
#        try:
#            max_dist=my_gui.max_dist.get()
#        except (ValueError, tk.TclError):
#            messagebox.showerror("Error","Maximum distance must be a number.")
#            return
        
        try:    
            label_freq=my_gui.label_freq.get()
        except (ValueError, tk.TclError):
            messagebox.showerror("Error","Label frequency must be a number.")
            return
        
        if label_freq < 1:
            messagebox.showerror("Error","Label frequency must be greater than 0.")
            return
            
        if (60%label_freq != 0) and (label_freq%60 != 0):
            messagebox.showerror("Error","Label frequency must be divisible by or a multiple of 60.")
            return
        
        if gridlines == True:
            try:    
                grid_freq=my_gui.grid_freq.get()
            except (ValueError, tk.TclError):
                messagebox.showerror("Error","Gridlines frequency must be a number.")
                return
            
            if grid_freq <= 0:
                messagebox.showerror("Error","Gridline frequency must be greater than 0.")
                return
        else:
            grid_freq = ''
        
        axes = my_gui.check5var.get()    
            
        l=my_gui.l.get()
        r=my_gui.r.get()
        t=my_gui.t.get()
        b=my_gui.b.get()
        try:
            legend_position=my_gui.legend_position.get()
        except (tk.TclError, ValueError):
            messagebox.showerror("Error","Legend position must be a number.")
            return
        try:
            ncols=my_gui.ncols.get()
        except (tk.TclError, ValueError):
            messagebox.showerror("Error","Number of columns must be an integer.")
            return
        try:
            size_x=my_gui.size_x.get()
        except (tk.TclError, ValueError):
            messagebox.showerror("Error","Horizontal size must be a number.")
            return
        try:
            size_y=my_gui.size_y.get()
        except (tk.TclError, ValueError):
            messagebox.showerror("Error","Vertical size must be a number.")
            return
            
        preview=my_gui.check4var.get()
        
        #Change any backslashes in directory.
        directory=directory.replace('\\', '/')
        
        if os.path.exists(directory) == False:
            messagebox.showerror("Error", "Directory not found.")
            return
        
        if os.path.isfile(directory+"/"+infilename) == False:
            messagebox.showerror("Error", "Input file not found in directory")
            return        
        
        #Navigate to directory containing the input files.
        os.chdir(directory)
        
        #Read the file with times and distances. Everything is read as a string.

        try:
            #y = np.genfromtxt(infilename, dtype=str, delimiter=',') #legacy
            
            minrow = (''.join(filter(str.isdigit, time_cell_start)))
            maxrow = (''.join(filter(str.isdigit, time_cell_end)))
            
            mincol = (''.join(filter(str.isalpha, time_cell_start)))
            maxcol = (''.join(filter(str.isalpha, time_cell_end)))
                              
                        
            locations, indices3, exceptions3 = readexcel(infilename,sheetname,'{}{}:{}{}'.format(location_column,minrow,location_column,maxrow),False,True)
            arrdep, indices2 , exceptions2 = readexcel(infilename,sheetname,'{}{}:{}{}'.format(arrdep_column,minrow,arrdep_column,maxrow),False,True)
            trains, indices1, exceptions1 = readexcel(infilename,sheetname,'{}{}:{}{}'.format(mincol,train_row,maxcol,train_row),False,True)
            cells, z, exceptions0 = readexcel(infilename,sheetname,'{}{}:{}{}'.format(mincol,minrow,maxcol,maxrow),True,False) #read(cell_range, only_times, as_string)
            
            y, max_dist = stitch(locations,trains,cells,arrdep,arrdep_labels)
            
            max_dist = max_dist+1
            min_dist = 0

            
            
            ##############debug only###############
            #e(y,'refdata.xlsx','VE slice1')
            #e(exceptions0,'refdata.xlsx','exc')
            #######################################
                        
        except ValueError as error123:
            messagebox.showerror("Error","ValueError raised. This should not usually happen; possible bug {} ".format(error123))
            return
        #Find split between up and down trains.
        n = np.shape(y)[0]
        #print(n)
        test=np.zeros(n)
        for i in range(n):
            if np.all(y[i,:]== ''):
                test[i]=1
        if len(np.where(test == 1)[0]) > 1:
            
            y = np.delete(y, (np.where(test == 1)[0][1:]), axis=0)
            z = np.delete(z, (np.where(test == 1)[0][1:]), axis=0)
            #y[1] = np.delete(y[1], (np.where(test == 1)[0][1:]), axis=0)
            
            n = np.shape(y)[0]
            
       
            
        if np.all(test==0) == True:
            n_directions = 1
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error", "No space between up and down trains. Graph cannot be produced.")
            root.destroy()
            return
            #raise SystemExit
        else:
            n_directions = 2
        #Remove extra empty lines at the end.
        empty_lines = []
        for i in range(n-1,0,-1):
            if np.all(y[i,:]== ''):
                empty_lines.append(i)
            else:
                break
            
        y = np.delete(y, (np.asarray(empty_lines)), axis=0)
        z = np.delete(z, (np.asarray(empty_lines)), axis=0)
        
        
        #Separate arrays for up and down trains.
        a = np.where(test == 1)[0][0]
        b1 = np.arange(n-a)+a
        b2 = np.arange(a+1)
        
        x1 = np.delete(y, (b1), axis=0)
        z1 = np.delete(z, (b1), axis=0)
        
        x2 = np.delete(y, (b2), axis=0)
        z2 = np.delete(z, (b2), axis=0)
        
        #Delete empty columns/
        n2 = np.shape(x1)[1]
        test2  =np.zeros(n2)
        for i in range(n2):
            if np.all(x1[:,i]==''):
                test2[i] =1
                
        c = np.where(test2 == 1)[0]
        x1 = np.delete(x1, (c), axis=1)
        z1 = np.delete(z1, (c), axis=1)
        
        n2 = np.shape(x2)[1]
        test2  =np.zeros(n2)
        for i in range(n2):
            if np.all(x2[:,i]==''):
                test2[i] =1
                
        c = np.where(test2 == 1)[0]
        x2 = np.delete(x2, (c), axis=1)
        z2 = np.delete(z2, (c), axis=1)
        
        #Remove adjustments and allowances.
        n = np.shape(x1)[0]
        test = np.zeros(n, dtype=np.bool_)
        for i in range(n): #Find where allowance and adjustment lines are.
            if 'Allowance' in x1[i,0]:
                test[i] = True
            elif 'Adjustment' in x1[i,0]:
                test[i] = True
            else:
                test[i] = False
        a = np.where(test == True)
        x1 = np.delete(x1, (a), axis=0) #Delete allowance and adjustment lines.
        z1 = np.delete(z1, (a), axis=0) #Delete allowance and adjustment lines.
        
        n = np.shape(x2)[0]
        test = np.zeros(n, dtype=np.bool_)
        for i in range(n): #Find where allowance and adjustment lines are.
            if 'Allowance' in x2[i,0]:
                test[i] = True
            elif 'Adjustment' in x2[i,0]:
                test[i] = True
            else:
                test[i] = False
        a = np.where(test == True)
        x2 = np.delete(x2, (a), axis=0) #Delete allowance and adjustment lines.
        z2 = np.delete(z2, (a), axis=0)
                      
                      
        #Remove platforms.
        plat_pos = np.where(x1[:,0] == 'Platform') #Find where platforms are.
        x1 = np.delete(x1, (plat_pos), axis=0) #Remove platforms from timetable.
        z1 = np.delete(z1, (plat_pos), axis=0) #Remove platforms from timetable.
        plat_pos2 = np.where(x2[:,0] == 'Platform') #Find where platforms are.
        x2 = np.delete(x2, (plat_pos2), axis=0) #Remove platforms from timetable.
        z2 = np.delete(z2, (plat_pos2), axis=0)
        
        #Remove lines.
        line_pos = np.where(x1[:,0] == 'Line') #Find where lines are.
        x1 = np.delete(x1, (line_pos), axis=0) #Remove lines from timetable.
        z1 = np.delete(z1, (line_pos), axis=0) #Remove lines from timetable.
        line_pos = np.where(x2[:,0] == 'Line') #Find where lines are.
        
        x2 = np.delete(x2, (line_pos), axis=0)
        z2 = np.delete(z2, (line_pos), axis=0)#Remove lines from timetable.
        
        #Get labels.
        labels1 = x1[0:2,:]
        labels2 = x2[0:2,:]
        labels1 = np.delete(labels1, (0,1), axis=1)
        labels2 = np.delete(labels2, (0,1), axis=1)
        labels = np.concatenate((labels1, labels2), axis=1)
        
        #Delete labels from arrays.
        x1 = np.delete(x1, (0,1), axis=0)
        z1 = np.delete(z1, (0,1), axis=0)
        x2 = np.delete(x2, (0,1), axis=0)
        z2 = np.delete(z2, (0,1), axis=0)
        
        #Put distances in separate array.
        dist = x1[:,1] #raw distance array (of up trains only?)
        n=np.size(dist)
        
        #Put locations and distances into separate array.
        
        location = OrderedSet(x1[:,0]) #orderedset will remove duplicates in col 0 of x1
        location = list(filter(lambda x: x!= '', location)) #remove blanks, convert back to list
        location1 = list(x1[:,0]) #un de-duped, uncleaned list of locations
        location2 = list(x2[:,0])
        dist1 = OrderedSet(x1[:,1]) #orderedset will remove duplicates in col 0 of x1
        dist1 = list(filter(lambda x: x!= '', dist1)) #remove blanks, convert back to list
        
        
        #Check distances are in correct format.
        for i in range(len(dist)):
            try: 
                distance = float(dist[i])
            except ValueError:
                #loc = location1[i]
                #if loc == "": 
                #    loc = location1[i-1]
                #messagebox.showerror("Error", "Incorrect distance format at "+loc+". Distance must be a number.")
                messagebox.showerror("Error","Incorrect distance format in cell "+str(exceladdressof(z1[:,1][i]))+". Distance must be a number.")
                
                return
                    
        #Convert distances to be marked on the graph to floats.
        dist2 = []
        try:
            for i in range(np.size(dist1)):
            	dist2.append(float(dist1[i]))          #dist2 contains the floats of distances
        except ValueError:
            messagebox.showerror("Error","non-number in dist array not caught. This should not happen.")
            #messagebox.showerror("Error","Incorrect distance1 format in cell "+str(exceladdressof(z1[:,1][i]))+". Distance must be a number.")
            
        #Test for missing distances and locations.
        if len(location) > len(dist1):
            messagebox.showerror("Error", "Missing distance.")
            return
         
        #Create arrays for train times.
        n21 = np.shape(x1)[1]-2 #num of trains?
        n_1 = np.shape(x1)[0]   #raw num of locations?
        n22 = np.shape(x2)[1]-2
        n_2 = np.shape(x2)[0]

        
        #Remove spaces and delete any blank lines.
        empty_lines = []
        for j in range(n21):
            for i in range(n_1):
                try:
                    test = mdates.date2num(dt.datetime.strptime(x1[:,j+2][i],"%d %H:%M:%S"))
                except ValueError:
                    try:
                        test = mdates.date2num(dt.datetime.strptime(x1[:,j+2][i],"%d %H:%M"))
                    except ValueError:
                        x1[:,j+2][i] = x1[:,j+2][i].replace(' ', '')
            if np.all(x1[:,j+2]== ''):
                empty_lines.append(i)
        x1 = np.delete(x1, (np.asarray(empty_lines)), axis=0)
        z1 = np.delete(z1, (np.asarray(empty_lines)), axis=0)
#################################################################################################### x1 is no longer mutated after this point        
        n21 = np.shape(x1)[1]-2
        n_1 = np.shape(x1)[0]
        list_times1 = np.zeros([n21,n_1])
        
        
        
        empty_lines = []
        for j in range(n22):
            for i in range(n_2):
                try:
                    test = mdates.date2num(dt.datetime.strptime(x2[:,j+2][i],"%d %H:%M:%S"))
                except ValueError:
                    try:
                        test = mdates.date2num(dt.datetime.strptime(x2[:,j+2][i],"%d %H:%M"))
                    except ValueError:
                        x2[:,j+2][i] = x2[:,j+2][i].replace(' ', '')
            if np.all(x2[:,j+2]== ''):
                empty_lines.append(i)
        x2 = np.delete(x2, (np.asarray(empty_lines)), axis=0)
        z2 = np.delete(z2, (np.asarray(empty_lines)), axis=0)
#################################################################################################### x2 is no longer mutated after this point                
        n22 = np.shape(x2)[1]-2
        n_2 = np.shape(x2)[0]
        list_times2 = np.zeros([n22,n_2])
        
        
        #Convert times into numpy datetime objects, replacing empty values by None.
        over_midnight = False
        for j in range(n21):
            for i in range(n_1):
                if x1[:,j+2][i] == '':
                    list_times1[j,i] = np.nan
                    continue
                try:
                    list_times1[j,i] = mdates.date2num(dt.datetime.strptime(x1[:,j+2][i],"%H:%M:%S"))
                except ValueError:
                    try:
                       list_times1[j,i] = mdates.date2num(dt.datetime.strptime(x1[:,j+2][i],"%H:%M"))
                    except ValueError:
                        try:
                           list_times1[j,i] = mdates.date2num(dt.datetime.strptime(x1[:,j+2][i],"%d %H:%M"))
                           over_midnight = True
                        except ValueError:
                            try:
                               list_times1[j,i] = mdates.date2num(dt.datetime.strptime(x1[:,j+2][i],"%d %H:%M:%S"))
                               over_midnight = True
                            except ValueError:
                                loc = location1[i]
                                if loc == "": 
                                    loc = location1[i-1]
                                #messagebox.showerror("Error", "Incorrect time format for train "+ str(j+1) + " at "+loc+". Time must be in hh:mm:ss or d hh:mm:ss format.")
                                messagebox.showerror("Error","Incorrect time format in cell "+str(exceladdressof(z1[:,j+2][i])))
                                return
        
        for j in range(n22):
            for i in range(n_2):
                if x2[:,j+2][i] == '':
                    list_times2[j,i] = np.nan
                    continue
                try:
                    list_times2[j,i] = mdates.date2num(dt.datetime.strptime(x2[:,j+2][i],"%H:%M:%S"))
                except ValueError:
                    try:
                        list_times2[j,i] = mdates.date2num(dt.datetime.strptime(x2[:,j+2][i],"%H:%M"))
                    except ValueError:
                        try:
                           list_times2[j,i] = mdates.date2num(dt.datetime.strptime(x2[:,j+2][i],"%d %H:%M"))
                           over_midnight = True
                        except ValueError:
                            try:
                               list_times2[j,i] = mdates.date2num(dt.datetime.strptime(x2[:,j+2][i],"%d %H:%M:%S"))
                               over_midnight = True
                            except ValueError:
                                loc = location2[i]
                                if loc == "": 
                                    loc = location2[i-1]
                                
                                messagebox.showerror("Error","Incorrect time format in cell "+str(exceladdressof(z2[:,j+2][i])))
                                return
        #Combine up and down trains.
        if np.shape(list_times1)[1] != np.shape(list_times2)[1]:
            messagebox.showerror("Error","Up and Down trains do not have the same number of locations or there are extra empty lines in your input file. Graph cannot be produced.")
            return
        
        #TGP2: sections here commented out so that we plot only Up trains
        n = np.shape(list_times1)[1]
        n21 = np.shape(list_times1)[0] 
        n22 = np.shape(list_times2)[0]
        n2 = n21 #+ n22
        list_times=np.zeros([n2,n])
        for i in range(n21):
            list_times[i,:] = list_times1[i,:]
        #for i in range(n22):
        #    list_times[i+n21,:] = np.flipud(list_times2[i,:])
        
        #Create figure and plot.
        fig = plt.figure(facecolor='white')
        ax = fig.add_subplot(111)
        
        ax.tick_params(axis='x',which='minor',bottom='off')
        if over_midnight == True:
            hms = DateFormatter("%d %H:%M:%S")
        else:
            hms = DateFormatter("%H:%M:%S")
        ax.xaxis.set_major_formatter(hms)
        #ax.xaxis.set_tick_params(labeltop='on')
        val=0
        arr = []
        if label_freq < 60:
            while val < 60:
                arr.append(val)
                val = val + label_freq
            tenmin = MinuteLocator(byminute=arr, interval=1)
        else:
            while val < 24:
                arr.append(val)
                val = val + label_freq/60
            tenmin = HourLocator(byhour=arr, interval=1)
        tenmin.MAXTICKS = 10000
        ax.xaxis.set_major_locator(tenmin)
        if gridlines == True:
            grid = grid_freq
        else:
            grid = 1
        minor_locator = AutoMinorLocator(label_freq/grid)
        minor_locator.MAXTICKS = 10000
        if gridlines == True:
            ax.xaxis.set_minor_locator(minor_locator)
            ax.grid(b= True, which='major', linestyle='-', color='0.4', zorder=0)
            ax.grid(b= True, which='minor', linestyle='-', color='0.7', zorder=0)
        
        
        for j in range(n2):
            #Check colours and labels
            if labels[0][j] == '':
                messagebox.showerror("Error", "No label found for train "+str(j+1)+".")
                return
            if (labels[1][j] not in matplotlib.colors.cnames) and (labels[1][j] not in ['w', 'k', 'r', 'b', 'y', 'g', 'm', 'c']):
                messagebox.showerror("Error", " \"" + labels[1][j]+ "\" is not a valid colour for train "+str(j+1)+".")
                return
                pass
            ax.plot(list_times[j][np.isfinite(list_times[j])], dist[np.isfinite(list_times[j])], color=labels[1][j], zorder=3)


        
        #if gridlines == True:
            #ax2.xaxis.set_minor_locator(minor_locator)
            #plt.grid(b= True, which='major', linestyle='-', color='0.4', zorder=0)
            #plt.grid(b= True, which='minor', linestyle='-', color='0.7', zorder=0)
                
        #Make the times look nice.
        #plt.gcf().autofmt_xdate()
        for label in ax.get_xticklabels():
            label.set_ha("right")
            label.set_rotation(30)

        

        
        
        #Set axis labels.
        ax.set_xlabel('Time')
        ax.yaxis.set_ticks(dist2)
        ax.yaxis.set_ticklabels(location)
        #Set font size for y tick labels.
        ytickfontsize = [tick.label.set_fontsize(7) for tick in ax.yaxis.get_major_ticks()]
        #Set axis limits. 
        plt.axis([min_time,max_time,min_dist,max_dist])
        #Create handles for legend.
        graph_labels = OrderedSet(labels[0])
        graph_labels = list(filter(lambda x: x!= '', graph_labels))
        colours = OrderedSet(labels[1])
        colours = list(filter(lambda x: x!= '', colours))
        n_labels = len(graph_labels)
        n_colours = len(colours)
        #Test if there are the same number of labels as colours.
        if n_labels < n_colours:
            messagebox.showwarning("Warning", "More colours than labels. Legend may not appear as expected.")
        if n_labels > n_colours:
            messagebox.showerror("Error", "More labels than colours. Graph cannot be produced.")
            return
            #raise SystemExit
            
        line = []
        for i in range(n_labels):
             line1 = mlines.Line2D([], [], color=colours[i], label=graph_labels[i])
             line.append(line1)
        #Create legend.
        plt.legend(handles=line, loc='upper right', ncol=ncols, bbox_to_anchor=(0.9, legend_position))
        #Change figure size.
        fig = plt.gcf()
        fig.set_size_inches(size_x, size_y)
        fig.subplots_adjust(left=l, bottom=b, top=t, right=r)
        #Add second x axis
        if axes == True:
            ax2 = ax.twiny()
            ax2.tick_params(axis='x',which='minor',bottom='off')
    
            X2tick_location= ax.xaxis.get_ticklocs() #Get the tick locations in data coordinates as a numpy array
            ax2.set_xticks(X2tick_location)
            ax2.set_xticklabels(X2tick_location)
            ax2.set_xlim([min_time, max_time])
            #ax2.xaxis_date()
            ax2.xaxis.set_major_formatter(hms)
    
            '''
            ax2.xaxis.set_major_locator(tenmin)
            '''
            for label in ax2.get_xticklabels():
                label.set_ha("left")
                label.set_rotation(30)
        
        if preview == True:
            def savecmd():
                root2.destroy()
                if pdf == True:
                    try:
                        plt.savefig(outfilename+'.pdf')
                    except PermissionError:
                        messagebox.showerror("Error","Output file is open. Close output file and try again.")
                        return
                    if png == False:
                        messagebox.showinfo("Graph produced", 'Graph has been written to '+outfilename+'.pdf'+' in '+directory+'.')
                if png == True:
                    try:
                        plt.savefig(outfilename+'.png', dpi=600)
                    except PermissionError:
                        messagebox.showerror("Error", "Output file is open. Close output file and try again.")
                        return
                    if pdf == False:
                        messagebox.showinfo("Graph produced", 'Graph has been written to '+outfilename+'.png'+' in '+directory+'.')
                if pdf == True and png == True:
                    messagebox.showinfo("Graphs produced", 'Graphs have been written to '+outfilename+'.pdf and '+outfilename+'.png'+' in '+directory+'.')
                plt.close()
                
                #Save graph variables.
                home_path = os.path.expanduser("~")
                if os.path.exists("H:\\") == True:
                    save_path = 'H:\Train Graph Plotter'
                else:
                    save_path = home_path+'\Train Graph Plotter'
                if os.path.exists(save_path) == False:
                        os.mkdir(save_path)
                f = open(save_path+'\saved_variables_v2.0Alpha.txt', 'w')
                variablelist=[directory, infilename, outfilename, pdf, png, gridlines, min_time, max_time, min_dist, max_dist, label_freq, grid_freq, axes, l, b, t, r, legend_position, ncols, size_x, size_y, preview, location_column, arrdep_labels, arrdep_column, train_row, time_cell_start, time_cell_end,sheetname]
                for item in variablelist:
                  f.write("%s\n" % item)
                f.close()
                
            def cancelcmd2():
                root2.destroy()
                plt.close()

#         
            #plt.savefig('H:/Train Graph Plotter/temp.pdf')
            #__f_tmp=glob.glob(pdfImg.PyGs({}).make_img_from_pdf("H:/Train Graph Plotter/temp.pdf")[1])[0]
            #__img=Image.open(__f_tmp)

            #__tk_img=tk.PhotoImage(__img)

            root2 = tk.Toplevel()
            root2.title("Graph preview")
            iconfile= open(tempfile,"wb")
            iconfile.write(icondata)
            iconfile.close()
            root2.iconbitmap(tempfile)
            os.remove(tempfile)            
            f1= tk.Frame(root2)        
            save_button = Button(f1, text="Save graph", command=savecmd)
            cancel_button = Button(f1, text="Cancel", command=cancelcmd2)
            save_button.pack(side=tk.LEFT)
            cancel_button.pack(side=tk.LEFT)
            f1.pack()
            f2 = tk.Frame(root2)
            vbar=tk.Scrollbar(f2, orient = tk.VERTICAL)
            vbar.pack(side=tk.RIGHT, fill=tk.Y)
            #canvas = tk.Canvas(f2)
            #canvas.create_image(0,0, image=__tk_img)
            canvas = FigureCanvasTkAgg(fig, master=f2)#, yscrollcommand=scrollbar.set)
            try:
                canvas.show() #will be depreceated in the future, use draw() instead
            except RuntimeError:
                messagebox.showerror("Error", "Graph contains too many tickmarks. Try increasing the label and gridline frequencies or reduce the size of the graph.")
                root2.destroy()
                plt.close()
                return
            canvas.get_tk_widget().pack(side=tk.BOTTOM,fill=tk.BOTH, expand=1)
                   
            hbar=tk.Scrollbar(f2, orient=tk.HORIZONTAL)
            hbar.pack(side=tk.TOP,fill=tk.X)

            canvas.get_tk_widget().config(xscrollcommand=hbar.set, yscrollcommand = vbar.set, scrollregion=(0,0, 2000, 2000))
            hbar.config(command=canvas.get_tk_widget().xview)
            vbar.config(command=canvas.get_tk_widget().yview)
            

            #canvas.config(scrollregion=(left, top, right, bottom))
            f2.pack()
            root2.resizable(0,0)
            return
        
        
        
        #Save figure and close plot.
        if pdf == True:
            try:
                plt.savefig(outfilename+'.pdf')
            except PermissionError:
                messagebox.showerror("Error","Output file is open. Close output file and try again.")
                return
            except RuntimeError:
                messagebox.showerror("Error", "Graph contains too many tickmarks. Try increasing the label and gridline frequencies or reduce the size of the graph.")
                plt.close()
                return
            if png == False:
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo("Graph produced", 'Graph has been written to '+outfilename+'.pdf'+' in '+directory+'.')
                root.destroy()
        if png == True:
            try:
                plt.savefig(outfilename+'.png', dpi=600)
            except PermissionError:
                messagebox.showerror("Error", "Output file is open. Close output file and try again.")
                return
            if pdf == False:
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo("Graph produced", 'Graph has been written to '+outfilename+'.png'+' in '+directory+'.')
                root.destroy()
        if pdf == True and png == True:
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Graphs produced", 'Graphs have been written to '+outfilename+'.pdf and '+outfilename+'.png'+' in '+directory+'.')
            root.destroy()
        plt.close()
        
        #Save graph variables.
        home_path = os.path.expanduser("~")
        if os.path.exists("H:\\") == True:
            save_path = 'H:\Train Graph Plotter'
        else:
            save_path = home_path+'\Train Graph Plotter'
        if os.path.exists(save_path) == False:
                os.mkdir(save_path)
        f = open(save_path+'\saved_variables_v2.0Alpha.txt', 'w')
        variablelist=[directory, infilename, outfilename, pdf, png, gridlines, min_time, max_time, min_dist, max_dist, label_freq, grid_freq, axes, l, b, t, r, legend_position, ncols, size_x, size_y, preview,location_column, arrdep_labels, arrdep_column, train_row, time_cell_start, time_cell_end,sheetname]
        for item in variablelist:
          f.write("%s\n" % item)
        f.close()

    #Define command for cancel button.
    def cancelcmd(self):
        #result=messagebox.askyesno("Quit", "Are you sure you want to quit?")
        if True:
            self.master.destroy()
            raise SystemExit
        else:
            return
    
    #Define command for clear button.
    def clearcmd(self):
        for i in range(len(self.entry_list)):
            self.entry_list[i].delete(0, "end")
        for i in range(len(self.check_list)):
            self.check_list[i].deselect()
     
    #Define command for restore defaults button.       
    def defaultscmd(self):
        for i in range(len(self.entry_list)):
            self.entry_list[i].delete(0, "end")
        for i in range(len(self.check_list)):
            self.check_list[i].deselect()
        self.e8.insert(0, 0.14)
        self.e9.insert(0, 0.93)
        self.e10.insert(0, 0.91)
        self.e11.insert(0, 0.08)
        self.e12.insert(0, 1.11)
        self.e13.insert(0, 1)
        self.e14.insert(0, 8)
        self.e15.insert(0, 7)
        self.e16.insert(0, 10)
        self.e17.insert(0, 1)
        self.e18.insert(0, 'A')
        self.e19.insert(0, 'arr,dep,pass')
        self.e20.insert(0, 'B')
        self.e21.insert(0, 2)
        self.e22.insert(0, 'C9')
        self.e23.insert(0, 'V58')
        self.check1.select()
    
    #Define command for closing window using X in top right hand corner.    
    def xwindow(self):
        #result=messagebox.askyesno("Quit", "Are you sure you want to quit?")
        if True:
            self.master.destroy()
            raise SystemExit
        else:
            return

    
           
            
#Create input box
input_ = tk.Tk()
my_gui = GUI(input_)
input_.mainloop()