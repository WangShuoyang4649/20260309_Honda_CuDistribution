import os
import sys
import datetime
from pathlib import Path

import math
import numpy as np
import pandas as pd

import subprocess
import win32gui
import win32console

import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from mpl_toolkits.axes_grid1 import make_axes_locatable

import tkinter as Tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

## Debug mode Handler
##   sys.tracebacklimit = 0: hide all traceback messages
##   sys.tracebacklimit = 1000: default value, show almost all traceback errors
sys.tracebacklimit = 0

## Log file record
def logoutput(filepath, logmessage, flag, logtype, promptflag):
    ## writing log file
    if len(filepath) > 0:
        file = open(filepath, flag)
        file.write('** '+logtype+' ** '+logmessage+'\n')
        file.close()
    ## prompt window or just output silently
    if promptflag:
        if logtype == 'Error':
            messagebox.showerror('Uniformity Evaluation Error', logmessage)
            raise HaltException('** '+logtype+' ** '+logmessage)
        elif logtype == 'Final':
            print('** '+logtype+' ** '+logmessage)
            messagebox.showinfo('Uniformity Evaluation Final', logmessage)
    else:
        print('** '+logtype+' ** '+logmessage)

## Rename prompt window name
win32console.SetConsoleTitle("Uniformity Evaluation Tool")

## Window Handler
def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


################
## GUI session
################
class UniformityEvaluation(Tk.Frame):

    def __init__(self,master):
        '''Initialization.'''

        ## define transfer information
        self.pt = PwParaTransfer()

        ## GUI Frame Creation
        super().__init__(master)
        self.canvas_create()
        ## create frame of Inputs input
        istart = self.all_inputs_frame_create(0)

    def canvas_create(self):
        '''Create canvas.'''

        ## start creating frame
        self.master.title('UniformityEvaluation')
        canvassize = str(self.pt.canvasX0)+'x'+str(self.pt.canvasY0)
        self.master.geometry(canvassize)

        ## create frame for base scene path input
        self.top = ttk.Frame(self.master, padding=10)
        self.top.grid(column=0, row=0, columnspan=4, rowspan=20, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

    def all_inputs_frame_create(self,istart):
        '''Create all inputs.'''

        iinputtitle = istart
        self.input_title_label = ttk.Label(self.top, text="** 入力設定 **", padding=(0,0), font=(15))
        self.input_title_label.grid(column=0, row=iinputtitle, sticky=Tk.W)

        idatafilepathinput = iinputtitle+1
        self.data_file_path_label = ttk.Label(self.top, text="データファイル（Excel）パス ", padding=(0,0), font=(15))
        self.data_file_path_label.grid(column=0, row=idatafilepathinput, sticky=Tk.E)
        self.data_file_path_value = Tk.StringVar(value = self.pt.data_file_path)
        self.data_file_path_entry = ttk.Entry(self.top, textvariable=self.data_file_path_value, width=self.pt.allinputlen)
        self.data_file_path_entry.grid(column=1, row=idatafilepathinput, sticky=Tk.W)
        self.data_file_path_entry.xview_moveto(1)
        self.data_file_path_button_open = ttk.Button(self.top, text='開く', command=self.open_data_file)
        self.data_file_path_button_open.grid(column=2, row=idatafilepathinput, sticky=Tk.W)

        isheetnumber = idatafilepathinput+1
        self.sheet_number_label = ttk.Label(self.top, text="シート番号 ", padding=(0,0), font=(15))
        self.sheet_number_label.grid(column=0, row=isheetnumber, sticky=Tk.E)
        self.sheet_number_value = Tk.IntVar(value = self.pt.sheet_number)
        self.sheet_number_entry = ttk.Entry(self.top, textvariable=self.sheet_number_value, width=self.pt.allinputlen0)
        self.sheet_number_entry.grid(column=1, row=isheetnumber, sticky=Tk.W)
        self.sheet_number_entry.bind("<Return>", lambda event: self.read_excel_data())

        ileftuppercorner = isheetnumber+1
        self.left_upper_corner_label = ttk.Label(self.top, text="データの左上のセル位置 ", padding=(0,0), font=(15))
        self.left_upper_corner_label.grid(column=0, row=ileftuppercorner, sticky=Tk.E)
        self.left_upper_corner_value = Tk.StringVar(value = self.pt.left_upper_corner)
        self.left_upper_corner_entry = ttk.Entry(self.top, textvariable=self.left_upper_corner_value, width=self.pt.allinputlen0)
        self.left_upper_corner_entry.grid(column=1, row=ileftuppercorner, sticky=Tk.W)
        self.left_upper_corner_entry.bind("<Return>", lambda event: self.update_width_height_show())

        irightlowercorner = ileftuppercorner+1
        self.right_lower_corner_label = ttk.Label(self.top, text="データの右下のセル位置 ", padding=(0,0), font=(15))
        self.right_lower_corner_label.grid(column=0, row=irightlowercorner, sticky=Tk.E)
        self.right_lower_corner_value = Tk.StringVar(value = self.pt.right_lower_corner)
        self.right_lower_corner_entry = ttk.Entry(self.top, textvariable=self.right_lower_corner_value, width=self.pt.allinputlen0)
        self.right_lower_corner_entry.grid(column=1, row=irightlowercorner, sticky=Tk.W)
        self.right_lower_corner_entry.bind("<Return>", lambda event: self.update_width_height_show())

        iwidthheightshow = irightlowercorner+1
        self.width_height_show_label = ttk.Label(self.top, text="データセル位置の範囲（幅 x 高さ） ", padding=(0,0), font=(15))
        self.width_height_show_label.grid(column=0, row=iwidthheightshow, sticky=Tk.E)
        self.width_height_show_value = Tk.StringVar(value = self.pt.width_height_show)
        self.width_height_show_entry = ttk.Entry(self.top, textvariable=self.width_height_show_value, width=self.pt.allinputlen0)
        self.width_height_show_entry.grid(column=1, row=iwidthheightshow, sticky=Tk.W)
        self.width_height_show_entry.config(state='disabled')

        iblank = iwidthheightshow+1
        self.Region1 = ttk.Label(self.top, text="", padding=(0,0), font=(15))
        self.Region1.grid(column=0, row=iblank, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

        istridexpattern = iblank+1
        self.stride_x_pattern_label = ttk.Label(self.top, text="分割の幅のパターン（X方向） ", padding=(0,0), font=(15))
        self.stride_x_pattern_label.grid(column=0, row=istridexpattern, sticky=Tk.E)
        self.stride_x_pattern_value = Tk.StringVar(value = self.pt.stride_x_pattern)
        self.stride_x_pattern_entry = ttk.Entry(self.top, textvariable=self.stride_x_pattern_value, width=self.pt.allinputlen0)
        self.stride_x_pattern_entry.grid(column=1, row=istridexpattern, sticky=Tk.W)

        istrideypattern = istridexpattern+1
        self.stride_y_pattern_label = ttk.Label(self.top, text="分割の幅のパターン（Y方向） ", padding=(0,0), font=(15))
        self.stride_y_pattern_label.grid(column=0, row=istrideypattern, sticky=Tk.E)
        self.stride_y_pattern_value = Tk.StringVar(value = self.pt.stride_y_pattern)
        self.stride_y_pattern_entry = ttk.Entry(self.top, textvariable=self.stride_y_pattern_value, width=self.pt.allinputlen0)
        self.stride_y_pattern_entry.grid(column=1, row=istrideypattern, sticky=Tk.W)

        istridecombination = istrideypattern+1
        self.stride_combination_label = ttk.Label(self.top, text="プロット用の分割の幅（X,Y） ", padding=(0,0), font=(15))
        self.stride_combination_label.grid(column=0, row=istridecombination, sticky=Tk.E)
        self.stride_combination_value = Tk.StringVar(value = self.pt.stride_combination)
        self.stride_combination_entry = ttk.Entry(self.top, textvariable=self.stride_combination_value, width=self.pt.allinputlen0)
        self.stride_combination_entry.grid(column=1, row=istridecombination, sticky=Tk.W)

        ibutton0 = istridecombination+1
        self.plot_data_division_button = ttk.Button(self.top, text='データ+分割プロット', command=self.plot_data_division)
        self.plot_data_division_button.grid(column=0, row=ibutton0, sticky=Tk.E)

        iblank = ibutton0+1
        self.Region2 = ttk.Label(self.top, text="", padding=(0,0), font=(15))
        self.Region2.grid(column=0, row=iblank, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

        icutinjection = iblank+1
        self.cut_injection_label = ttk.Label(self.top, text="注入物をクリップする ", padding=(0,0), font=(15))
        self.cut_injection_label.grid(column=0, row=icutinjection, sticky=Tk.E)
        self.cut_injection_value = Tk.BooleanVar(value = self.pt.cut_injection_flag)
        self.cut_injection_checkbutton = Tk.Checkbutton(self.top, variable=self.cut_injection_value)
        self.cut_injection_checkbutton.grid(column=1, row=icutinjection, sticky=Tk.W)

        icuconcentrationlowerthreshold = icutinjection+1
        self.cu_concentration_lower_threshold_label = ttk.Label(self.top, text="パターン認識用の閾値（下限） ", padding=(0,0), font=(15))
        self.cu_concentration_lower_threshold_label.grid(column=0, row=icuconcentrationlowerthreshold, sticky=Tk.E)
        self.cu_concentration_lower_threshold_value = Tk.DoubleVar(value = self.pt.cu_concentration_lower_threshold)
        self.cu_concentration_lower_threshold_entry = ttk.Entry(self.top, textvariable=self.cu_concentration_lower_threshold_value, width=self.pt.allinputlen0)
        self.cu_concentration_lower_threshold_entry.grid(column=1, row=icuconcentrationlowerthreshold, sticky=Tk.W)

        icuconcentrationhigherthreshold = icuconcentrationlowerthreshold+1
        self.cu_concentration_higher_threshold_label = ttk.Label(self.top, text="パターン認識用の閾値（上限） ", padding=(0,0), font=(15))
        self.cu_concentration_higher_threshold_label.grid(column=0, row=icuconcentrationhigherthreshold, sticky=Tk.E)
        self.cu_concentration_higher_threshold_value = Tk.DoubleVar(value = self.pt.cu_concentration_higher_threshold)
        self.cu_concentration_higher_threshold_entry = ttk.Entry(self.top, textvariable=self.cu_concentration_higher_threshold_value, width=self.pt.allinputlen0)
        self.cu_concentration_higher_threshold_entry.grid(column=1, row=icuconcentrationhigherthreshold, sticky=Tk.W)

        ibutton1 = icuconcentrationhigherthreshold+1
        self.plot_data_threshold_button = ttk.Button(self.top, text='データ+閾値プロット', command=self.plot_data_threshold)
        self.plot_data_threshold_button.grid(column=0, row=ibutton1, sticky=Tk.E)

        self.evaluation_button = ttk.Button(self.top, text='評価', command=self.uniformity_evaluation)
        self.evaluation_button.grid(column=1, row=ibutton1, sticky=Tk.W)

        iblank = ibutton1+1
        self.Region3 = ttk.Label(self.top, text="", padding=(0,0), font=(15))
        self.Region3.grid(column=0, row=iblank, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

        ioutputtitle = iblank+1
        self.output_title_label = ttk.Label(self.top, text="** 出力結果 **", padding=(0,0), font=(15))
        self.output_title_label.grid(column=0, row=ioutputtitle, sticky=Tk.W)

        imasscenterdistance = ioutputtitle+1
        self.mass_center_distance_label = ttk.Label(self.top, text="Cu-Alの重心の位置の距離 ", padding=(0,0), font=(15))
        self.mass_center_distance_label.grid(column=0, row=imasscenterdistance, sticky=Tk.E)
        self.mass_center_distance_value = Tk.DoubleVar(value = self.pt.mass_center_distance)
        self.mass_center_distance_entry = ttk.Entry(self.top, textvariable=self.mass_center_distance_value, width=self.pt.allinputlen0)
        self.mass_center_distance_entry.grid(column=1, row=imasscenterdistance, sticky=Tk.W)

        ivariance = imasscenterdistance+1
        self.variance_label = ttk.Label(self.top, text="分散 ", padding=(0,0), font=(15))
        self.variance_label.grid(column=0, row=ivariance, sticky=Tk.E)
        self.variance_value = Tk.DoubleVar(value = self.pt.variance)
        self.variance_entry = ttk.Entry(self.top, textvariable=self.variance_value, width=self.pt.allinputlen0)
        self.variance_entry.grid(column=1, row=ivariance, sticky=Tk.W)

        iprm4 = ivariance+1
        self.prm4_label = ttk.Label(self.top, text="Prm4 ", padding=(0,0), font=(15))
        self.prm4_label.grid(column=0, row=iprm4, sticky=Tk.E)
        self.prm4_value = Tk.DoubleVar(value = self.pt.Prm4)
        self.prm4_entry = ttk.Entry(self.top, textvariable=self.prm4_value, width=self.pt.allinputlen)
        self.prm4_entry.grid(column=1, row=iprm4, sticky=Tk.W)

        iblank = iprm4+1
        self.Region3 = ttk.Label(self.top, text="", padding=(0,0), font=(15))
        self.Region3.grid(column=0, row=iblank, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

        ibutton2 = iblank+1
        self.plot_mass_center_button = ttk.Button(self.top, text='Cu-Alの重心のプロット', command=self.plot_mass_center)
        self.plot_mass_center_button.grid(column=0, row=ibutton2, sticky=Tk.E)

        self.output_folder_open_button = ttk.Button(self.top, text='出力フォルダ開く', command=self.output_folder_open)
        self.output_folder_open_button.grid(column=1, row=ibutton2, sticky=Tk.W)

        return ibutton2+1
    
    def gui_refresh(self):
        '''Refresh GUI.'''

        ## create canvas
        self.canvas_create()
        ## clear all parameters (except Excel input path)
        self.pt.clearDict()
        ## create frame of Inputs input
        istart = self.all_inputs_frame_create(0)

    ##----------
    ## Actions
    ##----------
    def open_data_file(self):
        '''Open data file, read the numbers in the table.'''

        ## find ASCII folder path
        if self.pt.last_open_path is not None:
            self.pt.data_file_path = os.path.abspath( filedialog.askopenfilename(initialdir=self.pt.last_open_path) )
        else:
            self.pt.data_file_path = os.path.abspath( filedialog.askopenfilename(initialdir=os.getcwd()) )

        ## check whether there is an input
        if not self.pt.data_file_path.endswith('.xlsx'):
            self.data_file_path_value.set('')
            logoutput('', 'No Excel file input for execution. Please input an Excel file.', 'a', 'Error', True)
        else:
            ## record scene folder path as the last open path
            self.pt.last_open_path = os.path.dirname(self.pt.data_file_path)
            ## refresh GUI
            self.gui_refresh()

        ## make folder for output
        opedatetime = datetime.datetime.now()
        opedatetimestr = str(opedatetime.strftime('%Y%m%d_%H%M%S'))
        self.pt.output_folder_path = os.path.join( os.getcwd() , opedatetimestr+'_UniformityEvaluation_Outputs' )
        os.mkdir(self.pt.output_folder_path)
        ## make log file
        self.pt.log_file_path = os.path.join( self.pt.output_folder_path , opedatetimestr+'_UniformityEvaluation_Log.txt' )

        ## count the number of sheets in Excel file
        temp_file = pd.ExcelFile(self.pt.data_file_path)
        num_sheets = len(temp_file.sheet_names)
        logoutput(self.pt.log_file_path, f'The input file contains {num_sheets} sheets.', 'a', 'Report', False)

    def read_excel_data(self):
        '''Read excel data'''

        ## read sheet number
        self.pt.sheet_number = self.sheet_number_value.get()
        ## write log
        logoutput(self.pt.log_file_path, f'Reading data from input file from sheet #{self.pt.sheet_number}.', 'a', 'Report', False)
        ## read data
        self.df = pd.read_excel(self.pt.data_file_path, sheet_name=self.pt.sheet_number-1, header=None)  ## read assigned sheet
        ## fill NaN with 0
        self.df = self.df.fillna(0)
        
        ## convert df to numpy array
        self.df_np = self.df.values
        self.pt.df_np = self.df_np.copy()
        self.pt.df_np_copy = self.pt.df_np.copy()

        ## get number of column and row
        self.pt.height, self.pt.width = self.df_np.shape
        ## check whether data exists
        if (self.pt.height == 0) or (self.pt.width == 0):
            logoutput(self.pt.log_file_path, 'The sheet contains NO data. Please check the file.', 'a', 'Error', True)

        ## update GUI
        self.left_upper_corner_value.set('0, 0')
        self.right_lower_corner_value.set(f'{self.pt.width-1}, {self.pt.height-1}')
        self.width_height_show_value.set(f'{self.pt.width}, {self.pt.height}')
        ## write log
        logoutput(self.pt.log_file_path, 'Data range is updated in GUI.', 'a', 'Report', False)
        logoutput(self.pt.log_file_path, f'Original data size is {self.pt.width} x {self.pt.height}.', 'a', 'Report', False)

        ## print all possible stride / division combination
        self.all_stride_division_combination(self.pt.width, self.pt.height)

        ## refresh other inputs
        self.cut_injection_value.set(True)
        self.cu_concentration_lower_threshold_value.set(2)
        self.cu_concentration_higher_threshold_value.set(90)

        ## refresh outputs
        self.mass_center_distance_value.set(0)
        self.variance_value.set(0)
        self.prm4_value.set('0, 0, 0')

    def update_width_height_show(self):
        '''Update width and height by the input in left-upper and right-lower inputs.'''

        ## write log file
        logoutput(self.pt.log_file_path, 'Update width / height of data.', 'a', 'Report', False)

        ## get corner coordinate
        self.return_corner_coordinate()
        ## redefine width_height_show
        temp_x, temp_y = self.pt.x1-self.pt.x0+1, self.pt.y1-self.pt.y0+1
        self.width_height_show_value.set(f'{temp_x}, {temp_y}')

        ## print all possible stride / division combination
        self.all_stride_division_combination(temp_x, temp_y)

    def plot_data_division(self):
        '''Plot the data with region division grids.'''

        ## write log file
        logoutput(self.pt.log_file_path, 'Plot Cu concentration with divisions.', 'a', 'Report', False)

        ## get corner coordinate
        self.return_corner_coordinate()

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plot concentration
        fig, ax = plt.subplots()
        extent = (self.pt.x0-0.5, self.pt.x1+0.5, self.pt.y1+0.5, self.pt.y0-0.5)
        #im0 = ax.imshow(self.pt.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray_r')
        im0 = ax.imshow(self.pt.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')

        ## plot division grids
        self.pt.stride_x_pattern = self.stride_x_pattern_value.get()
        self.pt.stride_y_pattern = self.stride_y_pattern_value.get()
        self.pt.stride_combination = self.stride_combination_value.get()
        self.pt.stride_x_size, self.pt.stride_y_size = (int(value) for value in self.pt.stride_combination.replace(' ','').split(','))
        ## plot grid lines along X-direction
        xx = self.pt.x0-0.5
        while xx <= self.pt.x1+0.5:
            plt.plot([xx, xx], [self.pt.y0-0.5, self.pt.y1+0.5], color='red')
            xx += self.pt.stride_x_size
        ## plot grid lines along Y-direction
        yy = self.pt.y0-0.5
        while yy <= self.pt.y1+0.5:
            plt.plot([self.pt.x0-0.5, self.pt.x1+0.5], [yy, yy], color='red')
            yy += self.pt.stride_y_size
        ## plot ticks
        #plt.xticks(np.arange(self.pt.x0, self.pt.x1+1, self.pt.stride_x_size//2))
        #plt.yticks(np.arange(self.pt.y0, self.pt.y1+1, self.pt.stride_y_size//2))
        ## limit X/Y limits in 
        ax.set_xlim(self.pt.x0, self.pt.x1)
        ax.set_ylim(self.pt.y1, self.pt.y0)

        ## plot colorbar
        divider0 = make_axes_locatable(ax)
        cax0 = divider0.append_axes("right", size="4%", pad=0.2)
        fig.colorbar(im0, cax=cax0, ticks=levels)
        ## set title of plot
        ax.set_title(f"Concentration of Cu, sheet #{self.pt.sheet_number}, X stride = {self.pt.stride_x_size}, Y stride = {self.pt.stride_y_size}")

        imagefilename = f'stride_x_{self.pt.stride_x_size}_stride_y_{self.pt.stride_y_size}_grid_plot.png'
        plt.savefig(os.path.join(self.pt.output_folder_path, imagefilename))
        logoutput(self.pt.log_file_path, f'Image is saved as {imagefilename} in the output folder.', 'a', 'Report', False)
        if self.pt.plot_show_flag:
            plt.show()

    def plot_data_threshold(self):
        '''Plot the data with Cu concentration threshold.'''

        ## write log file
        logoutput(self.pt.log_file_path, f'Plot Cu concentration with thresholds.', 'a', 'Report', False)

        ## get corner coordinate
        self.return_corner_coordinate()

        ## remove injection in rasin part  <-- this part is manually executed
        self.pt.cut_injection_flag = self.cut_injection_value.get()
        if self.pt.cut_injection_flag:
            self.pt.df_np_copy[(self.pt.df_np_copy > 90) & (0 <= np.arange(self.pt.df_np_copy.shape[1])) & (np.arange(self.pt.df_np_copy.shape[1]) <= 100) & (np.arange(self.pt.df_np_copy.shape[0])[:, None] > 120)] = 0  ## np.nan
            self.df_np = self.pt.df_np_copy.copy()

        ## read stride size
        self.pt.stride_combination = self.stride_combination_value.get()
        self.pt.stride_x_size, self.pt.stride_y_size = (int(value) for value in self.pt.stride_combination.replace(' ','').split(','))

        ## find lower / upper contour positions
        self.lower_upper_contour_cal()

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plot concentration
        fig, ax = plt.subplots()
        extent = (self.pt.x0-0.5, self.pt.x1+0.5, self.pt.y1+0.5, self.pt.y0-0.5)
        #im0 = ax.imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray_r')
        im0 = ax.imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')

        plt.plot(self.lower_contour_x, self.lower_contour_y, color='red', lw=2)
        plt.plot(self.upper_contour_x, self.upper_contour_y, color='blue', lw=2)

        ## plot ticks
        #plt.xticks(np.arange(self.pt.x0, self.pt.x1+1, self.pt.stride_x_size//2))
        #plt.yticks(np.arange(self.pt.y0, self.pt.y1+1, self.pt.stride_y_size//2))
        ## limit X/Y limits in 
        ax.set_xlim(self.pt.x0, self.pt.x1)
        ax.set_ylim(self.pt.y1, self.pt.y0)

        ## plot colorbar
        divider0 = make_axes_locatable(ax)
        cax0 = divider0.append_axes("right", size="4%", pad=0.2)
        fig.colorbar(im0, cax=cax0, ticks=levels)
        ## set title of plot
        ax.set_title(f"Concentration of Cu, sheet #{self.pt.sheet_number}")


        imagefilename = f'Cu_concentration_threshold_{self.pt.cu_concentration_lower_threshold}-{self.pt.cu_concentration_higher_threshold}_plot.png'
        plt.savefig(os.path.join(self.pt.output_folder_path, imagefilename))
        logoutput(self.pt.log_file_path, f'Image is saved as {imagefilename} in the output folder.', 'a', 'Report', False)
        if self.pt.plot_show_flag:
            plt.show()

    def uniformity_evaluation(self):
        '''Evaluate the uniformity.'''

        ## write log file
        logoutput(self.pt.log_file_path, f'Start evaluation.', 'a', 'Report', False)

        ## get corner coordinate
        self.return_corner_coordinate()

        ## remove injection in rasin part  <-- this part is manually executed
        self.pt.cut_injection_flag = self.cut_injection_value.get()
        if self.pt.cut_injection_flag:
            self.pt.df_np_copy[(self.pt.df_np_copy > 90) & (0 <= np.arange(self.pt.df_np_copy.shape[1])) & (np.arange(self.pt.df_np_copy.shape[1]) <= 100) & (np.arange(self.pt.df_np_copy.shape[0])[:, None] > 120)] = 0  ## np.nan
            self.df_np = self.pt.df_np_copy.copy()

        ## find lower / upper contour positions
        self.lower_upper_contour_cal()
        ## cut contour boundary along Y-direction
        self.cut_along_Y_direction()
        ## cut contour boundary along X-direction
        self.cut_along_X_direction()
        
        ## calculate
        ## 1. distance between center of material
        ## 2. variance of all data
        self.center_distance_and_variance_whole_data()
        ## calculate Prm4 for each stride (X, Y)
        self.cal_Prm4()

        ## plot results
        self.plot_result_all()

        ## write log
        logoutput(self.pt.log_file_path, 'Evaluation Results (all):', 'a', 'Final', False)
        logoutput(self.pt.log_file_path, f'- Uniformity: Distance bewteen center of material = {self.pt.mass_center_distance}', 'a', 'Final', False)
        logoutput(self.pt.log_file_path, f'- Uniformity: Overall Variance = {self.pt.variance}', 'a', 'Final', False)
        logoutput(self.pt.log_file_path, f'- Uniformity: Prm4 value list:', 'a', 'Final', False)
        for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
            logoutput(self.pt.log_file_path, f'               - stride (X, Y) = ({stride_x:3}, {stride_y:3}): Prm4 = {self.Prm4_list[ii]}', 'a', 'Final', False)
        logoutput(self.pt.log_file_path, 'Evaluation of Uniformity is finished.', 'a', 'Final', True)

    def plot_mass_center(self):
        '''Plot mass center of Cu and Al.'''

        ## write log file
        logoutput(self.pt.log_file_path, f'Plot mass center of Cu / Al.', 'a', 'Report', False)

        ## plot concentration
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)

        fig, ax = plt.subplots()
        extent = (self.pt.x0-0.5, self.pt.x1+0.5, self.pt.y1+0.5, self.pt.y0-0.5)
        #im0 = ax.imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray_r')
        im0 = ax.imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
        ## limit X/Y limits in 
        ax.set_xlim(self.pt.x0, self.pt.x1)
        ax.set_ylim(self.pt.y1, self.pt.y0)
        ## plot colorbar
        divider0 = make_axes_locatable(ax)
        cax0 = divider0.append_axes("right", size="4%", pad=0.2)
        fig.colorbar(im0, cax=cax0, ticks=levels)
        ## plot mass center of Cu and Al
        ax.plot([self.df_np_Cu_xC],[self.df_np_Cu_yC],'x',color='red')
        ax.plot([self.df_np_Al_xC],[self.df_np_Al_yC],'x',color='blue')
        ax.text(self.df_np_Cu_xC,self.df_np_Cu_yC-20,'Center of Cu',size=10,color='red')
        ax.text(self.df_np_Al_xC,self.df_np_Al_yC+20,'Center of Al',size=10,color='blue')
        ## add title
        ax.set_title(f"Concentration of Cu, sheet #{self.pt.sheet_number}")

        imagefilename = 'Cu_Al_mass_center_distance_plot.png'
        plt.savefig(os.path.join(self.pt.output_folder_path, imagefilename))
        logoutput(self.pt.log_file_path, f'Image is saved as {imagefilename} in the output folder.', 'a', 'Report', False)
        if self.pt.plot_show_flag:
            plt.show()

    def output_folder_open(self):
        '''Open output folder.'''

        dirpath = Path(self.pt.output_folder_path)
        subprocess.Popen(["explorer", dirpath], shell=True)

    ##--------------
    ## Subordinary
    ##--------------
    def return_corner_coordinate(self):
        '''Read left-upper and right-lower coordinates.'''

        ## get left upper corner coordinate
        temp_string_pair = self.left_upper_corner_value.get().split(',')
        self.pt.x0, self.pt.y0 = int(temp_string_pair[0].strip()), int(temp_string_pair[1].strip())
        ## get right lower corner coordinate
        temp_string_pair = self.right_lower_corner_value.get().split(',')
        self.pt.x1, self.pt.y1 = int(temp_string_pair[0].strip()), int(temp_string_pair[1].strip())

    def common_divisors(self):
        '''Find all common divisors of self.pt.width and self.pt.height.'''

        # Step 1: Find the greatest common divisor (GCD) of a and b
        gcd_value = math.gcd(self.pt.width, self.pt.height)
        
        # Step 2: Find all divisors of the GCD
        divisors = []
        for i in range(1, gcd_value + 1):
            if gcd_value % i == 0:
                divisors.append(i)
        
        return divisors

    def all_divisors(self, length):
        '''Find all possible divisors for a given length.'''
        
        divisors = []
        for i in range(1, length+1):
            if length % i == 0:
                divisors.append(i)

        return divisors

    def all_stride_division_combination(self, width, height):
        '''Print all possible stride / division combination.'''

        ## find all possible stride patterns along X/Y-direction
        self.all_divisors_x_list = self.all_divisors(width)
        self.all_divisors_y_list = self.all_divisors(height)
        ## find all divisions along X/Y-direction
        self.all_division_x_list = [int(width / stride) for stride in self.all_divisors_x_list]
        self.all_division_y_list = [int(height / stride) for stride in self.all_divisors_y_list]
        ## combine the division set of X/Y, change it into a list
        division_x_set = set(self.all_division_x_list)
        division_y_set = set(self.all_division_y_list)
        division_xy_set = division_x_set.union(division_y_set)
        division_xy_list = list(division_xy_set)
        division_xy_list.sort()
        ## print all possible combination for division-stride
        logoutput(self.pt.log_file_path, f'All possible region divisions and strides along X/Y-direction:', 'a', 'Report', False)
        for division in division_xy_list:
            ## get correspondent stride along X-direction
            if division in self.all_division_x_list:
                div_index = self.all_division_x_list.index(division)
                stride_x_str = str(self.all_divisors_x_list[div_index])
            else:
                stride_x_str = ''
            ## get correspondent stride along Y-direction
            if division in self.all_division_y_list:
                div_index = self.all_division_y_list.index(division)
                stride_y_str = str(self.all_divisors_y_list[div_index])
            else:
                stride_y_str = ''
            ## print stride / division patterns
            logoutput(self.pt.log_file_path, f'  division = {division:4d}, stride X = {stride_x_str:>4}, stride Y = {stride_y_str:>4}.', 'a', 'Report', False)

        ## update stride patterns in GUI
        if len(self.all_divisors_x_list) >= 4:
            self.stride_x_pattern_value.set(', '.join([str(value) for value in self.all_divisors_x_list[-4:-1]]))
            self.pt.stride_x_size = self.all_divisors_x_list[-2]
        elif len(self.all_divisors_x_list) == 3:
            self.stride_x_pattern_value.set(', '.join([str(value) for value in self.all_divisors_x_list[-3:]]))
            self.pt.stride_x_size = self.all_divisors_x_list[-2]
        elif len(self.all_divisors_x_list) == 2:
            self.stride_pattern_value.set(', '.join([str(value) for value in self.all_divisors_x_list]))
            self.pt.stride_x_size = self.all_divisors_x_list[1]

        ## update stride patterns in GUI
        if len(self.all_divisors_y_list) >= 4:
            self.stride_y_pattern_value.set(', '.join([str(value) for value in self.all_divisors_y_list[-4:-1]]))
            self.pt.stride_y_size = self.all_divisors_y_list[-2]
        elif len(self.all_divisors_y_list) == 3:
            self.stride_y_pattern_value.set(', '.join([str(value) for value in self.all_divisors_y_list[-3:]]))
            self.pt.stride_y_size = self.all_divisors_y_list[-2]
        elif len(self.all_divisors_y_list) == 2:
            self.stride_pattern_value.set(', '.join([str(value) for value in self.all_divisors_y_list]))
            self.pt.stride_y_size = self.all_divisors_y_list[1]

        ## set stride combination for plotting
        self.stride_combination_value.set(str(self.pt.stride_x_size)+', '+str(self.pt.stride_y_size))

    def calculate_variance(self, data):
        non_NaN_num_count = np.count_nonzero(~np.isnan(data))
        non_NaN_num_average = np.nanmean(data)
        return np.nansum((data - non_NaN_num_average)**2) / non_NaN_num_count

    def contour_length(self,vertices):
        '''Function to calculate the length of a contour sub-path.'''

        diffs = np.diff(vertices, axis=0)
        return np.sum(np.sqrt((diffs ** 2).sum(axis=1)))  # Euclidean distance

    def find_longest_contour_line(self,contour):
        '''Function to find longest contour line.'''

        # Initialize variables to track the longest contour
        max_length = 0
        longest_contour_vertices = None

        # Iterate over all paths in the first collection
        #for path in contour.collections[0].get_paths():  <-- old version
        for path in contour.get_paths():
            vertices = path.vertices  # Get the vertices of the main path
            codes = path.codes        # Get the codes (used to separate sub-paths)

            # Separate sub-paths based on the MOVETO command (code 1)
            subpath_start = 0
            for i, code in enumerate(codes):
                if code == 1 and i > 0:  # MOVETO (start of a new sub-path)
                    subpath = vertices[subpath_start:i]
                    length = self.contour_length(subpath)
                    if length > max_length:
                        max_length = length
                        longest_contour_vertices = subpath
                    subpath_start = i

            # Handle the last sub-path
            subpath = vertices[subpath_start:]
            length = self.contour_length(subpath)
            if length > max_length:
                max_length = length
                longest_contour_vertices = subpath
        
        return longest_contour_vertices

    def lower_upper_contour_cal(self):

        ## write log file
        logoutput(self.pt.log_file_path, f'Find longest contour lines correspond to lower / upper threshold.', 'a', 'Report', False)

        ## get contour levels (Cu concentration threshold)
        self.pt.cu_concentration_lower_threshold = self.cu_concentration_lower_threshold_value.get()
        self.pt.cu_concentration_higher_threshold = self.cu_concentration_higher_threshold_value.get()

        ## get corner coordinate
        self.return_corner_coordinate()

        ## get coordinate
        self.x, self.y = np.arange(self.pt.width), np.arange(self.pt.height)
        self.X, self.Y = np.meshgrid(self.x, self.y)

        ## get lower / upper contours
        temp_fig = plt.figure()
        self.lower_contour = plt.contour(self.X, self.Y, self.df_np, levels=[self.pt.cu_concentration_lower_threshold])
        self.upper_contour = plt.contour(self.X, self.Y, self.df_np, levels=[self.pt.cu_concentration_higher_threshold])
        plt.close(temp_fig)

        ## find the longest contour for both upper and lower levels
        lower_longest_contour = self.find_longest_contour_line(self.lower_contour)
        upper_longest_contour = self.find_longest_contour_line(self.upper_contour)

        ## separate x and y for both contour levels
        self.lower_contour_x = np.array([int(round(value)) for value in lower_longest_contour[:,0]])
        self.lower_contour_y = np.array([int(round(value)) for value in lower_longest_contour[:,1]])
        self.upper_contour_x = np.array([int(round(value)) for value in upper_longest_contour[:,0]])
        self.upper_contour_y = np.array([int(round(value)) for value in upper_longest_contour[:,1]])

        ## check whether upper contour has discontinuity  <-- important in defining a closed region
        upper_sort_index = np.argsort(self.upper_contour_x)  ## sort upper contour by x
        self.upper_contour_x = self.upper_contour_x[upper_sort_index]
        self.upper_contour_y = self.upper_contour_y[upper_sort_index]

        upper_contour_x_copy = self.upper_contour_x.copy()
        for ii in range(len(upper_contour_x_copy)-1):
            x_gap = upper_contour_x_copy[ii+1] - upper_contour_x_copy[ii]
            if x_gap > 1:
                add_x_array = np.array(list(range(upper_contour_x_copy[ii], upper_contour_x_copy[ii+1])))
                add_y_array = np.array([self.upper_contour_y[ii]]*len(add_x_array))
                self.upper_contour_x = np.concatenate((self.upper_contour_x, add_x_array))
                self.upper_contour_y = np.concatenate((self.upper_contour_y, add_y_array))

        upper_sort_index = np.argsort(self.upper_contour_x)  ## sort upper contour by x again
        self.upper_contour_x = self.upper_contour_x[upper_sort_index]
        self.upper_contour_y = self.upper_contour_y[upper_sort_index]

    def cut_along_Y_direction(self):
        '''Cut contour along Y-direction.'''

        ## write log file
        logoutput(self.pt.log_file_path, 'Cut contour along Y-direction.', 'a', 'Report', False)
        
        ## build container for contour line
        yy_upper_container = []
        yy_lower_container = []
        for xx in range(self.pt.width):
            ii_lower_all = np.where(self.lower_contour_x == xx)
            ii_upper_all = np.where(self.upper_contour_x == xx)

            yy_lower_all = self.lower_contour_y[ii_lower_all]
            yy_upper_all = self.upper_contour_y[ii_upper_all]
            yy_upper_all.sort()
            yy_lower_all.sort()

            if len(yy_upper_container) > 0:
                yy_upper_container.append(list(yy_upper_all))
            else:
                yy_upper_container = [list(yy_upper_all)]

            if len(yy_lower_container) > 0:
                yy_lower_container.append(list(yy_lower_all))
            else:
                yy_lower_container = [list(yy_lower_all)]

        ## reset data above upper contour and below lowest lower contour
        for ii in range(self.pt.width):
            y_upper = yy_upper_container[ii][0]
            y_lower = yy_lower_container[ii][-1]
            self.df_np[0:y_upper+1,ii] = np.nan
            self.df_np[y_lower:,ii] = np.nan

        ## remove small residue
        for ii in range(self.pt.width):
            non_nan_count = np.count_nonzero(np.isnan(self.df_np[:,ii]))
            if non_nan_count > self.pt.height-6:  ##  <-- set manually
                self.df_np[:,ii] = np.nan

    def cut_along_X_direction(self):
        '''Cut contour along X-direction.'''

        ## write log file
        logoutput(self.pt.log_file_path, 'Cut contour along X-direction.', 'a', 'Report', False)

        x_mid = int(round(self.pt.width/2))

        ## cut on the left
        temp_fig = plt.figure()
        left_contour = plt.contour(self.X[:,0:x_mid], self.Y[:,0:x_mid], self.df_np[:,0:x_mid], levels=[self.pt.cu_concentration_lower_threshold])
        plt.close(temp_fig)
        ## find the longest contour for lower level
        left_longest_contour = self.find_longest_contour_line(left_contour)
        ## separate x and y for contour level
        left_contour_x = np.array([int(round(value)) for value in left_longest_contour[:,0]])
        left_contour_y = np.array([int(round(value)) for value in left_longest_contour[:,1]])

        if False:
            # Plot contours
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(left_contour_x, left_contour_y, color='red', lw=2)
            plt.title('cutted data, along X-direction (from left)')
            plt.show()

        ## build container for contour line
        xx_left_container = []
        for yy in range(self.pt.height):

            jj_left_all = np.where(left_contour_y == yy)
            xx_left_all = left_contour_x[jj_left_all]
            xx_left_all.sort()

            if len(xx_left_container) > 0:
                xx_left_container.append(list(xx_left_all))
            else:
                xx_left_container = [list(xx_left_all)]

            ## reset data below lowest lower contour
            if len(xx_left_container[yy]) > 0:
                x_left = xx_left_container[yy][-1]
                self.df_np[yy,0:x_left] = np.nan

        ## cut on the right
        temp_fig = plt.figure()
        right_contour = plt.contour(self.X[:,x_mid:], self.Y[:,x_mid:], self.df_np[:,x_mid:], levels=[self.pt.cu_concentration_lower_threshold])
        plt.close(temp_fig)
        ## find the longest contour for and lower level
        right_longest_contour = self.find_longest_contour_line(right_contour)
        ## separate x and y for contour level
        right_contour_x = np.array([int(round(value)) for value in right_longest_contour[:,0]])
        right_contour_y = np.array([int(round(value)) for value in right_longest_contour[:,1]])

        if False:
            # Plot contours
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(right_contour_x, right_contour_y, color='blue', lw=2)
            plt.title('cutted data, along X-direction (from right)')
            plt.show()

        ## build container for contour line
        xx_right_container = []
        for yy in range(self.pt.height):

            jj_right_all = np.where(right_contour_y == yy)
            xx_right_all = right_contour_x[jj_right_all]
            xx_right_all.sort()

            if len(xx_right_container) > 0:
                xx_right_container.append(list(xx_right_all))
            else:
                xx_right_container = [list(xx_right_all)]

            ## reset data below lowest lower contour
            if len(xx_right_container[yy]) > 0:
                x_right = xx_right_container[yy][0]
                self.df_np[yy,x_right:] = np.nan

    def center_distance_and_variance_whole_data(self):
        '''Calculate distance between material center and variance of whole data.'''

        ## split data into Cu and Al
        self.df_np_Cu = self.df_np[self.pt.y0:self.pt.y1+1, self.pt.x0:self.pt.x1+1].copy()
        self.df_np_Al = 100-self.df_np[self.pt.y0:self.pt.y1+1, self.pt.x0:self.pt.x1+1].copy()
        ## clip X/Y coordinate
        self.x_clip, self.y_clip = np.arange(self.pt.x0, self.pt.x1+1), np.arange(self.pt.y0, self.pt.y1+1)
        self.X_clip, self.Y_clip = np.meshgrid(self.x_clip, self.y_clip)

        ## calculate "center of material"
        self.df_np_Cu_xC = np.nansum(self.df_np_Cu * self.X_clip) / np.nansum(self.df_np_Cu)
        self.df_np_Cu_yC = np.nansum(self.df_np_Cu * self.Y_clip) / np.nansum(self.df_np_Cu)
        self.df_np_Al_xC = np.nansum(self.df_np_Al * self.X_clip) / np.nansum(self.df_np_Al)
        self.df_np_Al_yC = np.nansum(self.df_np_Al * self.Y_clip) / np.nansum(self.df_np_Al)

        ## calculate distance between "center of material"
        self.pt.mass_center_distance = np.sqrt( (self.df_np_Cu_xC-self.df_np_Al_xC)**2 + (self.df_np_Cu_yC-self.df_np_Al_yC)**2 )
        self.mass_center_distance_value.set(self.pt.mass_center_distance)
        ## calculate variance
        self.pt.variance = self.calculate_variance(self.df_np_Cu)
        self.variance_value.set(self.pt.variance)

    def cal_Prm4(self):
        '''Calculate Prm4 for each stride.'''

        ## get stride list in X/Y-direction
        self.stride_x_list = [int(value) for value in self.stride_x_pattern_value.get().replace(' ','').split(',')]
        self.stride_y_list = [int(value) for value in self.stride_y_pattern_value.get().replace(' ','').split(',')]

        ## check whether the length of these two list matches
        stride_x_list_len, stride_y_list_len = len(self.stride_x_list), len(self.stride_y_list)
        if stride_x_list_len != stride_y_list_len:
            logoutput(self.pt.log_file_path, 'The number of stride pattern does NOT match. Please check the input of stride pattern.', 'a', 'Error', True)

        ## calculate clipped width / height
        self.pt.width_clip, self.pt.height_clip = self.pt.x1-self.pt.x0+1, self.pt.y1-self.pt.y0+1

        ## remove the stride_x, stride_y = width_clip, height_clip
        if (self.pt.width_clip in self.stride_x_list) and (self.pt.height_clip in self.stride_y_list):
            self.stride_x_list.remove(self.pt.width_clip)
            self.stride_y_list.remove(self.pt.height_clip)
            ## reset length of stride pattern list
            stride_x_list_len, stride_y_list_len = len(self.stride_x_list), len(self.stride_y_list)

        self.Prm4_list = [0]*stride_x_list_len
        self.Prm4_plot_dict = {}
        for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
            logoutput(self.pt.log_file_path, f'Calculate evaluation for stride (X, Y) = {stride_x}, {stride_y}.', 'a', 'Report', False)
            ## calculate block number along each X/Y-direction
            block_num_x, block_num_y = int(self.pt.width_clip/stride_x), int(self.pt.height_clip/stride_y)
            ## save numpy data into list by block number = block_num_x * j + i
            room_Cu_average_list = [0]*int(block_num_x*block_num_y)
            room_plot_list = [False]*int(block_num_x*block_num_y)
            for j in range(block_num_y):
                for i in range(block_num_x):
                    ## room position in list
                    n = block_num_x * j + i
                    ## get numpy data of the nth room
                    jstart, jend = self.pt.y0+j*stride_y, self.pt.y0+(j+1)*stride_y
                    istart, iend = self.pt.x0+i*stride_x, self.pt.x0+(i+1)*stride_x
                    df_np_room = self.df_np[jstart:jend, istart:iend]
                    ## count number of non-NaN data in the nth room
                    df_np_room_data_count = np.count_nonzero(~np.isnan(df_np_room))
                    ## calculate average Cu concentration value in the nth room
                    if df_np_room_data_count == 0:
                        room_Cu_average = np.nan
                    else:
                        room_Cu_average = np.nanmean(df_np_room)
                    ## calculate ratio of non-NaN data in the nth room
                    room_valid_data_ratio = df_np_room_data_count / int(stride_x * stride_y) * 100
                    ## assign average Cu concentration value into the list
                    if room_valid_data_ratio > 10:
                        room_Cu_average_list[n] = room_Cu_average
                        room_plot_list[n] = True  ## will plot this region
                    else:
                        room_Cu_average_list[n] = np.nan
                        room_plot_list[n] = False ## will NOT plot this region
            ## convert average Cu list into numpy array
            room_Cu_average_np_array = np.array(room_Cu_average_list)
            ## calculate average Cu concentration across all rooms
            all_Cu_variance = self.calculate_variance(room_Cu_average_np_array)
            ## add variance to the list
            self.Prm4_list[ii] = all_Cu_variance
            ## add plotting information to the dictionary
            self.Prm4_plot_dict[(stride_x,stride_y)] = room_plot_list
        
        self.prm4_value.set(', '.join([str(value) for value in self.Prm4_list]))

    def plot_result_all(self):
        '''Plot all results.'''

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plotting extent size
        extent = (self.pt.x0-0.5, self.pt.x1+0.5, self.pt.y1+0.5, self.pt.y0-0.5)

        ## plot according to the number of stride pattern
        nstride = len(self.stride_x_list)
        ## number of plot in horizontal direction
        nplot = nstride+1
        ## plot concentration
        if nstride <= 2:
            nx, ny = nplot, 1
        elif nstride == 3:
            nx, ny = 2, 2
        elif nstride <= 5:
            nx, ny = 3, 2
        ## set containers
        fig_x, fig_y = 6.4*nx, 4.8*ny
        fig, axes = plt.subplots(ny, nx, figsize=(fig_x, fig_y))
        im, divider, cax = np.empty_like(axes), np.empty_like(axes), np.empty_like(axes)

        ## plot Cu concentration
        if nstride <= 2:
            #im0 = ax.imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray_r')
            im[0] = axes[0].imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
            ## plot mass center positions
            axes[0].plot([self.df_np_Cu_xC], [self.df_np_Cu_yC], 'x', color='red')
            axes[0].plot([self.df_np_Al_xC], [self.df_np_Al_yC], 'x', color='blue')
            axes[0].text(self.df_np_Cu_xC, self.df_np_Cu_yC-20, 'mass center of Cu', color='red')
            axes[0].text(self.df_np_Al_xC, self.df_np_Al_yC+20, 'mass center of Al', color='blue')
            ## set title of plot
            axes[0].set_title(f"Concentration of Cu, sheet #{self.pt.sheet_number}")
            ## plot colorbar
            divider[0] = make_axes_locatable(axes[0])
            cax[0] = divider[0].append_axes("right", size="4%", pad=0.2)
            fig.colorbar(im[0], cax=cax[0], ticks=levels)

            ## iterate the stride
            for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
                ## plot rooms over the Cu concentration
                im[ii+1] = axes[ii+1].imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
                ## plot grid lines along X-direction
                xx = self.pt.x0-0.5
                while xx <= self.pt.x1+0.5:
                    axes[ii+1].plot([xx, xx], [self.pt.y0-0.5, self.pt.y1+0.5], color='red')
                    xx += stride_x
                ## plot grid lines along Y-direction
                yy = self.pt.y0-0.5
                while yy <= self.pt.y1+0.5:
                    axes[ii+1].plot([self.pt.x0-0.5, self.pt.x1+0.5], [yy, yy], color='red')
                    yy += stride_y
                ## plot a square over the imshow if the room is NOT taken into calculation
                block_num_x = int(self.pt.width_clip/stride_x)
                for n, value in enumerate(self.Prm4_plot_dict[(stride_x, stride_y)]):
                    i, j = int(n%block_num_x), int(n//block_num_x)
                    if not value:
                        istart, jstart = self.pt.x0+i*stride_x, self.pt.y0+j*stride_y
                        square = patches.Rectangle((istart, jstart), stride_x, stride_y, facecolor='blue', alpha=0.2)
                        axes[ii+1].add_patch(square)

                ## set axe range
                axes[ii+1].set_xlim(self.pt.x0, self.pt.x1)
                axes[ii+1].set_ylim(self.pt.y1, self.pt.y0)
                ## set title of plot
                axes[ii+1].set_title(f"Concentration of Cu, sheet #{self.pt.sheet_number}, stride (X, Y) = {stride_x}, {stride_y}")
                ## plot colorbar
                divider[ii+1] = make_axes_locatable(axes[ii+1])
                cax[ii+1] = divider[ii+1].append_axes("right", size="4%", pad=0.2)
                fig.colorbar(im[ii+1], cax=cax[ii+1], ticks=levels)
        elif nstride <= 5:
            #im0 = ax.imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray_r')
            im[0,0] = axes[0,0].imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
            ## plot mass center positions
            axes[0,0].plot([self.df_np_Cu_xC], [self.df_np_Cu_yC], 'x', color='red')
            axes[0,0].plot([self.df_np_Al_xC], [self.df_np_Al_yC], 'x', color='blue')
            axes[0,0].text(self.df_np_Cu_xC, self.df_np_Cu_yC-20, 'mass center of Cu', color='red')
            axes[0,0].text(self.df_np_Al_xC, self.df_np_Al_yC+20, 'mass center of Al', color='blue')
            ## set title of plot
            axes[0,0].set_title(f"Concentration of Cu, sheet #{self.pt.sheet_number}")
            ## plot colorbar
            divider[0,0] = make_axes_locatable(axes[0,0])
            cax[0,0] = divider[0,0].append_axes("right", size="4%", pad=0.2)
            fig.colorbar(im[0,0], cax=cax[0,0], ticks=levels)

            ## iterate the stride
            for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
                ip = ii+1  ## index of plot including the overall plot
                ipi, ipj = int(ip%nx), int(ip//nx)
                ## plot rooms over the Cu concentration
                im[ipj,ipi] = axes[ipj,ipi].imshow(self.df_np[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
                ## plot grid lines along X-direction
                xx = self.pt.x0-0.5
                while xx <= self.pt.x1+0.5:
                    axes[ipj,ipi].plot([xx, xx], [self.pt.y0-0.5, self.pt.y1+0.5], color='red')
                    xx += stride_x
                ## plot grid lines along Y-direction
                yy = self.pt.y0-0.5
                while yy <= self.pt.y1+0.5:
                    axes[ipj,ipi].plot([self.pt.x0-0.5, self.pt.x1+0.5], [yy, yy], color='red')
                    yy += stride_y
                ## plot a square over the imshow if the room is NOT taken into calculation
                block_num_x = int(self.pt.width_clip/stride_x)
                for n, value in enumerate(self.Prm4_plot_dict[(stride_x, stride_y)]):
                    i, j = int(n%block_num_x), int(n//block_num_x)
                    if not value:
                        istart, jstart = self.pt.x0+i*stride_x, self.pt.y0+j*stride_y
                        square = patches.Rectangle((istart, jstart), stride_x, stride_y, facecolor='blue', alpha=0.2)
                        axes[ipj,ipi].add_patch(square)

                ## set axe range
                axes[ipj,ipi].set_xlim(self.pt.x0, self.pt.x1)
                axes[ipj,ipi].set_ylim(self.pt.y1, self.pt.y0)
                ## set title of plot
                axes[ipj,ipi].set_title(f"Concentration of Cu, sheet #{self.pt.sheet_number}, stride (X, Y) = {stride_x}, {stride_y}")
                ## plot colorbar
                divider[ipj,ipi] = make_axes_locatable(axes[ipj,ipi])
                cax[ipj,ipi] = divider[ipj,ipi].append_axes("right", size="4%", pad=0.2)
                fig.colorbar(im[ipj,ipi], cax=cax[ipj,ipi], ticks=levels)

                ## cover the last plot if nstride = 4
                if nstride == 4:
                    axes[1,2].axis('off')

        imagefilename = 'Cu_Al_mass_center_distance_Prm4_plot.png'
        plt.savefig(os.path.join(self.pt.output_folder_path, imagefilename))
        logoutput(self.pt.log_file_path, f'Image is saved as {imagefilename} in the output folder.', 'a', 'Report', False)
        if self.pt.plot_show_flag:
            plt.show()

################
## Parameters
################
class PwParaTransfer:
    def __init__(self):
        self.InitialCanvasSettings()
        self.InitialPathLength()
        self.clearDict()
        
        ## last open path and output folder path
        self.last_open_path = None
        self.output_folder_path = None
        self.log_file_path = None
        self.data_file_path = None

    def clearDict(self):
        ## input variables
        self.sheet_number = 1
        self.left_upper_corner = '0, 0'
        self.x0, self.y0 = 0, 0
        self.right_lower_corner = '99, 99'
        self.x1, self.y1 = 99, 99
        self.width_height_show = '100, 100'
        self.width, self.height = 100, 100
        self.width_clip, self.height_clip = 100, 100
        self.stride_x_pattern = '5, 10, 20'
        self.stride_y_pattern = '5, 10, 20'
        self.stride_combination = '10, 10'
        self.stride_x_size = 10
        self.stride_y_size = 10
        self.cut_injection_flag = True
        self.cu_concentration_lower_threshold = 2
        self.cu_concentration_higher_threshold = 90

        ## output result
        self.mass_center_distance = 0
        self.variance = 0
        self.Prm4 = '0, 0, 0'

        ## data container
        self.df_np = None
        self.df_np_copy = None

        ## plotting flag
        self.plot_show_flag = True

    def InitialCanvasSettings(self):
        ## initial canvas size
        self.canvasX0 = 800
        self.canvasY0 = 580

    def InitialPathLength(self):
        ## initial all inputs length
        self.allinputlen0 = 20
        self.allinputlen = 60

################
## Halt
################
class HaltException(Exception): pass
##-----------------------------------------


################
## Main
################
if __name__ == "__main__":
    root = Tk.Tk()
    app = UniformityEvaluation(master=root)
    app.mainloop()
