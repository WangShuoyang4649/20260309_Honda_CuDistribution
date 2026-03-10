import os
import sys
import datetime
from pathlib import Path

import numpy as np
import pandas as pd

import subprocess
import win32gui
import win32console

import matplotlib.pyplot as plt
from mpl_toolkits.axes_grid1 import make_axes_locatable

import tkinter as Tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

## Debug mode Handler
##   sys.tracebacklimit = 0: hide all traceback messages
##   sys.tracebacklimit = 1000: default value, show almost all traceback errors
#sys.tracebacklimit = 0

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

    ## Create canvas for all
    def canvas_create(self):
        '''Create canvas.'''

        ## start creating frame
        self.master.title('UniformityEvaluation')
        canvassize = str(self.pt.canvasX0)+'x'+str(self.pt.canvasY0)
        self.master.geometry(canvassize)

        ## create frame for base scene path input
        self.top = ttk.Frame(self.master, padding=10)
        self.top.grid(column=0, row=0, columnspan=4, rowspan=20, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

    ## create frame of scene list input
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
        self.data_file_path_button_open = ttk.Button(self.top, text='開く', command=self.open_data_file)
        self.data_file_path_button_open.grid(column=2, row=idatafilepathinput, sticky=Tk.W)

        itabnumber = idatafilepathinput+1
        self.tab_number_label = ttk.Label(self.top, text="タブ数 ", padding=(0,0), font=(15))
        self.tab_number_label.grid(column=0, row=itabnumber, sticky=Tk.E)
        self.tab_number_value = Tk.StringVar(value = self.pt.tab_number)
        self.tab_number_entry = ttk.Entry(self.top, textvariable=self.tab_number_value, width=self.pt.allinputlen0)
        self.tab_number_entry.grid(column=1, row=itabnumber, sticky=Tk.W)
        self.tab_number_entry.bind("<Return>", lambda event: self.read_excel_data())

        ileftuppercornerinput = itabnumber+1
        self.left_upper_corner_label = ttk.Label(self.top, text="データの左上のセル位置 ", padding=(0,0), font=(15))
        self.left_upper_corner_label.grid(column=0, row=ileftuppercornerinput, sticky=Tk.E)
        self.left_upper_corner_value = Tk.StringVar(value = self.pt.left_upper_corner)
        self.left_upper_corner_entry = ttk.Entry(self.top, textvariable=self.left_upper_corner_value, width=self.pt.allinputlen0)
        self.left_upper_corner_entry.grid(column=1, row=ileftuppercornerinput, sticky=Tk.W)
        self.left_upper_corner_entry.bind("<Return>", lambda event: self.update_width_height_show())

        irightlowercornerinput = ileftuppercornerinput+1
        self.right_lower_corner_label = ttk.Label(self.top, text="データの右下のセル位置 ", padding=(0,0), font=(15))
        self.right_lower_corner_label.grid(column=0, row=irightlowercornerinput, sticky=Tk.E)
        self.right_lower_corner_value = Tk.StringVar(value = self.pt.right_lower_corner)
        self.right_lower_corner_entry = ttk.Entry(self.top, textvariable=self.right_lower_corner_value, width=self.pt.allinputlen0)
        self.right_lower_corner_entry.grid(column=1, row=irightlowercornerinput, sticky=Tk.W)
        self.right_lower_corner_entry.bind("<Return>", lambda event: self.update_width_height_show())

        iwidthheightshow = irightlowercornerinput+1
        self.width_height_show_label = ttk.Label(self.top, text="データセル位置の範囲（幅 x 高さ） ", padding=(0,0), font=(15))
        self.width_height_show_label.grid(column=0, row=iwidthheightshow, sticky=Tk.E)
        self.width_height_show_value = Tk.StringVar(value = self.pt.width_height_show)
        self.width_height_show_entry = ttk.Entry(self.top, textvariable=self.width_height_show_value, width=self.pt.allinputlen0)
        self.width_height_show_entry.grid(column=1, row=iwidthheightshow, sticky=Tk.W)
        self.width_height_show_entry.config(state='disabled')

        iblank = iwidthheightshow+1
        self.Region1 = ttk.Label(self.top, text="", padding=(0,0), font=(15))
        self.Region1.grid(column=0, row=iblank, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

        ibutton0 = iblank+1
        self.plot_data_button = ttk.Button(self.top, text='データプロット', command=self.plot_data)
        self.plot_data_button.grid(column=0, row=ibutton0, sticky=Tk.E)

        self.evaluation_button = ttk.Button(self.top, text='評価', command=self.uniformity_evaluation)
        self.evaluation_button.grid(column=1, row=ibutton0, sticky=Tk.W)

        iblank = ibutton0+1
        self.Region2 = ttk.Label(self.top, text="", padding=(0,0), font=(15))
        self.Region2.grid(column=0, row=iblank, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

        ioutputtitle = iblank+1
        self.output_title_label = ttk.Label(self.top, text="** 出力結果 **", padding=(0,0), font=(15))
        self.output_title_label.grid(column=0, row=ioutputtitle, sticky=Tk.W)

        imasscenterdistance = ioutputtitle+1
        self.mass_center_distance_label = ttk.Label(self.top, text="Cu-Alの重心の位置の距離 ", padding=(0,0), font=(15))
        self.mass_center_distance_label.grid(column=0, row=imasscenterdistance, sticky=Tk.E)
        self.mass_center_distance_value = Tk.IntVar(value = self.pt.mass_center_distance)
        self.mass_center_distance_entry = ttk.Entry(self.top, textvariable=self.mass_center_distance_value, width=self.pt.allinputlen0)
        self.mass_center_distance_entry.grid(column=1, row=imasscenterdistance, sticky=Tk.W)

        ivariance = imasscenterdistance+1
        self.variance_label = ttk.Label(self.top, text="分散 ", padding=(0,0), font=(15))
        self.variance_label.grid(column=0, row=ivariance, sticky=Tk.E)
        self.variance_value = Tk.IntVar(value = self.pt.variance)
        self.variance_entry = ttk.Entry(self.top, textvariable=self.variance_value, width=self.pt.allinputlen0)
        self.variance_entry.grid(column=1, row=ivariance, sticky=Tk.W)

        iblank = ivariance+1
        self.Region3 = ttk.Label(self.top, text="", padding=(0,0), font=(15))
        self.Region3.grid(column=0, row=iblank, sticky=Tk.N+Tk.S+Tk.E+Tk.W)

        ibutton1 = iblank+1
        self.plot_mass_center_button = ttk.Button(self.top, text='Cu-Alの重心のプロット', command=self.plot_mass_center)
        self.plot_mass_center_button.grid(column=0, row=ibutton1, sticky=Tk.E)

        self.output_folder_open_button = ttk.Button(self.top, text='出力フォルダ開く', command=self.output_folder_open)
        self.output_folder_open_button.grid(column=1, row=ibutton1, sticky=Tk.W)

        return ibutton1+1

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
            ## refresh all inputs / outputs
            self.data_file_path_value.set('')
            self.left_upper_corner_value.set('0, 0')
            self.right_lower_corner_value.set('99, 99')
            self.width_height_show_value.set('100, 100')
        else:
            ## record scene folder path as the last open path
            self.pt.last_open_path = os.path.dirname(self.pt.data_file_path)
            ## set file path in GUI
            self.data_file_path_value.set(self.pt.data_file_path)
            ## refresh other inputs / outputs
            self.mass_center_distance_value.set('0')
            self.variance_value.set('0')

        ## make folder for output
        opedatetime = datetime.datetime.now()
        opedatetimestr = str(opedatetime.strftime('%Y%m%d_%H%M%S'))
        self.pt.output_folder_path = os.path.join( os.getcwd() , opedatetimestr+'_UniformityEvaluation_Outputs' )
        os.mkdir(self.pt.output_folder_path)
        ## make log file
        self.pt.log_file_path = os.path.join( self.pt.output_folder_path , opedatetimestr+'_UniformityEvaluation_Log.txt' )

        ## read data in the first tab
        self.read_excel_data()

    def read_excel_data(self):
        '''Read excel data'''

        ## read tab number
        self.pt.tab_number = int(self.tab_number_value.get())
        ## write log
        logoutput(self.pt.log_file_path, f'Reading data from input file from tab #{self.pt.tab_number}.', 'a', 'Report', False)
        ## read data
        self.df = pd.read_excel(self.pt.data_file_path, sheet_name=self.pt.tab_number-1, header=None)  ## read first sheet
        ## check whether NaN column exists in the file
        ## 1. drop columns with all NaN
        column_drop_flag = False
        self.df_copy = self.df.copy()
        for ii in range(len(self.df_copy.columns)-1, -1, -1):
            df_column_data = self.df_copy.iloc[:,ii].to_numpy()
            if np.all(np.isnan(df_column_data)):
                column_drop_flag = True
                self.df = self.df.drop(ii, axis=1)
        ## write log
        if column_drop_flag:
            self.df = self.df.set_axis(np.arange(len(self.df.iloc[0,:])), axis=1)
            logoutput(self.pt.log_file_path, 'Columns with all NaN are all deleted.', 'a', 'Report', False)
        ## 2. drop rows with all Nan
        row_drop_flag = False
        self.df_copy = self.df.copy()
        for jj in range(len(self.df_copy.iloc[:,0])-1, -1, -1):
            df_row_data = self.df_copy.iloc[jj,:].to_numpy()
            if np.all(np.isnan(df_row_data)):
                row_drop_flag = True
                self.df = self.df.drop(jj, axis=0)
        ## write log
        if row_drop_flag:
            self.df.reset_index(inplace=True, drop=True)  ## reset index
            logoutput(self.pt.log_file_path, 'Rows with all NaN are all deleted.', 'a', 'Report', False)
        
        ## convert df to numpy array
        self.np_data = self.df.values
        ## get number of column and row
        self.pt.height, self.pt.width = self.np_data.shape
        ## update GUI
        self.left_upper_corner_value.set('0, 0')
        self.right_lower_corner_value.set(f'{self.pt.width-1}, {self.pt.height-1}')
        self.width_height_show_value.set(f'{self.pt.width}, {self.pt.height}')
        ## write log
        logoutput(self.pt.log_file_path, 'Data range is updated in GUI.', 'a', 'Report', False)
        logoutput(self.pt.log_file_path, f'Original data size is {self.pt.width} x {self.pt.height}.', 'a', 'Report', False)

        ## refresh inputs / outputs
        self.mass_center_distance_value.set('0')
        self.variance_value.set('0')

    def update_width_height_show(self):
        '''Update width and height by the input in left-upper and right-lower inputs.'''

        ## get corner coordinate
        self.return_corner_coordinate()
        ## redefine width_height_show
        temp_x, temp_y = self.pt.x1-self.pt.x0+1, self.pt.y1-self.pt.y0+1
        self.width_height_show_value.set(f'{temp_x}, {temp_y}')

    def plot_data(self):
        '''Plot the data.'''

        ## get corner coordinate
        self.return_corner_coordinate()

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plot distribution
        fig, ax = plt.subplots()
        extent = (self.pt.x0-0.5, self.pt.x1+0.5, self.pt.y1+0.5, self.pt.y0-0.5)
        im0 = ax.imshow(self.np_data[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray_r')
        #im0 = ax.imshow(self.np_data[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
        plt.xticks(np.arange(self.pt.x0, self.pt.x1+1))
        plt.yticks(np.arange(self.pt.y0, self.pt.y1+1))
        ax.set_title(f"Concentration of Cu, tab #{self.pt.tab_number}")
        divider0 = make_axes_locatable(ax)
        cax0 = divider0.append_axes("right", size="4%", pad=0.2)
        fig.colorbar(im0, cax=cax0, ticks=levels)

        plt.savefig(os.path.join(self.pt.output_folder_path, 'original_plot.png'))

    def uniformity_evaluation(self):
        '''Evaluate the uniformity.'''

        ## get corner coordinate
        self.return_corner_coordinate()

        ## create coordinate
        x_data, y_data = np.arange(self.pt.x0, self.pt.x1+1), np.arange(self.pt.y0, self.pt.y1+1)
        ## generate mesh
        xx_data, yy_data = np.meshgrid(x_data, y_data)

        ## split data into Cu and Al
        self.np_data_Cu = self.np_data[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1].copy()
        self.np_data_Al = 100-self.np_data[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1].copy()

        ## calculate "center of material"
        self.np_data_Cu_xC = (self.np_data_Cu * xx_data).sum() / self.np_data_Cu.sum()
        self.np_data_Cu_yC = (self.np_data_Cu * yy_data).sum() / self.np_data_Cu.sum()
        self.np_data_Al_xC = (self.np_data_Al * xx_data).sum() / self.np_data_Al.sum()
        self.np_data_Al_yC = (self.np_data_Al * yy_data).sum() / self.np_data_Al.sum()

        ## calculate distance between "center of material"
        self.pt.mass_center_distance = np.sqrt( (self.np_data_Cu_xC-self.np_data_Al_xC)**2 + (self.np_data_Cu_yC-self.np_data_Al_yC)**2 )
        self.mass_center_distance_value.set(self.pt.mass_center_distance)
        ## calculate variance
        np_data_width = self.pt.x1+1-self.pt.x0
        np_data_height = self.pt.y1+1-self.pt.y0

        print(self.np_data_Cu.mean())
        self.pt.variance = ((self.np_data_Cu - self.np_data_Cu.mean())**2).sum() / (np_data_height*np_data_width)
        self.variance_value.set(self.pt.variance)

        ## write log
        logoutput(self.pt.log_file_path, 'Evaluation Results:', 'a', 'Final', False)
        logoutput(self.pt.log_file_path, f'- Uniformity: Distance bewteen center of material = {self.pt.mass_center_distance}', 'a', 'Final', False)
        logoutput(self.pt.log_file_path, f'- Uniformity: Variance = {self.pt.variance}', 'a', 'Final', False)
        logoutput(self.pt.log_file_path, 'Evaluation of Uniformity is finished.', 'a', 'Final', True)

    def plot_mass_center(self):
        '''Plot mass center of Cu and Al.'''

        ## plot distribution
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)

        fig, ax = plt.subplots()
        extent = (self.pt.x0-0.5, self.pt.x1+0.5, self.pt.y1+0.5, self.pt.y0-0.5)
        im0 = ax.imshow(self.np_data[self.pt.y0:self.pt.y1+1,self.pt.x0:self.pt.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray_r')
        plt.xticks(np.arange(self.pt.x0, self.pt.x1+1))
        plt.yticks(np.arange(self.pt.y0, self.pt.y1+1))
        divider0 = make_axes_locatable(ax)
        cax0 = divider0.append_axes("right", size="4%", pad=0.2)
        ax.plot([self.np_data_Cu_xC],[self.np_data_Cu_yC],'x',color='red')
        ax.plot([self.np_data_Al_xC],[self.np_data_Al_yC],'x',color='blue')
        ax.set_title(f"Concentration of Cu, tab #{self.pt.tab_number}")
        ax.text(self.np_data_Cu_xC,self.np_data_Cu_yC-1,'Center of Cu',size=10,color='red')
        ax.text(self.np_data_Al_xC,self.np_data_Al_yC+1,'Center of Al',size=10,color='blue')
        fig.colorbar(im0, cax=cax0, ticks=levels)

        plt.savefig(os.path.join(self.pt.output_folder_path, 'mass_center_distance.png'))

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

    def calculate_variance(self, data):
        try:
            width, height = data.shape
            return np.sqrt( ((data - data.mean())**2).sum() / (height*width) )
        except ValueError:
            width = data.shape
            return np.sqrt( ((data - data.mean())**2).sum() / width )

        
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

    def clearDict(self):
        ## input variables
        self.data_file_path = ''
        self.tab_number = 1
        self.left_upper_corner = '0, 0'
        self.x0, self.y0 = 0, 0
        self.right_lower_corner = '99, 99'
        self.x1, self.y1 = 99, 99
        self.width_height_show = '100, 100'
        self.width, self.height = 100, 100

        ## output result
        self.mass_center_distance = 0
        self.variance = 0

    def InitialCanvasSettings(self):
        ## initial canvas size
        self.canvasX0 = 800
        self.canvasY0 = 400

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
