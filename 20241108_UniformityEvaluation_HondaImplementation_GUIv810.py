import os
import datetime
from pathlib import Path

import matplotlib
matplotlib.use('module://pwpy.plot_backend')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import matplotlib.patches as patches
from mpl_toolkits.axes_grid1 import make_axes_locatable


import numpy as np
import pandas as pd

import pwpy as PW

class ContourCut:
    def __init__(self, session, scene):
        '''Initialize class.'''

        self.session = session
        self.scene = scene

    def run(self):
        '''Run main function.'''

        if self.ask_inputs():
            ## prepare image data
            self.prepare_image_data()
            
            ## cut image by contour levels
            self.cut_along_Y_direction()
            self.cut_along_X_direction()
            self.show_compare_image_plot()

            ## import data as air data
            self.create_csv_file()
            self.import_air()

            ## evaluate uniformity
            self.uniformity_evaluation()
            self.cal_Prm4()
            ## plot result
            self.plot_result_all()

            ## write results
            self.write_results()

            ## adjust camera
            self.adjust_camera()

    def ask_inputs(self):
        '''Ask inputs of 1) data path, 2) sheet number 3) lower threshold 4) upper threshold.'''

        a = PW.Input()
        self.xlsx_path = a.add_file('Data file path (xlsx)')
        self.sheet_num = a.add_integer('Sheet number', 0)
        self.upper_left_point = a.add_string('Clip (upper left point)', '0, 0')
        self.lower_right_point = a.add_string('Clip (lower right point)', '511, 383')
        self.stride_pattern_x = a.add_string('Division stride (X-direction)', '64, 128, 256')
        self.stride_pattern_y = a.add_string('Division stride (Y-direction)', '48, 96, 192')
        self.cut_injection = a.add_boolean('Cut injection', True)
        self.contour_level0 = a.add_float('Cutting lower threshold', 2)
        self.contour_level1 = a.add_float('Cutting upper threshold', 90)

        if not a.ask():
            return False
        return True

    def prepare_image_data(self):
        '''Prepare image data.'''

        ## create output folder
        datetimestr = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder_name = f"{datetimestr}_CuDistribution_Output"
        self.output_folder_path = os.path.join(self.scene.path(PW.SCENE_PATH_root_dir), output_folder_name)
        os.mkdir(self.output_folder_path)
        print(f'** REPORT ** Create output folder path {self.output_folder_path}.')

        print('** REPORT ** Prepare image data.')

        ## convert file path and sheet number
        self.xlsx_path = Path(self.xlsx_path.get())
        self.sheet_num = self.sheet_num.get()

        ## read file
        self.df = pd.read_excel(self.xlsx_path, sheet_name=self.sheet_num, header=None)  ## read first sheet
        ## fill NaN with 0
        self.df = self.df.fillna(0)
        ## convert dataframe into numpy array
        self.df_np = self.df.values
        ## save original data
        self.df_np_original = self.df_np.copy()

        ## get width / height of image
        self.height, self.width = self.df_np.shape
        ## check whether the data is of standard
        if (self.height != 384) or (self.width != 512):
            print(f'** ERROR ** The size of the input data is {self.width}x{self.height}, which is NOT a standard size (512x384).')
            print(f'** ERROR ** Please check the input again.')
            exit()

        ## get clipping area
        self.get_clipping_area()

        ## print all possible stride / division combination
        self.all_stride_division_combination(self.width_clip, self.height_clip)

        ## get stride list in X/Y-direction and check whether the strides are divisible to clipped width / height
        self.stride_check()

        ## cut off injection parts
        self.cut_injection = self.cut_injection.get()
        if self.cut_injection:
            ## remove injection in rasin part  <-- this part is manually executed
            self.df_np[(self.df_np > 90) & (0 <= np.arange(self.df_np.shape[1])) & (np.arange(self.df_np.shape[1]) <= 100) & (np.arange(self.df_np.shape[0])[:, None] > 120)] = 0

        ## get coordinate
        self.x, self.y = np.arange(self.width), np.arange(self.height)
        self.X, self.Y = np.meshgrid(self.x, self.y)

    def get_clipping_area(self):
        '''Get clipping area and check whether it is beyond the data region.'''

        self.upper_left_point = self.upper_left_point.get().replace(' ','').split(',')
        self.lower_right_point = self.lower_right_point.get().replace(' ','').split(',')
        self.x0, self.y0 = int(self.upper_left_point[0]), int(self.upper_left_point[1])
        self.x1, self.y1 = int(self.lower_right_point[0]), int(self.lower_right_point[1])
        ## check whether clipping area is inside the whole data region
        if (self.x0 >= self.width) or (self.x1 >= self.width) or (self.y0 >= self.height) or (self.y1 >= self.height):
            print(f'** ERROR ** The clipping area is beyong the original data region.')
            print(f'** ERROR ** Please adjust the clipping area to the standard size (512x384).')
            exit()
        ## find clipping area size
        self.width_clip, self.height_clip = self.x1-self.x0+1, self.y1-self.y0+1
        print(f'** REPORT ** The input data size is {self.width_clip}x{self.height_clip}.')

    def stride_check(self):
        '''Get stride list in X/Y-direction and check whether the strides are divisible to clipped width / height'''

        ## get stride list in X/Y-direction
        self.stride_x_list = [int(value) for value in self.stride_pattern_x.get().replace(' ','').split(',')]
        self.stride_y_list = [int(value) for value in self.stride_pattern_y.get().replace(' ','').split(',')]

        ## check whether the length of these two list matches
        self.stride_x_list_len, self.stride_y_list_len = len(self.stride_x_list), len(self.stride_y_list)
        if self.stride_x_list_len != self.stride_y_list_len:
            print('** ERROR ** The number of stride pattern does NOT match. Please check the input of stride pattern.')
            exit()
        
        ## check whether each stride is divisible by clipped width / height
        for stride_x in self.stride_x_list:
            if self.width_clip % stride_x != 0:
                print(f'** ERROR ** The stride {stride_x} along X-direction is NOT divisible to the clipped width.')
                print(f'** ERROR ** Please check the division lists above and input an appropriate stride.')
                exit()
        for stride_y in self.stride_y_list:
            if self.height_clip % stride_y != 0:
                print(f'** ERROR ** The stride {stride_y} along Y-direction is NOT divisible to the clipped height.')
                print(f'** ERROR ** Please check the division lists above and input an appropriate stride.')
                exit()

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
        print(f'** REPORT ** All possible region divisions and strides along X/Y-direction:')
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
            print(f'** REPORT **   division = {division:4d}, stride X = {stride_x_str:>4}, stride Y = {stride_y_str:>4}.')

    def all_divisors(self, length):
        '''Find all possible divisors for a given length.'''
        
        divisors = []
        for i in range(1, length+1):
            if length % i == 0:
                divisors.append(i)

        return divisors

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
        for path in contour.collections[0].get_paths():
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

    def cut_along_Y_direction(self):
        '''Cut contour along Y-direction.'''

        print('** REPORT ** Cut contour along Y-direction.')

        self.contour_level0 = self.contour_level0.get()
        self.contour_level1 = self.contour_level1.get()

        self.lower_contour = plt.contour(self.X, self.Y, self.df_np, levels=[self.contour_level0])
        self.upper_contour = plt.contour(self.X, self.Y, self.df_np, levels=[self.contour_level1])

        ## find the longest contour for both upper and lower levels
        lower_longest_contour = self.find_longest_contour_line(self.lower_contour)
        upper_longest_contour = self.find_longest_contour_line(self.upper_contour)

        ## separate x and y for both contour levels
        lower_contour_x = np.array([int(round(value)) for value in lower_longest_contour[:,0]])
        lower_contour_y = np.array([int(round(value)) for value in lower_longest_contour[:,1]])
        upper_contour_x = np.array([int(round(value)) for value in upper_longest_contour[:,0]])
        upper_contour_y = np.array([int(round(value)) for value in upper_longest_contour[:,1]])

        ## check whether upper contour has discontinuity  <-- important in defining a closed region
        upper_sort_index = np.argsort(upper_contour_x)  ## sort upper contour by x
        upper_contour_x = upper_contour_x[upper_sort_index]
        upper_contour_y = upper_contour_y[upper_sort_index]

        upper_contour_x_copy = upper_contour_x.copy()
        for ii in range(len(upper_contour_x_copy)-1):
            x_gap = upper_contour_x_copy[ii+1] - upper_contour_x_copy[ii]
            if x_gap > 1:
                add_x_array = np.array(list(range(upper_contour_x_copy[ii], upper_contour_x_copy[ii+1])))
                add_y_array = np.array([upper_contour_y[ii]]*len(add_x_array))
                upper_contour_x = np.concatenate((upper_contour_x, add_x_array))
                upper_contour_y = np.concatenate((upper_contour_y, add_y_array))

        upper_sort_index = np.argsort(upper_contour_x)  ## sort upper contour by x again
        upper_contour_x = upper_contour_x[upper_sort_index]
        upper_contour_y = upper_contour_y[upper_sort_index]

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(lower_contour_x, lower_contour_y, color='red', lw=2)
            plt.plot(upper_contour_x, upper_contour_y, color='blue', lw=2)
            plt.show()
            plt.close(f)

        ## build container for contour line
        yy_upper_container = []
        yy_lower_container = []
        for xx in range(self.width):
            ii_lower_all = np.where(lower_contour_x == xx)
            ii_upper_all = np.where(upper_contour_x == xx)

            yy_lower_all = lower_contour_y[ii_lower_all]
            yy_upper_all = upper_contour_y[ii_upper_all]
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
        for ii in range(self.width):
            y_upper = yy_upper_container[ii][0]
            y_lower = yy_lower_container[ii][-1]
            self.df_np[0:y_upper+1,ii] = np.nan
            self.df_np[y_lower:,ii] = np.nan

        ## remove small residue
        for ii in range(self.width):
            non_nan_count = np.count_nonzero(np.isnan(self.df_np[:,ii]))
            if non_nan_count > self.height-6:  ##  <-- set manually
                self.df_np[:,ii] = np.nan

        if False:
            # Plot plot after cutting along Y-direction
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.show()
            plt.close(f)

    def cut_along_X_direction(self):
        '''Cut contour along X-direction.'''

        print('** REPORT ** Cut contour along X-direction.')

        x_mid = int(round(self.width/2))
        self.df_np_copy = self.df_np.copy()

        ## cut on the left
        left_contour = plt.contour(self.X[:,0:x_mid], self.Y[:,0:x_mid], self.df_np[:,0:x_mid], levels=[self.contour_level0])
        ## find the longest contour for lower level
        left_longest_contour = self.find_longest_contour_line(left_contour)
        ## separate x and y for contour level
        left_contour_x = np.array([int(round(value)) for value in left_longest_contour[:,0]])
        left_contour_y = np.array([int(round(value)) for value in left_longest_contour[:,1]])

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(left_contour_x, left_contour_y, color='red', lw=2)
            plt.title('cutted data, along X-direction (from left)')
            plt.show()
            plt.close(f)

        ## build container for contour line
        xx_left_container = []
        for yy in range(self.height):

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
        right_contour = plt.contour(self.X[:,x_mid:], self.Y[:,x_mid:], self.df_np[:,x_mid:], levels=[self.contour_level0])
        ## find the longest contour for and lower level
        right_longest_contour = self.find_longest_contour_line(right_contour)
        ## separate x and y for contour level
        right_contour_x = np.array([int(round(value)) for value in right_longest_contour[:,0]])
        right_contour_y = np.array([int(round(value)) for value in right_longest_contour[:,1]])

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(right_contour_x, right_contour_y, color='blue', lw=2)
            plt.title('cutted data, along X-direction (from right)')
            plt.show()
            plt.close(f)

        ## build container for contour line
        xx_right_container = []
        for yy in range(self.height):

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

    def show_compare_image_plot(self):
        '''Show image (before/after cut) comparison.'''

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plotting extent size
        extent0 = (-0.5, self.width-0.5, self.height-0.5, -0.5)
        extent = (self.x0-0.5, self.x1+0.5, self.y1+0.5, self.y0-0.5)

        ## plot data frame into image
        plt.clf()
        fig, axes = plt.subplots(1, 2, figsize=(10, 3))
        im, divider, cax = np.empty_like(axes), np.empty_like(axes), np.empty_like(axes)

        im[0] = axes[0].imshow(self.df_np_original, cmap='gray', extent=extent0, vmin=vmin, vmax=vmax, aspect='auto')
        im[1] = axes[1].imshow(self.df_np, cmap='gray', extent=extent, vmin=vmin, vmax=vmax, aspect="auto")

        axes[0].set_xlim([0,self.width-1])
        axes[0].set_ylim([self.height-1,0])
        axes[1].set_xlim([self.x0,self.x1])
        axes[1].set_ylim([self.y1,self.y0])

        axes[0].set_title(f'Concentration of Cu, sheet #{self.sheet_num}, original data')
        axes[1].set_title(f'Concentration of Cu, sheet #{self.sheet_num}, cutted data')

        ## plot colorbar
        for ii in range(2):
            divider[ii] = make_axes_locatable(axes[ii])
            cax[ii] = divider[ii].append_axes("right", size="4%", pad=0.2)
            fig.colorbar(im[ii], cax=cax[ii], ticks=levels)

        imagepath = os.path.join(self.output_folder_path, 'contour_cut_compare_image_plot.png')
        plt.savefig(imagepath)
        print(f'** REPORT ** Output comparing image to {imagepath}.')

        plt.show()
        plt.close(fig)

    def create_csv_file(self):
        '''Create CSV file from image data.'''

        ## total number of data
        df_len = int( self.width * self.height )

        ## replot data frames
        x_coords, y_coords = np.indices((self.width, self.height))
        x_flat, y_flat, values_flat = x_coords.flatten(), y_coords.flatten(), self.df_np.transpose().flatten()
        df_new = pd.DataFrame({
            'x': x_flat,
            'y': y_flat,
            'z': [0]*df_len,
            'vx': values_flat,
            'vy': values_flat,
            'vz': values_flat,
        })

        ## drop data beyond clipping area
        df_new = df_new[df_new['x'] >= self.x0]
        df_new = df_new[df_new['x'] <= self.x1]
        df_new = df_new[df_new['y'] >= self.y0]
        df_new = df_new[df_new['y'] <= self.y1]

        ## drop NaN of df_new
        df_new.dropna(axis=0, how='any', inplace=True)

        ## get length of df_new
        self.df_new_length = len(df_new.index)

        ## find min / max of X / Y in cutted contour
        self.x_min_cut, self.x_max_cut = df_new['x'].min(), df_new['x'].max()
        self.y_min_cut, self.y_max_cut = df_new['y'].min(), df_new['y'].max()
        print('** REPORT ** Data is prepared into data frame.')

        ## output data into csv file
        output_data_path = os.path.join(self.output_folder_path, 'Cu_data.csv')
        df_new.to_csv(output_data_path, index=True, header=False)
        ## make list data
        output_list_path = os.path.join(self.output_folder_path, 'list_data.csv')
        with open(output_list_path, 'w') as outputfile:
            outputfile.write('0,Cu_data.csv\n')

        print('** REPORT ** Data is exported into CSV files.')

    def import_air(self):
        '''Import air node with created data.'''

        ## create air parameter
        p = self.scene.create_parameters(PW.PARAMETERS_import_air)
        p['filePath'] = os.path.join(self.output_folder_path, 'list_data.csv')
        ## submit task and wait
        q = self.session.task_queue
        t = q.submit(self.scene, p)
        q.wait(-1)

        ## change display of particle to show color
        for node in self.scene.nodes:
            if node.node_type == PW.NODE_air:
                node['material.useColorMap'] = True
                break
        self.scene.write()

        print('** REPORT ** Air node is imported into scene.')

    def uniformity_evaluation(self):
        '''Evaluate uniformity of mixing.'''

        print('** REPORT ** Start overall uniformity evaluation.')

        ## split data into Cu and Al
        self.df_np_Cu = self.df_np[self.y0:self.y1+1, self.x0:self.x1+1].copy()
        self.df_np_Al = 100-self.df_np[self.y0:self.y1+1, self.x0:self.x1+1].copy()
        ## clip X/Y coordinate
        self.x_clip, self.y_clip = np.arange(self.x0, self.x1+1), np.arange(self.y0, self.y1+1)
        self.X_clip, self.Y_clip = np.meshgrid(self.x_clip, self.y_clip)

        ## calculate "center of material"
        self.np_data_Cu_xC = np.nansum(self.df_np_Cu * self.X_clip) / np.nansum(self.df_np_Cu)
        self.np_data_Cu_yC = np.nansum(self.df_np_Cu * self.Y_clip) / np.nansum(self.df_np_Cu)
        self.np_data_Al_xC = np.nansum(self.df_np_Al * self.X_clip) / np.nansum(self.df_np_Al)
        self.np_data_Al_yC = np.nansum(self.df_np_Al * self.Y_clip) / np.nansum(self.df_np_Al)

        ## calculate distance between "center of material"
        self.mass_center_distance = np.sqrt( (self.np_data_Cu_xC-self.np_data_Al_xC)**2 + (self.np_data_Cu_yC-self.np_data_Al_yC)**2 )

        ## add probe point as center of material Cu
        self.scene.create_node(PW.NODE_probe_point)
        self.scene.write()
        for node in self.scene.nodes:
            if node.node_type == PW.NODE_probe_point:
                node.name = 'Cu_mass_center'
                node['probe.point.x'] = self.np_data_Cu_xC
                node['probe.point.y'] = self.np_data_Cu_yC
                node['probe.point.z'] = 0
        
        ## add probe point as center of material Al
        self.scene.create_node(PW.NODE_probe_point)
        self.scene.write()
        for node in self.scene.nodes:
            if (node.node_type == PW.NODE_probe_point) and (node.name != 'Cu_mass_center'):
                node.name = 'Al_mass_center'
                node['probe.point.x'] = self.np_data_Al_xC
                node['probe.point.y'] = self.np_data_Al_yC
                node['probe.point.z'] = 0

        self.variance = np.nansum((self.df_np_Cu - np.nanmean(self.df_np_Cu))**2) / self.df_new_length

    def cal_Prm4(self):
        '''Calculate Prm4 for each stride.'''

        print('** REPORT ** Start Prm4 evaluation.')

        ## remove the stride_x, stride_y = width_clip, height_clip
        if (self.width_clip in self.stride_x_list) and (self.height_clip in self.stride_y_list):
            self.stride_x_list.remove(self.width_clip)
            self.stride_y_list.remove(self.height_clip)
            ## reset length of stride pattern list
            self.stride_x_list_len, self.stride_y_list_len = len(self.stride_x_list), len(self.stride_y_list)

        self.Prm4_list = [0]*self.stride_x_list_len
        self.Prm4_plot_dict = {}
        for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
            print(f'** REPORT ** Calculate evaluation for stride (X, Y) = {stride_x}, {stride_y}.')
            ## calculate block number along each X/Y-direction
            block_num_x, block_num_y = int(self.width_clip/stride_x), int(self.height_clip/stride_y)
            ## save numpy data into list by block number = block_num_x * j + i
            room_Cu_average_list = [0]*int(block_num_x*block_num_y)
            room_plot_list = [False]*int(block_num_x*block_num_y)
            for j in range(block_num_y):
                for i in range(block_num_x):
                    ## room position in list
                    n = block_num_x * j + i
                    ## get numpy data of the nth room
                    jstart, jend = self.y0+j*stride_y, self.y0+(j+1)*stride_y
                    istart, iend = self.x0+i*stride_x, self.x0+(i+1)*stride_x
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
        
    def calculate_variance(self, data):
        non_NaN_num_count = np.count_nonzero(~np.isnan(data))
        non_NaN_num_average = np.nanmean(data)
        return np.nansum((data - non_NaN_num_average)**2) / non_NaN_num_count

    def plot_result_all(self):
        '''Plot all results.'''

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plotting extent size
        extent = (self.x0-0.5, self.x1+0.5, self.y1+0.5, self.y0-0.5)

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
        plt.clf()
        fig_x, fig_y = 6*nx, 4*ny
        #fig_x, fig_y = 6.4*nx, 4.8*ny
        fig, axes = plt.subplots(ny, nx, figsize=(fig_x, fig_y))
        im, divider, cax = np.empty_like(axes), np.empty_like(axes), np.empty_like(axes)

        ## plot Cu concentration
        if nstride <= 2:
            im[0] = axes[0].imshow(self.df_np[self.y0:self.y1+1,self.x0:self.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
            ## plot mass center positions
            axes[0].plot([self.np_data_Cu_xC], [self.np_data_Cu_yC], 'x', color='red')
            axes[0].plot([self.np_data_Al_xC], [self.np_data_Al_yC], 'x', color='blue')
            axes[0].text(self.np_data_Cu_xC, self.np_data_Cu_yC-20, 'mass center of Cu', color='red')
            axes[0].text(self.np_data_Al_xC, self.np_data_Al_yC+20, 'mass center of Al', color='blue')
            ## set title of plot
            axes[0].set_title(f"Concentration of Cu, sheet #{self.sheet_num}")
            ## plot colorbar
            divider[0] = make_axes_locatable(axes[0])
            cax[0] = divider[0].append_axes("right", size="4%", pad=0.2)
            fig.colorbar(im[0], cax=cax[0], ticks=levels)

            ## iterate the stride
            for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
                ## plot rooms over the Cu concentration
                im[ii+1] = axes[ii+1].imshow(self.df_np[self.y0:self.y1+1,self.x0:self.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
                ## plot grid lines along X-direction
                xx = self.x0-0.5
                while xx <= self.x1+0.5:
                    axes[ii+1].plot([xx, xx], [self.y0-0.5, self.y1+0.5], color='red')
                    xx += stride_x
                ## plot grid lines along Y-direction
                yy = self.y0-0.5
                while yy <= self.y1+0.5:
                    axes[ii+1].plot([self.x0-0.5, self.x1+0.5], [yy, yy], color='red')
                    yy += stride_y
                ## plot a square over the imshow if the room is NOT taken into calculation
                block_num_x = int(self.width_clip/stride_x)
                for n, value in enumerate(self.Prm4_plot_dict[(stride_x, stride_y)]):
                    i, j = int(n%block_num_x), int(n//block_num_x)
                    if not value:
                        istart, jstart = self.x0+i*stride_x, self.y0+j*stride_y
                        square = patches.Rectangle((istart, jstart), stride_x, stride_y, facecolor='blue', alpha=0.2)
                        axes[ii+1].add_patch(square)

                ## set axe range
                axes[ii+1].set_xlim(self.x0, self.x1)
                axes[ii+1].set_ylim(self.y1, self.y0)
                ## set title of plot
                axes[ii+1].set_title(f"Concentration of Cu, sheet #{self.sheet_num}, stride (X, Y) = {stride_x}, {stride_y}")
                ## plot colorbar
                divider[ii+1] = make_axes_locatable(axes[ii+1])
                cax[ii+1] = divider[ii+1].append_axes("right", size="4%", pad=0.2)
                fig.colorbar(im[ii+1], cax=cax[ii+1], ticks=levels)
        elif nstride <= 5:
            im[0,0] = axes[0,0].imshow(self.df_np[self.y0:self.y1+1,self.x0:self.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
            ## plot mass center positions
            axes[0,0].plot([self.np_data_Cu_xC], [self.np_data_Cu_yC], 'x', color='red')
            axes[0,0].plot([self.np_data_Al_xC], [self.np_data_Al_yC], 'x', color='blue')
            axes[0,0].text(self.np_data_Cu_xC, self.np_data_Cu_yC-20, 'mass center of Cu', color='red')
            axes[0,0].text(self.np_data_Al_xC, self.np_data_Al_yC+20, 'mass center of Al', color='blue')
            ## set title of plot
            axes[0,0].set_title(f"Concentration of Cu, sheet #{self.sheet_num}")
            ## plot colorbar
            divider[0,0] = make_axes_locatable(axes[0,0])
            cax[0,0] = divider[0,0].append_axes("right", size="4%", pad=0.2)
            fig.colorbar(im[0,0], cax=cax[0,0], ticks=levels)

            ## iterate the stride
            for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
                ip = ii+1  ## index of plot including the overall plot
                ipi, ipj = int(ip%nx), int(ip//nx)
                ## plot rooms over the Cu concentration
                im[ipj,ipi] = axes[ipj,ipi].imshow(self.df_np[self.y0:self.y1+1,self.x0:self.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
                ## plot grid lines along X-direction
                xx = self.x0-0.5
                while xx <= self.x1+0.5:
                    axes[ipj,ipi].plot([xx, xx], [self.y0-0.5, self.y1+0.5], color='red')
                    xx += stride_x
                ## plot grid lines along Y-direction
                yy = self.y0-0.5
                while yy <= self.y1+0.5:
                    axes[ipj,ipi].plot([self.x0-0.5, self.x1+0.5], [yy, yy], color='red')
                    yy += stride_y
                ## plot a square over the imshow if the room is NOT taken into calculation
                block_num_x = int(self.width_clip/stride_x)
                for n, value in enumerate(self.Prm4_plot_dict[(stride_x, stride_y)]):
                    i, j = int(n%block_num_x), int(n//block_num_x)
                    if not value:
                        istart, jstart = self.x0+i*stride_x, self.y0+j*stride_y
                        square = patches.Rectangle((istart, jstart), stride_x, stride_y, facecolor='blue', alpha=0.2)
                        axes[ipj,ipi].add_patch(square)

                ## set axe range
                axes[ipj,ipi].set_xlim(self.x0, self.x1)
                axes[ipj,ipi].set_ylim(self.y1, self.y0)
                ## set title of plot
                axes[ipj,ipi].set_title(f"Concentration of Cu, sheet #{self.sheet_num}, stride (X, Y) = {stride_x}, {stride_y}")
                ## plot colorbar
                divider[ipj,ipi] = make_axes_locatable(axes[ipj,ipi])
                cax[ipj,ipi] = divider[ipj,ipi].append_axes("right", size="4%", pad=0.2)
                fig.colorbar(im[ipj,ipi], cax=cax[ipj,ipi], ticks=levels)

                ## cover the last plot if nstride = 4
                if nstride == 4:
                    axes[1,2].axis('off')

        imagepath = os.path.join(self.output_folder_path, 'Cu_Al_mass_center_distance_Prm4_plot.png')
        plt.savefig(imagepath)
        print(f'** REPORT ** Output result image to {imagepath}.')

        plt.show()
        plt.close(fig)

    def write_results(self):
        '''Write all results into file and print.'''

        outputfilename = 'Uniformity_Evaluation_Results.txt'
        outputfilepath = os.path.join(self.output_folder_path, outputfilename)
        with open(outputfilepath, 'w') as file:
            file.write(f'Evaluation Results (all):\n')
            file.write(f'- Uniformity: Distance bewteen center of material = {self.mass_center_distance}\n')
            file.write(f'- Uniformity: Overall Variance = {self.variance}\n')
            file.write(f'- Uniformity: Prm4 value list:\n')
            for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
                file.write(f'               - stride (X, Y) = ({stride_x:3}, {stride_y:3}): Prm4 = {self.Prm4_list[ii]}\n')

        print(f'** FINAL ** Evaluation Results (all):')
        print(f'** FINAL ** - Uniformity: Distance bewteen center of material = {self.mass_center_distance}')
        print(f'** FINAL ** - Uniformity: Overall Variance = {self.variance}')
        print(f'** FINAL ** - Uniformity: Prm4 value list:')
        for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
            print(f'               - stride (X, Y) = ({stride_x:3}, {stride_y:3}): Prm4 = {self.Prm4_list[ii]}')
        print(f'** FINAL ** Evaluation of Uniformity is finished.')
        print(f'** FINAL ** Results are exported to {outputfilepath}.')

    def adjust_camera(self):
        '''Adjust camera for better view.'''

        xc = int(round((self.x_max_cut+self.x_min_cut)/2)) - 0.5
        yc = int(round((self.y_max_cut+self.y_min_cut)/2)) - 0.5

        ## change ortho of camera
        for node in self.scene.nodes:
            if node.name == 'camera':
                node['ortho'] = True
                node['transform.location'] = (xc, yc, 0)
                node['transform.rotation'] = (180, 0, 0)
                node['gaze'] = max(xc,yc)+1
                break
        
        ## change visibility of domain
        for node in self.scene.nodes:
            if node.name == 'domain':
                node['visible'] = False
        
        ## save scene settings
        self.scene.write()

def main():
    session = PW.Session()
    scene   = session.active_scene
    if scene != None:
        cc = ContourCut(session, scene)
        cc.run()

if 'Particleworks' in __name__:
    main()

