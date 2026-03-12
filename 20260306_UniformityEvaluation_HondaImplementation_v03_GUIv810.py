"""
Uniformity Evaluation - Honda Cu Distribution
Version: v03 GUI v810
Date: 2026-03-12
"""

import os
import time
import datetime
from pathlib import Path

import matplotlib
matplotlib.use('module://pwpy.plot_backend')
import matplotlib.pyplot as plt
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

        ## dx/dy threshold at the contact boundary, set manually
        self.dx_threshold = 6
        self.dy_threshold = 6

    def run(self):
        '''Run main function.'''

        if self.ask_inputs():
            ## prepare image data
            self.prepare_image_data()
            
            ## cut image by contour levels
            self.cut_feature = self.cut_feature.get()
            if self.cut_feature:
                self.record_contour_line_coordinate()
                self.cut_along_Y_direction()
                if self.cut_leftright.get():
                    self.cut_along_X_direction()
                    self.feature_cleansing()
                ## show compare image plotss
                self.show_compare_image_plot()

            ## import data as air data
            self.create_csv_file()
            self.import_air()

            ## evaluate uniformity
            self.uniformity_evaluation()
            self.cal_Prm4()

            ## evaluate slits
            self.xslit_flag = self.xslit_evaluate.get()
            if self.xslit_flag:
                self.slit_calc("x")
            self.yslit_flag = self.yslit_evaluate.get()
            if self.yslit_flag:
                self.slit_calc("y")

            ## plot result
            self.plot_result_all()
            ## plot result (color)
            self.plot_result_color()

            ## write results
            self.write_results()

            ## adjust camera
            self.adjust_camera()

    def ask_inputs(self):
        '''Ask inputs of 
            1) input file
            2) image clipping
            3) room division
            4) feature cut
            5) evaluate slit for X/Y
        '''

        ## clear existing nodes
        self.clear_nodes()

        ## ask inputs
        a_data_input = PW.Input()
        self.xlsx_path = a_data_input.add_file('Data file path (xlsx)')
        self.sheet_num = a_data_input.add_integer('Sheet number', 0)

        if not a_data_input.ask():
            return False
        else:
            ## get data size
            self.read_file_data_size()
            ## plot original data
            self.plot_original_data()
            time.sleep(3)  ## show the plot for 3 seconds

            ## get stride combination
            self.all_stride_division_combination(self.width, self.height, True)

            ## get stride lists:
            def get_stride_lists(dlist):
                if len(dlist) > 4:
                    lshow = dlist[1:4]
                elif len(dlist) == 4:
                    lshow = dlist[1:3]
                elif len(dlist) <= 3:
                    lshow = dlist[1]
                return [str(value) for value in lshow]

            ## get stride lists: X-direction
            x_stride_list_show = get_stride_lists(self.all_division_x_list)
            ## get stride lists: Y-direction
            y_stride_list_show = get_stride_lists(self.all_division_y_list)

            a_setting = PW.Input()
            self.upper_left_point = a_setting.add_string('Clip (upper left point)', '0, 0')
            self.lower_right_point = a_setting.add_string('Clip (lower right point)', f'{self.width-1}, {self.height-1}')
            self.stride_pattern_x = a_setting.add_string('Division stride (X-direction)', ', '.join(x_stride_list_show))
            self.stride_pattern_y = a_setting.add_string('Division stride (Y-direction)', ', '.join(y_stride_list_show))
            self.cut_injection = a_setting.add_boolean('Cut injection', True)
            self.cut_leftright = a_setting.add_boolean('Cut Left/Right', True)
            self.cut_feature = a_setting.add_boolean('Cut feature by thresholds', True)
            self.contour_level0 = a_setting.add_float('Cutting lower threshold', 2)
            self.contour_level1 = a_setting.add_float('Cutting upper threshold', 90)
            self.room_valid_data_ratio_threshold = a_setting.add_float('Filling rate threshold (%)', 10)
            self.xslit_evaluate = a_setting.add_boolean('Evaluate slits (X)', True)
            self.xslit_num = a_setting.add_integer('  1a) Slit number (X)', 10)
            self.xslit_csvfile_path = a_setting.add_file('  1b) Slit positions (X, csv)')
            self.xslit_width = a_setting.add_integer('  2) Slit width (X)', 25)
            self.yslit_evaluate = a_setting.add_boolean('Evaluate slits (Y)', False)
            self.yslit_num = a_setting.add_integer('  1a) Slit number (Y)', 10)
            self.yslit_csvfile_path = a_setting.add_file('  1b) Slit positions (Y, csv)')
            self.yslit_width = a_setting.add_integer('  2) Slit width (Y)', 25)

            if not a_setting.ask():
                return False
            return True

    def read_file_data_size(self):
        '''Prepare image data.'''

        print('** REPORT ** Prepare image data.')

        ## convert file path and sheet number
        self.xlsx_path = Path(self.xlsx_path.get())
        self.sheet_num = self.sheet_num.get()
        self.xlsx_name = Path(self.xlsx_path).stem

        ## read file
        print(f'** REPORT ** Read data in {self.xlsx_path}, sheet #{self.sheet_num}.')
        self.df = pd.read_excel(self.xlsx_path, sheet_name=self.sheet_num, header=None)
        ## fill NaN with 0
        self.df = self.df.fillna(0)
        ## convert dataframe into numpy array
        self.df_np = self.df.values
        ## save original data
        self.df_np_original = self.df_np.copy()

        ## get width / height of image
        self.height, self.width = self.df_np.shape

        ## create output folder
        datetimestr = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder_name = f"{datetimestr}_CuDistribution_Output_{self.xlsx_name}_{self.sheet_num}"
        self.output_folder_path = os.path.join(self.scene.path(PW.SCENE_PATH_root_dir), output_folder_name)
        os.mkdir(self.output_folder_path)
        print(f'** REPORT ** Create output folder path {self.output_folder_path}.')

    def plot_original_data(self):
        '''Show original image.'''

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plotting extent size
        extent0 = (-0.5, self.width-0.5, self.height-0.5, -0.5)

        ## plot data frame into image
        plt.clf()
        fig, axe = plt.subplots(1, 1, figsize=(5, 3))
        im, divider, cax = np.empty_like(axe), np.empty_like(axe), np.empty_like(axe)

        im = axe.imshow(self.df_np_original, cmap='gray', extent=extent0, vmin=vmin, vmax=vmax, aspect='auto')

        axe.set_xlim([-0.5,self.width-0.5])
        axe.set_ylim([self.height-0.5,-0.5])
        axe.set_aspect('equal', 'box')

        axe.set_title(f'Concentration of Cu, sheet #{self.sheet_num}, original data')

        ## plot colorbar
        divider = make_axes_locatable(axe)
        cax = divider.append_axes("right", size="4%", pad=0.2)
        fig.colorbar(im, cax=cax, ticks=levels)

        plt.show()
        plt.close(fig)

    def clear_nodes(self):
        '''Clear existing Air node and Point Probe nodes.'''

        for node in self.scene.nodes:
            if (node.node_type == PW.NODE_air) or (node.node_type == PW.NODE_probe_point):
                self.scene.delete_object(node)

    def prepare_image_data(self):
        '''Prepare data by 1) calculating divisible stride size
                           2) removing injection
        '''

        ## get clipping area
        self.get_clipping_area()

        ## print all possible stride / division combination
        if ((self.width == self.width_clip) and (self.height == self.height_clip)):
            self.all_stride_division_combination(self.width_clip, self.height_clip, False)
        else:
            self.all_stride_division_combination(self.width_clip, self.height_clip, True)
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
        def output_caution(stride_s):
            print(f'** CAUTION ** The stride {stride_s} along X-direction is NOT divisible to the clipped width.')
            print(f'** CAUTION ** Please check the division lists above and input an appropriate stride.')
            print(f'** CAUTION ** Otherwise, the data outside of the clipped region will be used for Prm4 calculation.')
        for stride_x in self.stride_x_list:
            if self.width_clip % stride_x != 0:
                output_caution(stride_x)
        for stride_y in self.stride_y_list:
            if self.height_clip % stride_y != 0:
                output_caution(stride_y)

    def all_stride_division_combination(self, width, height, print_flag):
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
        ## print all possible combination for division-stride if data is clipped
        if print_flag:
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
        else:
            pass

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

    def record_contour_line_coordinate(self):
        '''Record contour positions.'''

        self.contour_level0_value = self.contour_level0.get()
        self.contour_level1_value = self.contour_level1.get()

        self.lower_contour = plt.contour(self.X, self.Y, self.df_np, levels=[self.contour_level0_value])
        self.upper_contour = plt.contour(self.X, self.Y, self.df_np, levels=[self.contour_level1_value])

        ## find the longest contour for both upper and lower levels
        self.lower_longest_contour = self.find_longest_contour_line(self.lower_contour)
        self.upper_longest_contour = self.find_longest_contour_line(self.upper_contour)

        ## separate x and y for both contour levels
        self.lower_contour_x = np.array([int(round(value)) for value in self.lower_longest_contour[:,0]])
        self.lower_contour_y = np.array([int(round(value)) for value in self.lower_longest_contour[:,1]])
        self.upper_contour_x = np.array([int(round(value)) for value in self.upper_longest_contour[:,0]])
        self.upper_contour_y = np.array([int(round(value)) for value in self.upper_longest_contour[:,1]])

        ## build dataframe for operation convenience
        lower_contour_df = pd.DataFrame({"x": self.lower_contour_x, "y": self.lower_contour_y})
        upper_contour_df = pd.DataFrame({"x": self.upper_contour_x, "y": self.upper_contour_y})
        ## remove redundant rows, due to round() execution above
        lower_contour_df = lower_contour_df.drop_duplicates()
        upper_contour_df = upper_contour_df.drop_duplicates()
        ## seperate dataframe into numpy arrays again
        self.lower_contour_x = lower_contour_df["x"].to_numpy()
        self.lower_contour_y = lower_contour_df["y"].to_numpy()
        self.upper_contour_x = upper_contour_df["x"].to_numpy()
        self.upper_contour_y = upper_contour_df["y"].to_numpy()

    def get_marginal_contour_position(self, contour_m, contour_n, sFlag, pFlag):
        '''Get marginal contour positions, which is applicable for either orientation.
           Combination of sFlag and pFlag will change the calculation purpose,
           such as getting the outer-most or inner-most position (X or Y) of the contour position.'''

        idx = np.lexsort((sFlag * contour_m, contour_n))  ## primary sort key = contour_n, secondary sort key = contour_m
        x_sorted = contour_n[idx]  ## reorganize contour_n by the sorted indices
        ## keep only the first point for each unique contour_n after sorting
        first = np.r_[True, x_sorted[1:] != x_sorted[:-1]]
        keep_idx = idx[first]
        
        if pFlag:
            return contour_m[keep_idx], contour_n[keep_idx]
        else:
            return contour_n[keep_idx], contour_m[keep_idx]

    def feature_cleansing(self):
        '''Feature cleansing.'''

        ## reset NaN to -1
        self.df_np_nan_replace = np.nan_to_num(self.df_np, nan=-1)
        ## get feature >= 0
        self.zero_contour = plt.contour(self.X, self.Y, self.df_np_nan_replace, levels=[0])
        ## find the longest contour for 0 levels
        self.zero_longest_contour = self.find_longest_contour_line(self.zero_contour)
        ## separate x and y for 0 contour level
        self.zero_contour_x = np.array([int(round(value)) for value in self.zero_longest_contour[:,0]])
        self.zero_contour_y = np.array([int(round(value)) for value in self.zero_longest_contour[:,1]])

        ## build dataframe for operation convenience
        zero_contour_df = pd.DataFrame({"x": self.zero_contour_x, "y": self.zero_contour_y})
        ## remove redundant rows, due to round() execution above
        zero_contour_df = zero_contour_df.drop_duplicates()
        ## seperate dataframe into numpy arrays again
        self.zero_contour_x = zero_contour_df["x"].to_numpy()
        self.zero_contour_y = zero_contour_df["y"].to_numpy()

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(self.zero_contour_x, self.zero_contour_y, color='blue', lw=2)
            plt.title("cutted data + NaN-contour")
            plt.show()
            plt.close(f)

        # Create coordinate grid
        y, x = np.mgrid[:self.height, :self.width]
        points = np.vstack((x.ravel(), y.ravel())).T
        # Create polygon path
        poly = matplotlib.path.Path(np.column_stack((self.zero_contour_x, self.zero_contour_y)))
        # Check which pixels are inside
        mask = poly.contains_points(points)
        mask = mask.reshape(self.height, self.width)
        # Apply mask, get main feature
        self.df_np[~mask] = np.nan  # or 0

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.title("final result")
            plt.show()
            plt.close(f)

    def cut_along_Y_direction(self):
        '''Cut contour along Y-direction.'''

        print('** REPORT ** Cut contour along Y-direction.')

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

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(self.upper_contour_x, self.upper_contour_y, color='blue', lw=2)
            plt.plot(self.lower_contour_x, self.lower_contour_y, color='red', lw=2)
            plt.title("original data + longest contours")
            plt.show()
            plt.close(f)

        ## extract highest contour X/Y-position for upper contour
        north_contour_x, north_contour_y = self.get_marginal_contour_position(self.upper_contour_y, self.upper_contour_x, 1, False)
        ## extract lowest contour X/Y-position for lower contour
        south_contour_x, south_contour_y = self.get_marginal_contour_position(self.lower_contour_y, self.lower_contour_x, -1, False)

        if False:
            ## Plot contours with range of main feature along X-direction
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(north_contour_x, north_contour_y, color='blue', lw=2)
            plt.plot(south_contour_x, south_contour_y, color='red', lw=2)
            plt.title('original data + longest contour (wide)')
            plt.show()
            plt.close(f)

        ## reset data above upper contour and below lowest lower contour
        for ii in range(self.width):
            y_upper = north_contour_y[ii]
            y_lower = south_contour_y[ii]
            self.df_np[0:y_upper+1,ii] = np.nan
            self.df_np[y_lower:,ii] = np.nan

        if False:
            # Plot plot after cutting along Y-direction
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.title("cutted data (along Y-direction)")
            plt.show()
            plt.close(f)

    def cut_along_X_direction(self):
        '''Cut contour along X-direction.'''

        ## find lowest / highest contour X/Y positions

        print('** REPORT ** Cut contour along X-direction.')

        ## extract lowest contour X/Y-positions of upper contour
        north_contour_x, north_contour_y = self.get_marginal_contour_position(self.upper_contour_y, self.upper_contour_x, -1, False)
        ## extract highest contour X/Y-positions of lower contour
        south_contour_x, south_contour_y = self.get_marginal_contour_position(self.lower_contour_y, self.lower_contour_x, 1, False)

        if False:
            ## Plot contours with range of main feature along X-direction
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(north_contour_x, north_contour_y, color='blue', lw=2)
            plt.plot(south_contour_x, south_contour_y, color='red', lw=2)
            plt.title('cutted data + longest contour (narrow)')
            plt.show()
            plt.close(f)

        ## find main feature boundary along X-direction
        xleft, xright = 0, self.width-1
        ## find boundary of feature from left
        for ii in range(self.width):
            dy = south_contour_y[ii] - north_contour_y[ii]
            if dy > self.dy_threshold:
                xleft = ii
                break
        ## find boundary of feature (at separator line) from right
        for ii in range(self.width-1, -1, -1):
            dy = south_contour_y[ii] - north_contour_y[ii]
            if dy > self.dy_threshold:
                xright = ii
                break

        if False:
            ## Plot contours with range of main feature along X-direction
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(north_contour_x, north_contour_y, color='blue', lw=2)
            plt.plot(south_contour_x, south_contour_y, color='red', lw=2)
            plt.plot([xleft,xleft],[self.y0,self.y1], color="purple", lw=2, ls="--")
            plt.plot([xright,xright],[self.y0,self.y1], color="purple", lw=2, ls="--")
            plt.title('cutted data, range of main features (X)')
            plt.show()
            plt.close(f)

        ## find contours on left and right that will be removed from the original contour positions
        south_contour_x_cut = np.r_[south_contour_x[:xleft], south_contour_x[xright:]]
        south_contour_y_cut = np.r_[south_contour_y[:xleft], south_contour_y[xright:]]

        ## extract main feature range
        ## build dataframe for operation convenience
        south_contour_df = pd.DataFrame({"x": south_contour_x_cut, "y": south_contour_y_cut})
        lower_contour_df = pd.DataFrame({"x": self.lower_contour_x, "y": self.lower_contour_y})
        ## remove cut parts from the original contour positions to focus on main feature region
        lower_contour_df_leftover = (lower_contour_df.merge(south_contour_df, on=["x", "y"], how="left", indicator=True)
                                     .query('_merge == "left_only"').drop(columns="_merge"))
        ## remove isolated contour parts, keep the contour continuous in the main feature region
        dlen = np.inf
        while dlen > 0:
            df_len_org = len(lower_contour_df_leftover)
            lower_contour_df_leftover_squash = lower_contour_df_leftover.sort_values(by=["x","y"]).drop_duplicates(subset="x", keep="first")
            group_id = (lower_contour_df_leftover_squash["x"].diff() != 1).cumsum()
            lower_contour_df_discontinuous = lower_contour_df_leftover_squash.groupby(group_id).filter(lambda g: len(g) <= self.dx_threshold)
            lower_contour_df_leftover = (lower_contour_df_leftover.merge(lower_contour_df_discontinuous, on=["x", "y"], how="left", indicator=True)
                                        .query('_merge == "left_only"').drop(columns="_merge"))
            dlen = df_len_org - len(lower_contour_df_leftover)

        ## seperate df into numpy arrays
        lower_contour_x_leftover = lower_contour_df_leftover["x"].to_numpy()
        lower_contour_y_leftover = lower_contour_df_leftover["y"].to_numpy()
        
        if False:
            ## Plot contours with range of main feature along X-direction
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(lower_contour_x_leftover, lower_contour_y_leftover, color='red', lw=2)
            #plt.plot(lower_contour_x_leftover, lower_contour_y_leftover, color='red', lw=2, ls="None", marker="o")
            plt.plot([xleft,xleft],[self.y0,self.y1], color="purple", lw=2, ls="--")
            plt.plot([xright,xright],[self.y0,self.y1], color="purple", lw=2, ls="--")
            plt.title('cutted data, lower contour in main features (X)')
            plt.show()
            plt.close(f)

        ## find largest y-position and its x-position
        ymax_idx = np.argmax(lower_contour_y_leftover)
        ## find left / right contour X/Y-positions
        west_contour_x = lower_contour_x_leftover[ymax_idx:]
        west_contour_y = lower_contour_y_leftover[ymax_idx:]
        east_contour_x = lower_contour_x_leftover[:ymax_idx+1]
        east_contour_y = lower_contour_y_leftover[:ymax_idx+1]

        if False:
            ## Plot contours on the left / right
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(west_contour_x, west_contour_y, color='red', lw=2)
            plt.plot(east_contour_x, east_contour_y, color='blue', lw=2)
            plt.plot([lower_contour_x_leftover[ymax_idx]],[lower_contour_y_leftover[ymax_idx]], color="purple", lw=2, ls="None", marker="o")
            plt.title('cutted data, seperate contour (left/right)')
            plt.show()
            plt.close(f)

        ## extract only the outmost position of X for left side
        left_contour_x, left_contour_y = self.get_marginal_contour_position(west_contour_x, west_contour_y, 1, True)
        ## extract only the outmost position of X for right side
        right_contour_x, right_contour_y = self.get_marginal_contour_position(east_contour_x, east_contour_y, -1, True)

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(left_contour_x, left_contour_y, color='red', lw=2)
            plt.title('cutted data, along X-direction (from left)')
            plt.show()
            plt.close(f)

        ## reset data beyong left contour
        for jj, yy in enumerate(left_contour_y):
            x_left = left_contour_x[jj]
            self.df_np[yy,0:x_left+1] = np.nan

        if False:
            # Plot contours
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.plot(right_contour_x, right_contour_y, color='blue', lw=2)
            plt.title('cutted data, along X-direction (from right)')
            plt.show()
            plt.close(f)

        ## reset data beyong right contour
        for jj, yy in enumerate(right_contour_y):
            x_right = right_contour_x[jj]
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

        if self.cut_feature:
            axes[0].plot(self.lower_contour_x, self.lower_contour_y, color='red', lw=2)
            axes[0].plot(self.upper_contour_x, self.upper_contour_y, color='blue', lw=2)

        axes[0].set_xlim([0,self.width-1])
        axes[0].set_ylim([self.height-1,0])
        axes[0].set_aspect('equal', 'box')
        axes[1].set_xlim([self.x0,self.x1])
        axes[1].set_ylim([self.y1,self.y0])
        axes[1].set_aspect('equal', 'box')

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

        ## calculate average value for Cu and Al
        self.Cu_ave = np.nanmean(self.df_np_Cu)
        ## calculate variance
        self.variance = np.nansum((self.df_np_Cu - self.Cu_ave)**2) / self.df_new_length
        ## calculate standard deviation
        self.std = np.sqrt(self.variance)

    def cal_Prm4(self):
        '''Calculate Prm4 for each stride.'''

        print('** REPORT ** Start Block Standard Deviation evaluation.')

        ## get filling rate threshold
        self.room_valid_data_ratio_threshold = self.room_valid_data_ratio_threshold.get()

        ## remove the stride_x, stride_y = width_clip, height_clip
        if (self.width_clip in self.stride_x_list) and (self.height_clip in self.stride_y_list):
            self.stride_x_list.remove(self.width_clip)
            self.stride_y_list.remove(self.height_clip)
            ## reset length of stride pattern list
            self.stride_x_list_len, self.stride_y_list_len = len(self.stride_x_list), len(self.stride_y_list)

        self.Prm4_list = [0]*self.stride_x_list_len
        self.Prm4_plot_dict = {}
        self.Prm4_print_dict = {}
        for ii, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
            print(f'** REPORT ** Calculate evaluation for stride (X, Y) = {stride_x}, {stride_y}.')
            ## calculate block number along each X/Y-direction
            block_num_x, block_num_y = int(np.ceil(self.width_clip/stride_x)), int(np.ceil(self.height_clip/stride_y))
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
                    if room_valid_data_ratio > self.room_valid_data_ratio_threshold:
                        room_Cu_average_list[n] = room_Cu_average
                        room_plot_list[n] = True  ## will plot this region
                    else:
                        room_Cu_average_list[n] = np.nan
                        room_plot_list[n] = False ## will NOT plot this region
            ## convert average Cu list into numpy array
            room_Cu_average_np_array = np.array(room_Cu_average_list)
            ## calculate variance of average Cu concentration across all rooms
            all_Cu_variance = self.calculate_variance(room_Cu_average_np_array)
            ## add standard deviation to the list
            self.Prm4_list[ii] = np.sqrt(all_Cu_variance)
            ## add plotting information to the dictionary
            self.Prm4_plot_dict[(stride_x,stride_y)] = room_plot_list
            ## add average value information to the dictionary
            self.Prm4_print_dict[(stride_x,stride_y)] = room_Cu_average_list
        
    def calculate_variance(self, data):
        non_NaN_num_count = np.count_nonzero(~np.isnan(data))
        non_NaN_num_average = np.nanmean(data)
        return np.nansum((data - non_NaN_num_average)**2) / non_NaN_num_count

    def slit_calc(self, axis):
        """Calculate slit averages along X or Y direction.
           * axis : 'x' or 'y'
        """
        # ----- axis-dependent settings -----
        if axis == "x":
            csvfile_path = Path(self.xslit_csvfile_path.get())
            slit_num = self.xslit_num.get()
            slit_width = self.xslit_width.get()
            max_size = self.width
            other_size = self.height
        else:
            csvfile_path = Path(self.yslit_csvfile_path.get())
            slit_num = self.yslit_num.get()
            slit_width = self.yslit_width.get()
            max_size = self.height
            other_size = self.width

        # ----- read or generate slit positions -----
        pos_col = f"{axis.upper()}pos"
        if csvfile_path != Path("."):
            slit_pos_df = pd.read_csv(csvfile_path, names=[pos_col], header=None)
            mask = slit_pos_df[pos_col] >= max_size
            if mask.any():
                filtered_csv = os.path.join(self.output_folder_path, f"slit_{axis}_pos.csv")
                slit_pos_df = slit_pos_df[slit_pos_df[pos_col] < max_size]
                slit_pos_df.to_csv(filtered_csv, index=None)
                print(f"** CAUTION ** {axis.upper()}-positions larger than the domain size have been removed.")
                print(f"** CAUTION ** Filtered positions are exported to {filtered_csv}.")
        else:
            if slit_num == 0:
                print(f"** ERROR ** Please input an integer greater than 0 or import a CSV file in 1b) ({axis.upper()}).")
                exit()
            dpos = max_size // slit_num
            slit_list = [dpos * i for i in range(slit_num)]
            slit_pos_df = pd.DataFrame(np.array(slit_list), columns=[pos_col])

        setattr(self, f"slit_{axis}pos_df", slit_pos_df)

        # ----- calculate slit averages -----
        slit_df = pd.DataFrame({"Position Y|X": np.arange(other_size)})

        for value in slit_pos_df[pos_col]:
            if slit_width % 2 == 1:
                left = value - slit_width // 2
            else:
                left = value - slit_width // 2 + 1
            right = value + slit_width // 2 + 1

            left = max(0, left)
            right = min(max_size, right)

            if axis == "x":
                slit_ave = np.mean(self.df_np[:, left:right], axis=1)
            else:
                slit_ave = np.mean(self.df_np[left:right, :], axis=0)

            slit_df[value] = slit_ave

        # ----- output -----
        output_csv = os.path.join(self.output_folder_path, f"slit_{axis}_average_data.csv")
        if axis == "x":
            slit_df.to_csv(output_csv, index=False, na_rep="NaN")
        else:
            slit_df.T.to_csv(output_csv, header=False, na_rep="NaN")

        print(f"** REPORT ** Slits average along {axis.upper()}-direction is calculated.")
        print(f"** REPORT ** Slits average along {axis.upper()}-direction has been exported to {output_csv}.")

    def plot_result_all(self):
        '''Plot all results.'''

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plotting extent size
        extent = (self.x0-0.5, self.x1+0.5, self.y1+0.5, self.y0-0.5)

        ## plot according to the number of stride pattern
        nstride = len(self.stride_x_list)
        ## number of plots in horizontal / vertical direction
        ## + original/clipped distribution with mass center
        ## + plot concentration
        ## + slits plot (if any)
        nx = 2
        if self.xslit_flag or self.yslit_flag:
            nplots = nstride + 2
            if nstride%nx == 1:
                ny = int((nstride+1)/nx)+1
            else:
                ny = int(nstride//nx)+1
        else:
            nplots = nstride + 1
            if (nstride+1)%nx == 1:
                ny = int((nstride+1)/nx)+1
            else:
                ny = int((nstride+1)/nx)

        ## set containers
        plt.clf()
        fig_x, fig_y = 6*nx, 4*ny
        #fig_x, fig_y = 6.4*nx, 4.8*ny
        fig, axes = plt.subplots(ny, nx, figsize=(fig_x, fig_y))
        im, divider, cax = np.empty_like(axes), np.empty_like(axes), np.empty_like(axes)

        ## plot Cu concentration
        im[0,0] = axes[0,0].imshow(self.df_np[self.y0:self.y1+1,self.x0:self.x1+1], extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
        ## plot mass center positions
        axes[0,0].plot([self.np_data_Cu_xC], [self.np_data_Cu_yC], 'x', color='red')
        axes[0,0].plot([self.np_data_Al_xC], [self.np_data_Al_yC], 'x', color='blue')
        axes[0,0].text(self.np_data_Cu_xC, self.np_data_Cu_yC-20, 'mass center of Cu', color='red')
        axes[0,0].text(self.np_data_Al_xC, self.np_data_Al_yC+20, 'mass center of Al', color='blue')
        ## set grid ratio = 1
        axes[0,0].set_aspect('equal', 'box')
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
            block_num_x = int(np.ceil(self.width_clip/stride_x))
            for n, value in enumerate(self.Prm4_plot_dict[(stride_x, stride_y)]):
                i, j = int(n%block_num_x), int(n//block_num_x)
                istart, jstart = self.x0+i*stride_x, self.y0+j*stride_y
                if not value:
                    square = patches.Rectangle((istart, jstart), stride_x, stride_y, facecolor='blue', alpha=0.2)
                    axes[ipj,ipi].add_patch(square)

            ## set axe range
            axes[ipj,ipi].set_xlim(self.x0-0.5, self.x1+0.5)
            axes[ipj,ipi].set_ylim(self.y1+0.5, self.y0-0.5)
            ## set grid ratio = 1
            axes[ipj,ipi].set_aspect('equal', 'box')
            ## set title of plot
            axes[ipj,ipi].set_title(f"Concentration of Cu, sheet #{self.sheet_num}, stride (X, Y) = {stride_x}, {stride_y}")
            ## plot colorbar
            divider[ipj,ipi] = make_axes_locatable(axes[ipj,ipi])
            cax[ipj,ipi] = divider[ipj,ipi].append_axes("right", size="4%", pad=0.2)
            fig.colorbar(im[ipj,ipi], cax=cax[ipj,ipi], ticks=levels)

        ## plot if either flags is True
        ip = ii+2  ## index of plot including the overall plot
        ipi, ipj = int(ip%nx), int(ip//nx)
        if self.xslit_flag or self.yslit_flag:
            im[ipj,ipi] = axes[ipj,ipi].imshow(self.df_np, extent=extent, vmin=vmin, vmax=vmax, cmap='gray')
            ## plot slits along X-direction
            if self.xslit_flag:
                for xx in self.slit_xpos_df["Xpos"]:
                    axes[ipj,ipi].plot([xx, xx], [-0.5, self.height+0.5], color='purple', linestyle="dashed")
            else:
                pass
            ## plot slits along Y-direction
            if self.yslit_flag:
                for yy in self.slit_ypos_df["Ypos"]:
                    axes[ipj,ipi].plot([-0.5, self.width+0.5], [yy, yy], color='purple', linestyle="dashed")
            else:
                pass
            ## set axe range
            axes[ipj,ipi].set_xlim(-0.5, self.width+0.5)
            axes[ipj,ipi].set_ylim(self.height+0.5, -0.5)
            ## set title of plot
            axes[ipj,ipi].set_title(f"Average slit center positions, sheet #{self.sheet_num}")
            ## plot colorbar
            divider[ipj,ipi] = make_axes_locatable(axes[ipj,ipi])
            cax[ipj,ipi] = divider[ipj,ipi].append_axes("right", size="4%", pad=0.2)
            fig.colorbar(im[ipj,ipi], cax=cax[ipj,ipi], ticks=levels)
        else:
            if ipi+ipj*2 < nplots:
                axes[ipj,ipi].axis('off')
            else:
                pass

        ## cover the plots at the ends
        for ip in range(nplots, nx*ny):
            ipi, ipj = int(ip%nx), int(ip//nx)
            axes[ipj,ipi].axis('off')

        imagepath = os.path.join(self.output_folder_path, 'Cu_Al_mass_center_distance_bsd_plot.png')
        plt.savefig(imagepath)
        print(f'** REPORT ** Output result image to {imagepath}.')

        plt.show()
        plt.close(fig)

    def plot_result_color(self):
        '''Plot data frame into image (color).'''

        ## determine colorbar range and cadence
        vmin, vmax = 0, 100
        levels = np.linspace(vmin, vmax, 11)
        ## plotting extent size
        extent = (self.x0-0.5, self.x1+0.5, self.y1+0.5, self.y0-0.5)

        plt.clf()
        fig, axe = plt.subplots(1, 1, figsize=(5, 3))
        im, divider, cax = np.empty_like(axe), np.empty_like(axe), np.empty_like(axe)
        im = axe.imshow(self.df_np[self.y0:self.y1+1,self.x0:self.x1+1], cmap='jet', extent=extent, vmin=vmin, vmax=vmax, aspect="auto")
        ## plot mass center positions
        axe.plot([self.np_data_Cu_xC], [self.np_data_Cu_yC], marker='s', markersize=4, markerfacecolor='yellow', markeredgecolor="black")
        axe.plot([self.np_data_Al_xC], [self.np_data_Al_yC], marker='s', markersize=4, markerfacecolor='yellow', markeredgecolor="black")
        ## set limit
        axe.set_xlim([self.x0,self.x1])
        axe.set_ylim([self.y1,self.y0])
        ## set grid ratio = 1
        axe.set_aspect('equal', 'box')
        ## set title
        axe.set_title(f'Concentration of Cu, sheet #{self.sheet_num}, cutted data')
        ## plot colorbar
        divider = make_axes_locatable(axe)
        cax = divider.append_axes("right", size="4%", pad=0.2)
        fig.colorbar(im, cax=cax, ticks=levels)
        imagepath = os.path.join(self.output_folder_path, 'contour_image_plot_color.png')
        plt.savefig(imagepath)
        plt.close(fig)
        print(f'** REPORT ** Output data image to {imagepath}.')

    def write_results(self):
        '''Write all results into file and print.'''

        print_padding = '                                           '

        ## write to log file
        outputtxtfilename = 'Uniformity_Evaluation_Results.txt'
        outputtxtfilepath = os.path.join(self.output_folder_path, outputtxtfilename)
        with open(outputtxtfilepath, 'w') as file:
            file.write(f'Evaluation Results (all):\n')
            file.write(f'- Uniformity: Distance bewteen center of material = {self.mass_center_distance}\n')
            file.write(f'- Uniformity: Overall Average = {self.Cu_ave}\n')
            file.write(f'- Uniformity: Overall Standard Deviation = {self.std}\n')
            file.write(f'- Uniformity: Block Standard Deviation value list:\n')
            for kk, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
                ## calculate block number
                block_num_x, block_num_y = int(np.ceil(self.width_clip/stride_x)), int(np.ceil(self.height_clip/stride_y))    
                file.write(f'               - stride (X, Y) = ({stride_x:3}, {stride_y:3}): Block Standard Deviation = {self.Prm4_list[kk]}\n')
                file.write(f'                                           : Average in each room\n')
                for jj in range(block_num_y):
                    line_string = ''
                    for ii in range(block_num_x):
                        n = ii + jj * block_num_x
                        value = self.Prm4_print_dict[(stride_x,stride_y)][n]
                        line_string += '|'+'{:.2f}'.format(value).center(9,' ')
                    line_string = print_padding + line_string + '|\n'
                    file.write(print_padding+'-'*(10*block_num_x+1)+'\n')
                    file.write(line_string)
                file.write(print_padding+'-'*(10*block_num_x+1)+'\n')

        ## write main results to CSV file
        outputcsvfilename = 'Uniformity_Evaluation_Results.csv'
        outputcsvfilepath = os.path.join(self.output_folder_path, outputcsvfilename)
        with open(outputcsvfilepath, 'w') as file:
            file.write(f'Evaluation Results (all):\n')
            file.write(f'Uniformity: Distance bewteen center of material, {self.mass_center_distance}\n')
            file.write(f'Uniformity: Overall Average, {self.Cu_ave}\n')
            file.write(f'Uniformity: Overall Standard Deviation, {self.std}\n')
            file.write(f'Uniformity: Block Standard Deviation value list\n')
            for kk, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
                file.write(f'"stride (X, Y) = ({stride_x:3}, {stride_y:3}): Block Standard Deviation", {self.Prm4_list[kk]}\n')

        ## write block information to CSV file
        for kk, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
            roomfilename = f'room_average_{stride_x}x{stride_y}.csv'
            roomfilepath = os.path.join(self.output_folder_path, roomfilename)
            with open(roomfilepath, 'w') as file:
                ## calculate block number
                block_num_x, block_num_y = int(np.ceil(self.width_clip/stride_x)), int(np.ceil(self.height_clip/stride_y))    
                for jj in range(block_num_y):
                    line_string = ''
                    for ii in range(block_num_x):
                        n = ii + jj * block_num_x
                        value = self.Prm4_print_dict[(stride_x,stride_y)][n]
                        line_string += ','+str(value)
                    line_string = line_string[1:] + '\n'
                    file.write(line_string)

        ## write to console
        print(f'** FINAL ** Evaluation Results (all):')
        print(f'** FINAL ** - Uniformity: Distance bewteen center of material = {self.mass_center_distance}')
        print(f'** FINAL ** - Uniformity: Overall Average = {self.Cu_ave}')
        print(f'** FINAL ** - Uniformity: Overall Standard Deviation = {self.std}')
        print(f'** FINAL ** - Uniformity: Block Standard Deviation value list:')
        for kk, (stride_x, stride_y) in enumerate(zip(self.stride_x_list, self.stride_y_list)):
            ## calculate block number
            block_num_x, block_num_y = int(np.ceil(self.width_clip/stride_x)), int(np.ceil(self.height_clip/stride_y))
            print(f'               - stride (X, Y) = ({stride_x:3}, {stride_y:3}): Block Standard Deviation = {self.Prm4_list[kk]}')
            print(f'                                           : Average in each room')
            for jj in range(block_num_y):
                line_string = ''
                for ii in range(block_num_x):
                    n = ii + jj * block_num_x
                    value = self.Prm4_print_dict[(stride_x,stride_y)][n]
                    line_string += '|'+'{:.2f}'.format(value).center(9,' ')
                line_string = print_padding + line_string + '|'
                print(print_padding+'-'*(10*block_num_x+1))
                print(line_string)
            print(print_padding+'-'*(10*block_num_x+1))
        print(f'** FINAL ** Evaluation of Uniformity is finished.')
        print(f'** FINAL ** Results are exported to {outputtxtfilepath}.')
        print(f'** FINAL ** Results are exported to {outputcsvfilepath}.')

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


if 'Particleworks' in __name__:
    session = PW.Session()
    scene   = session.active_scene
    if scene != None:
        cc = ContourCut(session, scene)
        cc.run()