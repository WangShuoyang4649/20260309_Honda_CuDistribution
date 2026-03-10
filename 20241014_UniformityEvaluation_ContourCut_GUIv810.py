import os
import datetime
from pathlib import Path

import matplotlib
matplotlib.use('module://pwpy.plot_backend')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

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
            self.prepare_image_data()
            if self.cut_contour:
                self.cut_along_Y_direction()
                self.cut_along_X_direction()
                self.show_compare_image_plot()
            self.create_csv_file()
            self.import_air()
            self.uniformity_evaluation()
            self.adjust_camera()

    def ask_inputs(self):
        '''Ask inputs of 1) data path, 2) sheet number 3) lower threshold 4) upper threshold.'''

        a = PW.Input()
        self.xlsx_path = a.add_file('Data file path (xlsx)')
        self.sheet_num = a.add_integer('Sheet number', 0)
        self.cut_contour = a.add_boolean('Cut data', True)
        self.contour_level0 = a.add_float('Cutting lower threshold', 2)
        self.contour_level1 = a.add_float('Cutting upper threshold', 90)

        if not a.ask():
            return False
        return True

    def prepare_image_data(self):
        '''Prepare image data.'''

        print('** Report ** Prepare image data.')

        ## convert file path and sheet number
        self.xlsx_path = Path(self.xlsx_path.get())
        self.sheet_num = self.sheet_num.get()
        self.cut_contour = self.cut_contour.get()

        ## read file
        self.df = pd.read_excel(self.xlsx_path, sheet_name=self.sheet_num, header=None)  ## read first sheet
        ## fill NaN with 0
        self.df = self.df.fillna(0)
        ## convert dataframe into numpy array
        self.df_np = self.df.values
        ## save original data
        self.df_np_original = self.df_np.copy()

        if self.cut_contour:
            ## remove injection in rasin part  <-- this part is manually executed
            self.df_np[(self.df_np > 90) & (0 <= np.arange(self.df_np.shape[1])) & (np.arange(self.df_np.shape[1]) <= 100) & (np.arange(self.df_np.shape[0])[:, None] > 120)] = 0

        ## get width / height of image
        self.height, self.width = self.df_np.shape
        ## get coordinate
        self.x, self.y = np.arange(self.width), np.arange(self.height)
        self.X, self.Y = np.meshgrid(self.x, self.y)

        if False:
            ## plot data frame into image
            plt.clf()
            f = plt.figure(figsize=(4, 3))
            plt.imshow(self.df_np, cmap='gray')
            plt.show()
            plt.close(f)
 
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

        print('** Report ** Cut contour along Y-direction.')

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

        print('** Report ** Cut contour along X-direction.')

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

        ## plot data frame into image
        plt.clf()
        f = plt.figure(figsize=(8.5, 3))
        gs = gridspec.GridSpec(1, 3, width_ratios=[1, 1, 0.05], wspace=0.05)
        ax1 = f.add_subplot(gs[0])
        ax2 = f.add_subplot(gs[1])
        im1 = ax1.imshow(self.df_np_original, cmap='gray', vmin=0, vmax=100, aspect="auto")
        im2 = ax2.imshow(self.df_np, cmap='gray', vmin=0, vmax=100, aspect="auto")

        ax1.set_xlim([0,self.df_np_original.shape[1]])
        ax1.set_ylim([self.df_np_original.shape[0],0])
        ax2.set_xlim([0,self.df_np_original.shape[1]])
        ax2.set_ylim([self.df_np_original.shape[0],0])

        ax1.set_title('original data')
        ax2.set_title('cutted data')
        ax2.set_yticks([])

        cbar_ax = f.add_subplot(gs[2])
        f.colorbar(im1, cax=cbar_ax)
        plt.show()
        plt.close(f)

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
        ## drop NaN of df_new
        df_new.dropna(axis=0, how='any', inplace=True)
        ## get length of df_new
        self.df_new_length = len(df_new.index)
        ## find min / max of X / Y in cutted contour
        self.x_min_cut, self.x_max_cut = df_new['x'].min(), df_new['x'].max()
        self.y_min_cut, self.y_max_cut = df_new['y'].min(), df_new['y'].max()
        print('** Report ** Data is prepared into data frame.')

        ## create output folder
        datetimestr = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder_name = f"{datetimestr}_CuDistribution_Output"
        self.output_folder_path = os.path.join(self.scene.path(PW.SCENE_PATH_root_dir), output_folder_name)
        os.mkdir(self.output_folder_path)

        ## output data into csv file
        output_data_path = os.path.join(self.output_folder_path, 'Cu_data.csv')
        df_new.to_csv(output_data_path, index=True, header=False)
        ## make list data
        output_list_path = os.path.join(self.output_folder_path, 'list_data.csv')
        with open(output_list_path, 'w') as outputfile:
            outputfile.write('0,Cu_data.csv\n')

        print('** Report ** Data is exported into CSV files.')

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

        print('** Report ** Air node is imported into scene.')

    def uniformity_evaluation(self):
        '''Evaluate uniformity of mixing.'''

        ## split data into Cu and Al
        np_data_Cu = self.df_np.copy()
        np_data_Al = 100-self.df_np.copy()

        ## calculate "center of material"
        np_data_Cu_xC = np.nansum(np_data_Cu * self.X) / np.nansum(np_data_Cu)
        np_data_Cu_yC = np.nansum(np_data_Cu * self.Y) / np.nansum(np_data_Cu)
        np_data_Al_xC = np.nansum(np_data_Al * self.X) / np.nansum(np_data_Al)
        np_data_Al_yC = np.nansum(np_data_Al * self.Y) / np.nansum(np_data_Al)

        ## calculate distance between "center of material"
        mass_center_distance = np.sqrt( (np_data_Cu_xC-np_data_Al_xC)**2 + (np_data_Cu_yC-np_data_Al_yC)**2 )

        ## add probe point as center of material Cu
        self.scene.create_node(PW.NODE_probe_point)
        self.scene.write()
        for node in self.scene.nodes:
            if node.node_type == PW.NODE_probe_point:
                node.name = 'Cu_mass_center'
                node['probe.point.x'] = np_data_Cu_xC
                node['probe.point.y'] = np_data_Cu_yC
                node['probe.point.z'] = 0
        
        ## add probe point as center of material Al
        self.scene.create_node(PW.NODE_probe_point)
        self.scene.write()
        for node in self.scene.nodes:
            if (node.node_type == PW.NODE_probe_point) and (node.name != 'Cu_mass_center'):
                node.name = 'Al_mass_center'
                node['probe.point.x'] = np_data_Al_xC
                node['probe.point.y'] = np_data_Al_yC
                node['probe.point.z'] = 0

        variance = np.nansum((np_data_Cu - np.nanmean(np_data_Cu))**2) / self.df_new_length

        ## write result
        print(f'** Evaluation Results:')
        print(f'** - Uniformity: Distance bewteen center of material = {mass_center_distance}')
        print(f'** - Uniformity: Variance = {variance}')
        print(f'** Evaluation of Uniformity is finished.')

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

