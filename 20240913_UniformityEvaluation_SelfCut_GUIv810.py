import os
import datetime
from pathlib import Path

import matplotlib
matplotlib.use('module://pwpy.plot_backend')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import pwpy as PW

class FeatureCut:
    def __init__(self, session, scene):
        '''Initialize class.'''

        self.session = session
        self.scene = scene

    def run(self):
        '''Run main function.'''

        if self.ask_inputs():
            self.show_image_plot()
            if self.ask_feature_bbox():
                self.cut_normalize_image()
                self.create_csv_file()
                self.import_air()
                self.uniformity_evaluation()
                self.adjust_camera()

    def ask_inputs(self):
        '''Ask inputs of 1) data path, 2) sheet number.'''

        a = PW.Input()
        self.xlsx_path = a.add_string('Data file path (xlsx)', '')
        self.sheet_num = a.add_integer('Sheet number', 0)

        if not a.ask():
            return False
        return True

    def show_image_plot(self):
        '''Show image in console.'''

        ## convert file path and sheet number
        self.xlsx_path = Path(self.xlsx_path.get())
        self.sheet_num = self.sheet_num.get()

        ## read file
        self.df = pd.read_excel(self.xlsx_path, sheet_name=self.sheet_num, header=None)  ## read first sheet
        ## get width / height of image
        self.img_height = len(self.df)
        self.img_width = len(self.df.columns)

        ## plot data frame into image
        plt.clf()
        f = plt.figure(figsize=(4, 3))
        #a = f.add_subplot(1,1,1)
        plt.imshow(self.df, cmap='gray')

        plt.show()
        plt.close(f)

    def ask_feature_bbox(self):
        '''Ask feature bounding box boundaries.'''

        a = PW.Input()
        self.x_left = a.add_integer('X Left Boundary', 0)
        self.x_right = a.add_integer('X Right Boundary', self.img_width-1)
        self.y_lower = a.add_integer('Y Lower Boundary', 0)
        self.y_upper = a.add_integer('Y Upper Boundary', self.img_height-1)
        if not a.ask():
            return False
        return True

    def cut_normalize_image(self):
        '''Cut image by input size.'''

        ## cut image
        self.x_left, self.x_right = self.x_left.get(), self.x_right.get()
        self.y_lower, self.y_upper = self.y_lower.get(), self.y_upper.get()
        self.df_np = self.df.values
        self.df_np_cut = self.df_np[self.y_lower:self.y_upper+1, self.x_left:self.x_right+1]

        print('** Report ** Cut image (normalized) is shown below.')
        print(f'** Report ** - Cutting range (X-direction): {self.x_left}-{self.x_right}')
        print(f'** Report ** - Cutting range (Y-direction): {self.y_lower}-{self.y_upper}')

        plt.clf()
        f = plt.figure(figsize=(4, 3))
        #a = f.add_subplot(1,1,1)
        plt.imshow(self.df_np_cut, cmap='gray')

        plt.show()
        plt.close(f)

    def create_csv_file(self):
        '''Create CSV file from image data.'''

        ## total number of data
        self.img_width_cut, self.img_height_cut = self.x_right+1-self.x_left, self.y_upper+1-self.y_lower
        df_len = int( self.img_width_cut * self.img_height_cut )

        ## read file
        self.df = pd.DataFrame(self.df_np_cut)
        ## replot data frames
        x_coords, y_coords = np.indices((self.img_width_cut, self.img_height_cut))
        x_flat, y_flat, values_flat = x_coords.flatten(), y_coords.flatten(), self.df_np_cut.transpose().flatten()
        df_new = pd.DataFrame({
            'x': x_flat,
            'y': y_flat,
            'z': [0]*df_len,
            'vx': values_flat,
            'vy': values_flat,
            'vz': values_flat,
        })
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

        ## create coordinate
        x_data, y_data = np.arange(self.img_width_cut), np.arange(self.img_height_cut)
        ## generate mesh
        xx_data, yy_data = np.meshgrid(x_data, y_data)

        ## split data into Cu and Al
        np_data_Cu = self.df_np_cut.copy()
        np_data_Al = 100-self.df_np_cut.copy()

        ## calculate "center of material"
        np_data_Cu_xC = (np_data_Cu * xx_data).sum() / np_data_Cu.sum()
        np_data_Cu_yC = (np_data_Cu * yy_data).sum() / np_data_Cu.sum()
        np_data_Al_xC = (np_data_Al * xx_data).sum() / np_data_Al.sum()
        np_data_Al_yC = (np_data_Al * yy_data).sum() / np_data_Al.sum()

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

        variance = ((np_data_Cu - np_data_Cu.mean())**2).sum() / (self.img_height_cut*self.img_width_cut)

        ## write result
        print(f'** Evaluation Results:')
        print(f'** - Uniformity: Distance bewteen center of material = {mass_center_distance}')
        print(f'** - Uniformity: Variance = {variance}')
        print(f'** Evaluation of Uniformity is finished.')

    def adjust_camera(self):
        '''Adjust camera for better view.'''

        xc = int(len(self.df_np_cut[0,:])/2) - 0.5
        yc = int(len(self.df_np_cut[:,0])/2) - 0.5

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
        fc = FeatureCut(session, scene)
        fc.run()

if 'Particleworks' in __name__:
    main()