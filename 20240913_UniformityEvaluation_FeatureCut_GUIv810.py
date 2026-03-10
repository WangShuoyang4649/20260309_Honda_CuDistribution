import os
import datetime
from pathlib import Path

import matplotlib
matplotlib.use('module://pwpy.plot_backend')
import matplotlib.pyplot as plt
import cv2
import numpy as np
import pandas as pd

from ultralytics import YOLO

import pwpy as PW

class FeatureCut:
    def __init__(self, session, scene):
        '''Initialize class.'''

        self.session = session
        self.scene = scene
        self.conf = 0.005  ## <-- change this parameter if necessary

    def run(self):
        '''Run main function.'''

        if self.ask_image_path():
            self.show_image_plot()
            if self.ask_feature_bbox():
                self.cut_normalize_image()
                if self.ask_cut_further():
                    self.cut_image_further()
                self.create_csv_file()
                self.import_air()
                self.uniformity_evaluation()
                self.adjust_camera()

    def ask_image_path(self):
        '''Ask image input path.'''

        a = PW.Input()
        self.image_path = a.add_string('Data file path (png, jpg)','')
    
        if not a.ask():
            return False
        return True

    def show_image_plot(self):
        '''Show image in console.'''

        ## read image
        self.image_path = Path(self.image_path.get()[8:])
        self.image = cv2.imread(self.image_path)

        ## using YOLO to detect feature
        model = YOLO('yolov8n.pt', verbose=False)  ## hide category message output
        results = model(self.image_path, conf=self.conf)

        ## find bounding boxes
        self.bboxes = results[0].boxes.xyxy
        ## find number of boxes
        nboxes = len(self.bboxes)

        if nboxes > 0:
            ## plot image with the bounding boxes
            plt.clf()
            fig_width = int(4*nboxes)
            f = plt.figure(figsize=(fig_width, 3))
            for ii, bbox in enumerate(self.bboxes):
                a = f.add_subplot(1,nboxes,ii+1)
                a.imshow(self.image, cmap='gray')
                x0, y0, x1, y1 = bbox
                plt.plot([x0,x0],[y0,y1],color='red')
                plt.plot([x1,x1],[y0,y1],color='red')
                plt.plot([x0,x1],[y0,y0],color='red')
                plt.plot([x0,x1],[y1,y1],color='red')
                plt.title(f'#{ii}')
            plt.show()
            plt.close(f)
        else:
            print('** Caution ** No feature was detected. Please tune down self.conf.')
            exit()

    def ask_feature_bbox(self):
        '''Ask which bounding box to be used.'''

        a = PW.Input()
        self.box_select = a.add_integer('Select number of bounding box', 0)

        if not a.ask():
            return False
        return True

    def cut_normalize_image(self):
        '''Cut image by selected box.'''

        ## convert image to grey
        self.image_grey = cv2.cvtColor(self.image, cv2.COLOR_RGB2GRAY)
        ## normalize image to 0-100
        self.image_grey = (self.image_grey-np.min(self.image_grey))/(np.max(self.image_grey)-np.min(self.image_grey))*100
        ## cut image
        x0, y0, x1, y1 = self.bboxes[self.box_select.get()]
        self.x_left, self.x_right = int(x0), int(x1)
        self.y_lower, self.y_upper = int(y0), int(y1)
        self.image_grey_cut = self.image_grey[self.y_lower:self.y_upper+1, self.x_left:self.x_right+1]

        ## get X / Y center and width / height
        self.x_width = self.x_right+1-self.x_left
        self.y_height = self.y_upper+1-self.y_lower
        self.x_center = int((self.x_right-self.x_left)/2)
        self.y_center = int((self.y_upper-self.y_lower)/2)

        print('** Report ** Cut image (normalized) is shown below.')
        print(f'** Report ** - Cutting range (X-direction): {self.x_left}-{self.x_right}')
        print(f'** Report ** - Cutting range (Y-direction): {self.y_lower}-{self.y_upper}')

        plt.clf()
        f = plt.figure(figsize=(4, 3))
        #a = f.add_subplot(1,1,1)
        plt.imshow(self.image_grey_cut, cmap='gray')
        plt.plot([0, self.x_width], [self.y_center, self.y_center], color='blue')
        plt.plot([self.x_center, self.x_center], [0, self.y_height], color='blue')
        plt.xlim([0, self.x_width])
        plt.ylim([self.y_height, 0])

        plt.show()
        plt.close(f)

    def ask_cut_further(self):
        '''Ask whether to cut image further.'''

        a = PW.Input()
        self.x_center = a.add_integer('Feature center @ X-position', self.x_center)
        self.y_center = a.add_integer('Feature center @ Y-position', self.y_center)
        self.x_cut_width = a.add_integer('Cutting half width along X-direction', int(self.x_width/2))
        self.y_cut_height = a.add_integer('Cutting half height along Y-direction', int(self.y_height/2))

        if not a.ask():
            return False
        return True

    def cut_image_further(self):
        '''Cut image further.'''

        ## get new X / Y center
        self.x_center = self.x_center.get()
        self.y_center = self.y_center.get()
        ## get cutting ratio
        self.x_cut_width = self.x_cut_width.get()
        self.y_cut_height = self.y_cut_height.get()
        ## get new bounding box
        self.x_left = self.x_center - self.x_cut_width
        self.x_right = self.x_center + self.x_cut_width
        self.y_lower = self.y_center - self.y_cut_height
        self.y_upper = self.y_center + self.y_cut_height
        ## cut image
        self.image_grey_cut_copy = self.image_grey_cut.copy()
        self.image_grey_cut = self.image_grey_cut[self.y_lower:self.y_upper+1, self.x_left:self.x_right+1]

        print('** Report ** Re-cut image is shown below.')
        print(f'** Report ** - Cutting range (X-direction): {self.x_left}-{self.x_right}')
        print(f'** Report ** - Cutting range (Y-direction): {self.y_lower}-{self.y_upper}')

        plt.clf()
        f = plt.figure(figsize=(4, 3))
        #a = f.add_subplot(1,1,1)
        plt.imshow(self.image_grey_cut_copy, cmap='gray')
        plt.plot([self.x_left, self.x_left], [self.y_lower, self.y_upper], color='red')
        plt.plot([self.x_right, self.x_right], [self.y_lower, self.y_upper], color='red')
        plt.plot([self.x_left, self.x_right], [self.y_lower, self.y_lower], color='red')
        plt.plot([self.x_left, self.x_right], [self.y_upper, self.y_upper], color='red')

        plt.show()
        plt.close(f)

    def create_csv_file(self):
        '''Create CSV file from image data.'''

        ## total number of data
        self.img_width_cut, self.img_height_cut = self.x_right+1-self.x_left, self.y_upper+1-self.y_lower
        df_len = int( self.img_width_cut * self.img_height_cut )

        ## read file
        self.df = pd.DataFrame(self.image_grey_cut)
        ## replot data frames
        x_coords, y_coords = np.indices((self.img_width_cut, self.img_height_cut))
        x_flat, y_flat, values_flat = x_coords.flatten(), y_coords.flatten(), self.image_grey_cut.transpose().flatten()
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
        np_data_Cu = self.image_grey_cut.copy()
        np_data_Al = 100-self.image_grey_cut.copy()

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

        xc = int(len(self.image_grey_cut[0,:])/2) - 0.5
        yc = int(len(self.image_grey_cut[:,0])/2) - 0.5

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

