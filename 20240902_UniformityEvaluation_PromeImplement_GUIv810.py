import os
import sys
import datetime

import numpy as np
import pandas as pd

##---------
## Inputs  
##---------
sys.path.append(r'D:\Universal\Particleworks_Installation\Particleworks_InstallPath\Particleworks_8.1.0_240617\bin')
xlsx_file_path = r'D:\2024\202406\20240603_Honda_CuDistribution\Honda_DummyData.xlsx'
sheet_num = 4  ## start from 0
##---------

import pwpy as PW

def convert_xlsx_file(scene):
    '''Convert xlsx file into csv file.'''

    ## read file
    df = pd.read_excel(xlsx_file_path, sheet_name=sheet_num, header=None)  ## read first sheet
    ## get number of row / column
    df_row_num = len(df)
    df_column_num = len(df.columns)
    ## total number of data
    df_len = int(df_row_num * df_column_num)

    ## convert df to new df with position
    df_new = pd.DataFrame(np.zeros((df_len, 7)))
    ## 1. 1 column = index
    df_new[0] = np.arange(df_len)
    ## 2. 2 - 4 column = position, 5 column
    for ii, row in df.iterrows():
        for jj, value in enumerate(row):
            ipos = int( ii * df_column_num + jj )
            df_new.iloc[ipos,1] = jj
            df_new.iloc[ipos,2] = ii
            df_new.iloc[ipos,3] = 0
            df_new.iloc[ipos,4] = value
            df_new.iloc[ipos,5] = value
            df_new.iloc[ipos,6] = value

    ## create output folder
    datetimestr = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder_name = f"{datetimestr}_CuDistribution_Output"
    output_folder_path = os.path.join(scene.path(PW.SCENE_PATH_root_dir), output_folder_name)
    os.mkdir(output_folder_path)

    ## output data into csv file
    output_data_path = os.path.join(output_folder_path, 'Cu_data.csv')
    df_new.to_csv(output_data_path, index=False, header=False)
    ## make list data
    output_list_path = os.path.join(output_folder_path, 'list_data.csv')
    with open(output_list_path, 'w') as outputfile:
        outputfile.write('0,Cu_data.csv\n')
    
    return df, output_folder_path

def import_air(session, scene, output_folder_path):
    '''Import air node with created data.'''

    ## create air parameter
    p = scene.create_parameters(PW.PARAMETERS_import_air)
    p['filePath'] = os.path.join(output_folder_path, 'list_data.csv')
    ## submit task and wait
    q = session.task_queue
    t = q.submit(scene, p)
    q.wait(-1)

    ## change display of particle to show color
    for node in scene.nodes:
        if node.node_type == PW.NODE_air:
            node['material.useColorMap'] = True
            break
    scene.write()

def uniformity_evaluation(scene, df):
    '''Evaluate uniformity of mixing.'''

    ## corner coordinate
    x0, x1 = 0, len(df.columns)
    y0, y1 = 0, len(df)

    ## data in numpy form
    np_data = df.values

    ## create coordinate
    x_data, y_data = np.arange(x0, x1), np.arange(y0, y1)
    ## generate mesh
    xx_data, yy_data = np.meshgrid(x_data, y_data)

    ## split data into Cu and Al
    np_data_Cu = np_data[y0:y1,x0:x1].copy()
    np_data_Al = 100-np_data[y0:y1,x0:x1].copy()

    ## calculate "center of material"
    np_data_Cu_xC = (np_data_Cu * xx_data).sum() / np_data_Cu.sum()
    np_data_Cu_yC = (np_data_Cu * yy_data).sum() / np_data_Cu.sum()
    np_data_Al_xC = (np_data_Al * xx_data).sum() / np_data_Al.sum()
    np_data_Al_yC = (np_data_Al * yy_data).sum() / np_data_Al.sum()

    ## calculate distance between "center of material"
    mass_center_distance = np.sqrt( (np_data_Cu_xC-np_data_Al_xC)**2 + (np_data_Cu_yC-np_data_Al_yC)**2 )

    ## add probe point as center of material Cu
    scene.create_node(PW.NODE_probe_point)
    scene.write()
    for node in scene.nodes:
        if node.node_type == PW.NODE_probe_point:
            node.name = 'Cu_mass_center'
            node['probe.point.x'] = np_data_Cu_xC
            node['probe.point.y'] = np_data_Cu_yC
            node['probe.point.z'] = 0
    
    ## add probe point as center of material Al
    scene.create_node(PW.NODE_probe_point)
    scene.write()
    for node in scene.nodes:
        if (node.node_type == PW.NODE_probe_point) and (node.name != 'Cu_mass_center'):
            node.name = 'Al_mass_center'
            node['probe.point.x'] = np_data_Al_xC
            node['probe.point.y'] = np_data_Al_yC
            node['probe.point.z'] = 0

    ## calculate variance
    np_data_width = x1-x0
    np_data_height = y1-y0

    variance = ((np_data_Cu - np_data_Cu.mean())**2).sum() / (np_data_height*np_data_width)

    ## write result
    print(f'** Evaluation Results:')
    print(f'** - Uniformity: Distance bewteen center of material = {mass_center_distance}')
    print(f'** - Uniformity: Variance = {variance}')
    print(f'** Evaluation of Uniformity is finished.')

def adjust_camera(scene, df):
    '''Adjust camera for better view.'''

    xc = int(len(df.iloc[0,:])/2) - 0.5
    yc = int(len(df.iloc[:,0])/2) - 0.5

    ## change ortho of camera
    for node in scene.nodes:
        if node.name == 'camera':
            node['ortho'] = True
            node['transform.location'] = (xc, yc, 0)
            node['transform.rotation'] = (180, 0, 0)
            node['gaze'] = max(xc,yc)+1
            break

def main():
    session = PW.Session()
    scene   = session.active_scene
    if scene != None:
        df, output_folder_path = convert_xlsx_file(scene)
        import_air(session, scene, output_folder_path)
        uniformity_evaluation(scene, df)
        adjust_camera(scene, df)


if 'Particleworks' in __name__:
    main()
