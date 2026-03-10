import os

import numpy as np
import pandas as pd

import matplotlib.pyplot as plt
from mpl_toolkits.axes_grid1 import make_axes_locatable

input_filepath = r'D:\2024\202406\20240603_Honda_CuDistribution\raw_data.xlsx'  ## input filepath

df_low = pd.read_excel(input_filepath, sheet_name=0, header=None)  ## read "low" sheet
df_high = pd.read_excel(input_filepath, sheet_name=1, header=None) ## read "high" sheet

## convert df to numpy array
np_low = df_low.values
np_high = df_high.values

## get number of column and row
np_low_height, np_low_width = np_low.shape
np_high_height, np_high_width = np_high.shape

## create coordinate
x_low, y_low = np.arange(np_low_width), np.arange(np_low_height)
x_high, y_high = np.arange(np_high_width), np.arange(np_high_height)

## generate mesh
xx_low, yy_low = np.meshgrid(x_low, y_low)
xx_high, yy_high = np.meshgrid(x_high, y_high)

## split data into Cu and Al
np_low_Cu, np_low_Al = np_low.copy(), 100-np_low.copy()
np_high_Cu, np_high_Al = np_high.copy(), 100-np_high.copy()

## calculate "center of material"
np_low_Cu_xC = (np_low_Cu * xx_low).sum() / np_low_Cu.sum()
np_low_Cu_yC = (np_low_Cu * yy_low).sum() / np_low_Cu.sum()
np_low_Al_xC = (np_low_Al * xx_low).sum() / np_low_Al.sum()
np_low_Al_yC = (np_low_Al * yy_low).sum() / np_low_Al.sum()

np_high_Cu_xC = (np_high_Cu * xx_high).sum() / np_high_Cu.sum()
np_high_Cu_yC = (np_high_Cu * yy_high).sum() / np_high_Cu.sum()
np_high_Al_xC = (np_high_Al * xx_high).sum() / np_high_Al.sum()
np_high_Al_yC = (np_high_Al * yy_high).sum() / np_high_Al.sum()

## calculate distance between "center of material"
np_low_d = np.sqrt( (np_low_Cu_xC-np_low_Al_xC)**2 + (np_low_Cu_yC-np_low_Al_yC)**2 )
np_high_d = np.sqrt( (np_high_Cu_xC-np_high_Al_xC)**2 + (np_high_Cu_yC-np_high_Al_yC)**2 )

## output "center of material"
print()
print('Print "Center of Material": ')
print(f'- Low Uniformity: Distance bewteen center of material = {np_low_d}')
print(f'- High Uniformity: Distance bewteen center of material = {np_high_d}')

## calculate variance
np_low_Cu_V = ((np_low_Cu - np_low_Cu.mean())**2).sum() / (np_low_height*np_low_width)
np_high_Cu_V = ((np_high_Cu - np_high_Cu.mean())**2).sum() / (np_high_height*np_high_width)
print()
print('Print "Variance of Material": ')
print(f'- Low Uniformity: Variance = {np_low_Cu_V}')
print(f'- High Uniformity: Variance = {np_high_Cu_V}')

## plot distribution
vmin, vmax = 0, 100
levels = np.linspace(vmin, vmax, 11)

fig, ax = plt.subplots(2,figsize=(8,9))

im0 = ax[0].imshow(np_low_Cu, vmin=vmin, vmax=vmax, cmap='gray_r')
divider0 = make_axes_locatable(ax[0])
cax0 = divider0.append_axes("right", size="4%", pad=0.2)
ax[0].plot([np_low_Cu_xC],[np_low_Cu_yC],'x',color='red')
ax[0].plot([np_low_Al_xC],[np_low_Al_yC],'x',color='blue')
ax[0].set_title("Concentration of Cu, Uniformity: low")
ax[0].text(np_low_Cu_xC,np_low_Cu_yC-1,'Center of Cu',size=10,color='red')
ax[0].text(np_low_Al_xC,np_low_Al_yC+3,'Center of Al',size=10,color='blue')
fig.colorbar(im0, cax=cax0, ticks=levels)

im1 = ax[1].imshow(np_high_Cu, vmin=vmin, vmax=vmax, cmap='gray_r')
divider1 = make_axes_locatable(ax[1])
cax1 = divider1.append_axes("right", size="4%", pad=0.2)
ax[1].plot([np_high_Cu_xC],[np_high_Cu_yC],'x',color='red')
ax[1].plot([np_high_Al_xC],[np_high_Al_yC],'x',color='blue')
ax[1].set_title("Concentration of Cu, Uniformity: high")
ax[1].text(np_high_Cu_xC,np_high_Cu_yC-2,'Center of Cu',size=10,color='red')
ax[1].text(np_high_Al_xC,np_high_Al_yC+4,'Center of Al',size=10,color='blue')
fig.colorbar(im1, cax=cax1, ticks=levels)
plt.show()