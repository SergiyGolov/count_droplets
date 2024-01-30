#!/usr/bin/env python3

import matplotlib.pyplot as plt 
import numpy as np

bins=[0]+[round(3.5/10*(x+1),2) for x in range(10)]+[3.85]

def plot_bins(data,title,image_paths):
    clipped_data=[np.clip(x,0,3.51) for x in data]
    fig=plt.figure(f'Droplet diameter histogram for {title}')
    plt.hist(clipped_data, bins=bins, edgecolor="black",label=image_paths)
    if len(data)>1:
        plt.legend(loc='upper right')

    plt.xticks(ticks=bins,labels=bins[:-1]+['>3.5'])
    plt.xlabel('Diameter (µm)')
    plt.ylabel('Number of droplets')
    plt.title(f'Droplet diameter histogram for\n{title}')
    plt.show()