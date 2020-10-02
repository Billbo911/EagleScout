import os
import csv
import glob
import pandas as pd
import numpy as np

os.chdir ("/EagleScout")
path = '.'

if os.path.exists('combined.csv'):
    os.remove ('combined.csv')


files_in_dir = [ f for f in glob.glob('*-*.csv')]

for filenames in files_in_dir:
    df = pd.read_csv(filenames)
    fName, fExt = (os.path.splitext(filenames))
    sName = fName.split('-')
    df.to_csv('combined.csv', index_label = (sName[0]), mode = 'a')
