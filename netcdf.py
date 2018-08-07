#Python 3.5
#recommend using Anaconda or Miniconda 

import pandas as pd
import netCDF4 as nc4
import itertools

#read your Excel tab or tabs
#recommend to read each tab as its own dataframe, then assign each tab
#to its own group
#sheetname is the exact name of your excel worksheet

df = pd.read_excel('MO_2006.xlsx', sheetname = 'num_crash')
#df2 = pd.read_excel('MO_2006.xlsx', sheetname = 'num_injured')
#df3 = pd.read_excel('MO_2006.xlsx', sheetname = 'num_fatal')

#assign the columns in each dataframe to arrays
#change the name of the dataframe (df, df2, etc) and column names (['Rain'])
#depending on which sheet you're reading
 
indx = list(itertools.chain.from_iterable(df[['Index']].as_matrix().tolist()))
yr = list(itertools.chain.from_iterable(df[[2006]].as_matrix().tolist()))
fRain = list(itertools.chain.from_iterable(df[['Freezing Rain']].as_matrix().tolist()))
rn = list(itertools.chain.from_iterable(df[['Rain']].as_matrix().tolist()))
sHail = list(itertools.chain.from_iterable(df[['Sleet/Hail']].as_matrix().tolist()))
snw = list(itertools.chain.from_iterable(df[['Snow']].as_matrix().tolist()))
tot = list(itertools.chain.from_iterable(df[['Total wx']].as_matrix().tolist()))

#create a netCDF dataset file
#the 'w' is so you can write to it
f = nc4.Dataset('dataset.nc', 'w', format='NETCDF4')

#create a group for each of your excel tabs

nCrash = f.createGroup('num_crash')
#nInjure = f.createGroup('num_injured')
#nFatal = f.createGroup('num_fatal')


#specify the dimensions that you want to assign to each group
nCrash.createDimension('indx', len(indx))
nCrash.createDimension('yr', len(yr))
nCrash.createDimension('fRain', len(fRain))
nCrash.createDimension('rn', len(rn))
nCrash.createDimension('sHail', len(sHail))
nCrash.createDimension('snw', len(snw))
nCrash.createDimension('tot', len(tot))


#build variables to allocate some storage for your data 
index = nCrash.createVariable('Index', 'f4', 'indx')
year = nCrash.createVariable('Year', 'f4', 'yr')
freezRain = nCrash.createVariable('Freezing Rain', 'f4', 'fRain')
rain = nCrash.createVariable('Rain', 'f4', 'rn')
#slash is an illegal character, so changed to hyphen
sleetHail = nCrash.createVariable('Sleet-Hail', 'f4', 'sHail')
snow = nCrash.createVariable('Snow', 'f4', 'snw')
total = nCrash.createVariable('Total wx', 'f4', 'tot')

#pass data to each of your variables
#the [:] at the end of the variable instance is necessary

index[:] = indx
year[:] = yr
freezRain[:] = fRain
rain[:] = rn
sleetHail[:] = sHail
snow[:] = snw
total[:] = tot

#close the file
f.close()

#to reopen and read it, do this:
f = nc4.Dataset('dataset.nc', 'r')