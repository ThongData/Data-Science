# -*- coding: utf-8 -*-
"""
Created on Wed Mar  6 19:52:30 2024

@author: Thong.Nguyen
"""

import pandas as pd
import numpy as np
import os
import glob

cwd = os.getcwd()


def append(path):
    frames = []
    frames_timeline =[]
    for root, dirs, files in os.walk(path):
        for file in files:
            file_with_path = os.path.join(root, file)
            print(file_with_path)
            
            df = pd.read_excel(file_with_path, sheet_name="Res Plan",header=None)
            
            df["Project Name"] =np.where(df[2]=='Project Name',df[3], None)
            df.loc[:,"Project Name"] = df.loc[:,"Project Name"].ffill()
            
            df_2= df.copy()
            new_list=[]
            Actual_idx= list(df_2[df_2[2]=='Actual Timeline'].index+1)
            for j in Actual_idx:
                for i in range(8):
                    new_list.append(j+i)
            df_timeline=df_2.iloc[new_list,:]
            df_timeline = df_timeline.iloc[:,np.r_[2:5,-1]]
            df_timeline.columns =[["Phase","Start_Date","End_Date","Project Name"]]
            frames_timeline.append(df_timeline)

            df_=df.iloc[23:,:].copy()
            df_.columns = df_.iloc[0]

            df_=df_[(df_.iloc[:,1].notnull()) & (df_.iloc[:,1]!=0)]
            df_=df_.iloc[:,np.r_[4,24:len(df_.columns)]]
            df_.columns.values[0]="Emp_Name"
            df_=df_.iloc[:,~df_.columns.isnull()]
            
            df_.columns.values[1:-1]= pd.to_datetime(df_.columns[1:-1]).strftime('%Y-%m-%d')
            df_.columns.values[-1]="Project Name"
            df_=pd.melt(df_,id_vars =["Emp_Name",'Project Name'], value_vars =list(df_.columns[1:-1]))
            df_["value"]=df_["value"].fillna(0)
            df_.columns.values[2]="Date_"
            df_.iloc[:,2] = pd.to_datetime(df_.iloc[:,2], format='%Y-%m-%d')

            frames.append(df_)
    df = pd.concat(frames, axis=0, ignore_index=True)
    df_ = pd.concat(frames_timeline, axis=0, ignore_index=True)

    return df, df_

 # path：The folder path where storage all the excel files 
df,df_ = append(cwd)


df.to_excel("merged_excel.xlsx")
df_.to_excel("merged_timeline.xlsx")



df = pd.read_excel("C:/Users/thong.nguyen/Desktop/TCCC/★Dashboard/New Dashboard/Data/Project Tracker & CM Management/Sparkling/Fanta/ExLOOP Bench Mark Plan FA Core 2024.xlsx", sheet_name="Res Plan",header=None)
df["Project Name"] =np.where(df[2]=='Project Name',df[3], None)
df.loc[:,"Project Name"] = df.loc[:,"Project Name"].ffill()
new_list=[]
Actual_idx= list(df[df[2]=='Actual Timeline'].index+1)
for j in Actual_idx:
    for i in range(8):
        new_list.append(j+i)
df_timeline=df.iloc[new_list,:]
df_timeline = df_timeline.iloc[:,np.r_[2:4,-1]]
df_timeline.columns =[["Phase","Delivery Date","Project Name"]]

df_timeline.to_excel("merged_timeline.xlsx")



# df = pd.read_excel("C:/Users/thong.nguyen/Desktop/TCCC/★Dashboard/New Dashboard/Data/Project Tracker & CM Management/Sparkling/Fanta/ExLOOP Bench Mark Plan FA Core 2024.xlsx", sheet_name="Res Plan",header=None)
# df["Project Name"] =np.where(df[2]=='Project Name',df[3], None)
# df.loc[:,"Project Name"] = df.loc[:,"Project Name"].ffill()
# df_=df.iloc[23:,:].copy()
# df_.columns = df_.iloc[0]

# df_=df_[(df_.iloc[:,1].notnull()) & (df_.iloc[:,1]!=0)]

# df_=df_.iloc[:,np.r_[4,24:len(df_.columns)]]
# df_.columns.values[0]="Emp_Name"


# df_=df_.iloc[:,~df_.columns.isnull()]

# df_.columns.values[1:-1]= pd.to_datetime(df_.columns[1:-1]).strftime('%Y-%m-%d')
# df_.columns.values[-1]="Project Name"
# df_=pd.melt(df_,id_vars =["Emp_Name",'Project Name'], value_vars =list(df_.columns[1:-1]))
# df_["value"]=df_["value"].fillna(0)
# df_.columns.values[2]="Date_"
# df_.iloc[:,2] = pd.to_datetime(df_.iloc[:,2], format='%Y-%m-%d')




# df_=df_.iloc[:,np.r_[4,24:len(df_.columns)]]

# df_.columns.values[1:-1]= pd.to_datetime(df_.columns[1:-1]).strftime('%Y-%m-%d')

# df_=df_[(df_.iloc[:,1].notnull()) & (df_.iloc[:,1]!=0)]
# df_=pd.melt(df_,id_vars =["Emp_Name",'Project Name'], value_vars =list(df_.columns[1:-1]))
# df_["value"]=df_["value"].fillna(0)






# df["Project Name"] =np.where(df[2]=='Project Name',df[3], None)
# df.loc[:,"Project Name"] = df.loc[:,"Project Name"].ffill()


# df_=df.iloc[22:,:].copy()
# # test=df.applymap(lambda x : 'Week' in str(x)).any()[24:]
# # test_= df.loc[:,24:].iloc[:test]


# df_["count"]=df_.apply(lambda x: np.sum(x.astype(str).str.contains("Week")) ,axis=1)
# df_.columns.values[24:-2]= df_.iloc[np.argmax(df_["count"])+2,24:-2]



# test_.iloc[:,:].iloc[23,:]

# test_.iloc[:,1]
# df_= df_.transpose()

# df_[23]=df_[23].fillna(0)

# df_=df_.reset_index().rename(columns={"index":"filter_"})
# df_lastrow = df_[-1:]
# df_= df_[:-1]

# df_["filter_"]=df_["filter_"].astype(int)

# # df_=df_.loc[(df_[23]!=0) | (df_['filter_']<=23)]
# df_=df_.loc[(df_['filter_']<=80)]


# df_=pd.concat([df_,df_lastrow],ignore_index=True)

# df_=df_.transpose()


# df_=df_[(df_.iloc[:,1].notnull()) & (df_.iloc[:,1]!=0)]

# df_=df_.iloc[1:,np.r_[4,23:len(df_.columns)]]

# df_.columns = [*df_.columns[:-1], 'Project Name']
# df_.columns.values[4]="Emp_Name"
# df_=df_[(df_.iloc[:,1].notnull()) & (df_.iloc[:,1]!=0)]

# df_=df_.iloc[:,np.r_[4,24:len(df_.columns)-1]]

# df_=df_.iloc[:,~df_.columns.isnull()]


# df_.columns.values[1:-1]= pd.to_datetime(df_.columns[1:-1]).strftime('%Y-%m-%d')

# df_=pd.melt(df_,id_vars =["Emp_Name",'Project Name'], value_vars =list(df_.columns[1:-1]))
# df_["value"]=df_["value"].fillna(0)

# df_.to_excel("merged_excel.xlsx")







