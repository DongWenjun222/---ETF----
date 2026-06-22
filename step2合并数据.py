import pandas as pd
import numpy as np
from sklearn.metrics import mean_squared_error
import matplotlib.pyplot as plt
from tqdm import tqdm
from scipy.optimize import minimize
import datetime as dt
from WindPy import w
w.start()
plt.rcParams['font.sans-serif'] = ['SimHei'] # 设置字体为黑体
plt.rcParams['axes.unicode_minus'] = False # 解决坐标轴负号显示问题
import warnings
warnings.filterwarnings("ignore")

def merge_data():
    data_AAplus_更新 = pd.read_excel('期限结构_中债AA+_更新.xlsx')
    data_AAplus_历史 = pd.read_excel('期限结构_中债AA+_历史.xlsx')
    data_AAA_更新 = pd.read_excel('期限结构_中债AAA_更新.xlsx')
    data_AAA_历史 = pd.read_excel('期限结构_中债AAA_历史.xlsx')

    ## 合并 data_AAA_历史,data_AAA_更新
    for column in data_AAA_更新.columns[2:]:
        #if column not in data_AAA_历史.columns:
        data_AAA_历史[column] = data_AAA_更新[column]  

    for column in data_AAplus_更新.columns[2:]:
        #if column not in data_AAplus_历史.columns:
        data_AAplus_历史[column] = data_AAplus_更新[column]  
    
    data_AAA_历史.to_excel('期限结构_中债AAA_历史.xlsx',index=False)
    data_AAplus_历史.to_excel('期限结构_中债AA+_历史.xlsx',index=False)
    return None
if __name__ == "__main__":
    merge_data()

