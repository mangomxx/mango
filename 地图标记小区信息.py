# -*- coding: utf-8 -*-
"""
Created on  AUG 15  2019

@author:芒果果哟啊
"""

from __future__ import division
import pandas as pd
import folium
from folium import plugins
import webbrowser
import math
# posi=pd.read_csv("C:\Users\18810\Desktop\Test\\2015Cities-CHINA.csv")

posi=pd.read_excel("E:\工作\python实现给经纪人发个人报告\Test\\小区名称极坐标.xlsx")


# 创建以高德地图为底图的密度图：
def map ():
    map_osm = folium.Map(
     location=[39.99043611,116.4467194],   #地图打开的中心位置
     zoom_start=15,              #放大程度  数值越大越详细
     control_scale=True,   #添加比例尺
     no_touch=True,#closePopupOnClick=False,
     tiles='http://webrd02.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}',
     attr="&copy; <a href='http://ditu.amap.com/'>高德地图</a>"
     )
    a=plugins.MarkerCluster().add_to(map_osm)
    return a


def item (map):
    for name,row in posi.iterrows():
      # icon = folium.Icon(color='red', icon='house')
      tooltip = folium.Tooltip("{0}:{1}".format(row["小区名称"], row["房屋总数"]),
                               sticky=False,permanent=True)
                               # ,direction="left")
      popup = folium.Popup("{0}:{1}".format(row["小区名称"], row["房屋总数"]),
                           parse_html=True, max_width=2650,show=False
                           # auto_pan=False,autoClose=False
                           )
      bd_lng=row["经度"]
      bd_lat=row["纬度"]
      X_PI = 3.14159265358979324 * 3000.0 / 180.0
      x = bd_lng - 0.0065
      y = bd_lat - 0.006
      z = math.sqrt(x * x + y * y) - 0.00002 * math.sin(y * X_PI)
      theta = math.atan2(y, x) - 0.000003 * math.cos(x * X_PI)
      lng = z * math.cos(theta)
      lat = z * math.sin(theta)
      b=folium.CircleMarker([lat, lng],  fill=True,fill_color='#6495ED' ,stroke=False,
                          fillOpacity=0.5,  tooltip=tooltip).add_to(map)
    # radius = 100,
    return b
if __name__ == "__main__":
    map1=map()
    map2=item(map1)
    file_path = r"C:\Users\18810\Desktop\小区1.html"
    map2.save(file_path)     # 保存为html文件
    webbrowser.open(file_path)  # 默认浏览器打开


