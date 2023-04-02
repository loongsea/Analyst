import streamlit as st
import pandas as pd
import numpy as np
import time
import datetime
import random

st.set_page_config(
    page_title="初中成绩分析程序",    #页面标题
    page_icon=":rainbow:",        #icon
    # layout="wide",                #页面布局
    initial_sidebar_state="auto"  #侧边栏
)

sidebar = st.sidebar.radio(
    "导航栏",
    ("首页", "项目管理", "用户管理", "权限管理")
)

st.title("人力资源评分系统")
html_temp = """
<div style ="backgroud-color:tomato;padding:10px">
<h2 style="color:white;text-align:center;">人力</h2>
</div>
"""











