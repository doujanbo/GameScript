import cv2
import numpy
import aircv as ac
import os
import sys
import time
from ctypes import *
import win32com
import win32api
import win32com.client
import json
from PIL import ImageGrab
import random
import pyttsx3
import pyautogui

#进程内注册插件,模块所在的路径按照实际位置修改
#hkmdll = windll.LoadLibrary("D:/GameScript/mouse/wyhkm.dll")

#创建对象
#try:
wyhkm=win32com.client.Dispatch("wyp.hkm")
#except:
#print("创建对象失败!")
#sys.exit(0)
#获得模块版本号
version=wyhkm.GetVersion()
print("无涯键鼠盒子模块版本："+hex(version))
#查找设备,这个只是例子,参数中的VID和PID要改成实际值

DevId=wyhkm.SearchDevice(0x2612, 0x1701, 0)

if DevId==-1:
    print("未找到无涯键鼠盒子")
    sys.exit(0)
#打开设备
if not wyhkm.Open(DevId):
    print("打开无涯键鼠盒子失败")
    sys.exit(0)
#打开资源管理器快捷键Win+E


#大图中找小图，并返回各种坐标和相似度
def matchImg(imgobj,num,confidencevalue=0.5):
    imsrc = ac.imread("./screenshot/background.bmp")
    imobj = ac.imread(imgobj)
    return ac.find_all_template(imsrc, imobj, threshold = num)

#截屏并保存
def screenshot(bmpname):
    im = ImageGrab.grab()
    im.save('./screenshot/' + bmpname + '.bmp','bmp')

#输入大图和小图找出小图在大图中的位置和相似度，如果没找到就返回None
def ReturnCoordinates(imgobj, Similarity):

    screenshot("background")
    try:
        return matchImg("./screenshot/" + imgobj + ".bmp", Similarity)
    except:
        print("没找到")
        return None

def MoveToAdd(x, y):

    time.sleep(random.randint(66, 100) * 0.001)
    x = x + random.randint(-2,3)
    y = y + random.randint(-2,3)
    wyhkm.MoveTo(x, y)
    time.sleep(random.randint(66, 100) * 0.001)

def highestXY(coordinates):
    i = 0
    xy = coordinates[0]["result"]
    for k in coordinates:
        if k["result"][1] < xy[1]:
            xy = k["result"]
    return xy

def MoveToBMP(name):
    xy = (0, 0)
    xy = highestXY(ReturnCoordinates(name, 0.999))
    x, y = xy
    MoveToAdd(x, y)  # 移动到小行星的位置

def locking(name):

    xy = (0, 0)
    xy = highestXY(ReturnCoordinates(name, 0.999))
    x, y = xy
    MoveToAdd(x, y)  # 移动到小行星的位置
    wyhkm.LeftClick()  # 左健单击
    wyhkm.Keypress('Ctrl')

engine = pyttsx3.init()
def sayhello(strhello):
    pass

    engine.say(strhello)
    engine.runAndWait()


time.sleep(3)
#locking("xiaoxinxing")
MoveToBMP("xiaoxinxing")
sayhello("运行完毕")

print("hello world")
input()