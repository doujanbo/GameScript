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



#大图中找小图，并返回各种坐标和相似度
def matchImg(imgobj,num,confidencevalue=0.5):
    imsrc = ac.imread("./screenshot/background.png")
    imobj = ac.imread(imgobj)
    return ac.find_all_template(imsrc, imobj, threshold = num)

#截屏并保存
def screenshot(bmpname):
    im = ImageGrab.grab()
    im.save('./screenshot/' + bmpname + '.png','png')

#输入大图和小图找出小图在大图中的位置和相似度，如果没找到就返回None
def ReturnCoordinates(imgobj, Similarity):

    screenshot("background")
    try:
        return matchImg("./screenshot/" + imgobj + ".png", Similarity)
    except:
        sayhello("没有找到图片")
        return None

def MoveToAdd(x, y):
    '''移动到指定坐标'''

    time.sleep(random.randint(66, 100) * 0.001)
    x = x + random.randint(-2,3)
    y = y + random.randint(-2,3)
    pyautogui.moveTo(x, y, 0.5, pyautogui.easeInQuad)
    time.sleep(random.randint(66, 100) * 0.001)

def highestXY(coordinates):
    '''取出最高处的匹配图片的坐标'''
    i = 0
    xy = coordinates[0]["result"]
    for k in coordinates:
        if k["result"][1] < xy[1]:
            xy = k["result"]
    return xy

def FindBMP(name):
    '''移动到指定图片'''
    i = ReturnCoordinates(name, 0.995)
    if i:
        sayhello("发现此图片")
        return len(i)
    else:
        sayhello("没有发现此图片")
        return False


def MoveToBMP2(name):
    '''移动到指定图片'''

    time.sleep(random.randint(66, 100) * 0.001)
    xy = (0, 0)
    i = ReturnCoordinates(name, 0.95)
    if i:
        xy = highestXY(i)
        x, y = xy
        MoveToAdd(x, y)
        time.sleep(random.randint(66, 100) * 0.001)
        return len(i)
    else:
        return False

def MoveToBMP(name):
    '''移动到指定图片'''

    time.sleep(random.randint(66, 100) * 0.001)
    xy = (0, 0)
    i = ReturnCoordinates(name, 0.995)
    if i:
        xy = highestXY(i)
        x, y = xy
        MoveToAdd(x, y)
        time.sleep(random.randint(66, 100) * 0.001)
        return len(i)
    else:
        return False


def locking():
    '''锁定矿石'''
    xy = (0, 0)
    xy = highestXY(ReturnCoordinates("xiaoxinxing", 0.995))
    x, y = xy
    MoveToAdd(x, y)  # 移动到小行星的位置
    pyautogui.click(button='left')  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    pyautogui.press('ctrl')

    MoveToAdd(x, y + 16)
    pyautogui.click(button='left')  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    pyautogui.press('ctrl')

    MoveToAdd(x, y + 32)
    pyautogui.click(button='left')   # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    pyautogui.press('ctrl')

    MoveToAdd(x, y + 48)
    pyautogui.click(button='left')   # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    pyautogui.press('ctrl')

    MoveToAdd(x, y + 64)
    pyautogui.click(button='left')  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    pyautogui.press('ctrl')

    MoveToAdd(x, y + 80)
    pyautogui.click(button='left')  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    pyautogui.press('ctrl')
    time.sleep(random.randint(66, 100) * 0.001)

engine = pyttsx3.init()
def sayhello(strhello):
    pass
'''
    engine.say(strhello)
    engine.runAndWait()
    '''
def sayhello2(strhello):
    pass
'''
    engine.say(strhello)
    engine.runAndWait()
'''

def open_gun():
    if MoveToBMP("gun"):
        pyautogui.click(button='left')  # 左健单击

        sayhello2("打开矿枪成功")
        return True
    else:
        sayhello2("打开矿枪失败")

        return False





time.sleep(3)

MoveToBMP("eve")
pyautogui.click(button='left')
time.sleep(0.5)

while True:
    #如果有锁定的矿，没有正在挖矿
    if FindBMP("mineral") and (FindBMP("mining") is not True):
        sayhello2("打开矿枪")
        open_gun()
        time.sleep(1)
        open_gun()
    #有锁定的矿并且正在挖矿
    if FindBMP("mineral") and FindBMP("mining"):
        sayhello2("正在挖矿")
        time.sleep(3)
    #如果只有一个锁定的矿
    if FindBMP("mineral") < 2:
        sayhello2("锁定")
        locking()

    if MoveToBMP("Orebin"):
        pyautogui.click(button='left')
        if MoveToBMP2("Warehouseore"):
            sayhello2("转移矿石")
            pyautogui.mouseDown(button='left')
            time.sleep(random.randint(66, 100) * 0.001)
            MoveToBMP("fleet")
            pyautogui.mouseUp(button='left')
            MoveToBMP("666666")
        else:
            MoveToBMP("666666")



print("hello world")
input()