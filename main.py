import cv2
import numpy
import aircv as ac
import os
import sys
import time
from ctypes import *
import win32com
import win32api
import win32gui
import win32ui
import win32con
import win32com.client
import json
from PIL import ImageGrab
import random
import pyttsx3
import pyautogui
from ctypes import windll
from PIL import Image

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
    x = x + random.randint(-3,4)
    y = y + random.randint(-3,4)
    wyhkm.MoveTo(x, y)
    time.sleep(random.randint(66, 100) * 0.001)

def highestXY(coordinates):
    '''取出最高处的匹配图片的坐标'''
    i = 0
    try:
        xy = coordinates[0]["result"]
        for k in coordinates:
            if k["result"][1] < xy[1]:
                xy = k["result"]
        return xy
    except:
        print(coordinates)
        return False

def FindBMP2(name):
    '''移动到指定图片'''
    i = ReturnCoordinates(name, 0.95)
    if i:
        sayhello("发现此图片")
        return len(i)
    else:
        sayhello("没有发现此图片")
        return False

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
    try:
        x, y = xy
    except:
        return False
    MoveToAdd(x, y)  # 移动到小行星的位置
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    MoveToAdd(x, y + 16)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    MoveToAdd(x, y + 32)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    MoveToAdd(x, y + 48)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    MoveToAdd(x, y + 64)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')

    MoveToAdd(x, y + 80)
    wyhkm.LeftClick()  # 左健单击
    time.sleep(random.randint(66, 100) * 0.001)
    wyhkm.Keypress('Ctrl')
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
        wyhkm.LeftClick()  # 左健单击

        sayhello2("打开矿枪成功")
        return True
    else:
        sayhello2("打开矿枪失败")

        return False




'''
time.sleep(3)
MoveToBMP("eve")
wyhkm.LeftClick()
time.sleep(0.5)
'''

def find_xiaoxinxing():



    MoveToAdd(456, 298)
    wyhkm.RightClick()
    time.sleep(0.5)
    if MoveToBMP("xiaodai"):
        time.sleep(0.5)
        if MoveToBMP("xiaoxinxingdai"):
            while FindBMP("yueqian"):
                wyhkm.MoveRP(0,13)
                time.sleep(0.5)
            while bool(FindBMP("yueqian")) is not True:
                wyhkm.MoveRP(0, 4)
                time.sleep(0.2)
            if MoveToBMP("yueqian"):
                wyhkm.LeftClick()

def showtext():
    titles = set()

    def foo(hwnd, mouse):
        # 去掉下面这句就所有都输出了，但是我不需要那么多
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            titles.add(win32gui.GetWindowText(hwnd))

    win32gui.EnumWindows(foo, 0)
    lt = [t for t in titles if t]
    lt.sort()
    for t in lt:
        print(t)




def calculating_coordinates(win_name):
    hwnd = win32gui.FindWindow(0, win_name)  # 父句柄
    windowRec = win32gui.GetWindowRect(hwnd)  # 目标子句柄窗口的坐标

    while True:
        tempt = win32api.GetCursorPos()  # 记录鼠标所处位置的坐标
        x = tempt[0] - windowRec[0]  # 计算相对x坐标
        y = tempt[1] - windowRec[1]  # 计算相对y坐标
        print(x, y)
        time.sleep(0.5)  # 每0.5s输出一次

def save_bmp(win_name):
    # 获取要截取窗口的句柄
    hwnd = win32gui.FindWindow(None, win_name)


    # Change the line below depending on whether you want the whole window
    # or just the client area.
    # left, top, right, bot = win32gui.GetClientRect(hwnd)
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bot - top

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)

    # Change the line below depending on whether you want the whole window
    # or just the client area.
    # result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)
    print
    result

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    if result == 1:
        # PrintWindow Succeeded
        im.save("test.png")
        im.show()


def fun():
    hwnd = win32gui.FindWindow(0,"星战前夜：晨曦 - 窦逍遥")
    print(win32gui.GetWindowText(hwnd))
    x = 1378
    y = 266
    x = x + random.randint(-3, 4)
    y = y + random.randint(-3, 4)
    long_position = win32api.MAKELONG(x, y)  # 模拟鼠标指针 传送到指定坐标
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)  # 模拟鼠标按下
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)  # 模拟鼠标弹起

#save_bmp("新建文本文档.txt - 记事本")
#calculating_coordinates("星战前夜：晨曦 - 窦逍遥")

#fun()
#showtext()






'''
#逆戟鲸
while True:
    if FindBMP2("mineral") < 2:
        locking()
    if MoveToBMP("xiaoxinxing"):
        wyhkm.LeftClick()
        if MoveToBMP("jiejin"):
            wyhkm.LeftClick()
            time.sleep(5)
    if FindBMP("konxian"):
        if MoveToBMP("kuangwu"):
            wyhkm.RightClick()
            if MoveToBMP("kaicai"):
                wyhkm.LeftClick()
                time.sleep(3)
                MoveToAdd(1440, 900)
'''

print("hello world")
input()