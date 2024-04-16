import sys

import pyautogui
import time
import xlrd
import pyperclip
import time
import os

message = "————————————————————————————————————————————————————\n" \
          "操作当中涉及点击鼠标，输入、粘贴数据，这两者组合可能将代码写入脚本当中，需要主要不要造成不良影响！"
print(message)
def dataCheck(sheet1):
    """
    检查给定表格的数据是否符合指定格式。
    参数:
    - sheet1: 一个表格对象，需要对其进行数据格式检查。
    返回值:
    - checkCmd: 布尔值，如果所有数据检查都通过，则为True，否则为False。
    """

    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:
        print("没数据啊哥")
        checkCmd = False
    # 每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0
                                  and cmdType.value != 7.0 and cmdType.value != 8.0 and cmdType.value != 9.0):
            print('第', i + 1, "行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 7.0 热键组合，内容不能为空
        if cmdType.value == 7.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 8.0 时间，内容不能为空
        if cmdType.value == 8.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 9.0 系统命令集模式，内容不能为空
        if cmdType.value == 9.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd


def mouseClick(clickTimes, lOrR, img, reTry):
    """
    自动点击屏幕指定图像的位置。

    参数:
    - clickTimes: 点击次数
    - lOrR: 点击的鼠标按钮，'left' 或 'right'
    - img: 图像文件路径，用于在屏幕上定位点击位置
    - reTry: 重试次数。如果为1，则一直尝试直到找到图像并点击；如果为-1，则一直尝试直到程序结束；如果大于1，则尝试指定次数。
    换言之-单击，长按，双击三击等
    返回值:
    无
    """

    if reTry == 1:
        while True:
            base_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
            print(base_dir)
            img = os.path.join(base_dir, img)
            if os.path.exists(img):
                print("找到匹配图片", img)
            # location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            location = pyautogui.locateCenterOnScreen(img, confidence=0.8)
            print(location)
            if location is not None:
                print("点击", img)
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    # elif reTry == -1:
    #     while True:
    #         location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    #         if location is not None:
    #             pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
    #         time.sleep(0.1)
    # elif reTry > 1:
    #     i = 1
    #     while i < reTry + 1:
    #         location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    #         if location is not None:
    #             pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
    #             print("重复")
    #             i += 1
    #         time.sleep(0.1)


def cmd_action(cmdType, cmdValue=None):
    """
    根据指令类型执行相应的操作。
    """
    if cmdType.value == 0.0:
        time.sleep(cmdValue)
        print("等待完成")
    elif cmdType.value == 1.0:
        # 根据图片单击操作
        mouseClick(1, 'left', cmdValue, 1)
        print("单击完成")
    elif cmdType.value == 2.0:
        # 按键按下“1”
        pyautogui.press('1')
        print("按下1")
    elif cmdType.value == 3.0:
        # 输入密码
        passwd = '11071135Ecust#'
        # 复制密码，粘贴密码
        pyperclip.copy(passwd)
        pyautogui.hotkey('ctrl', 'v')
        print("输入密码完成")
    elif cmdType.value == 4.0:
        # 键入“enter”
        pyautogui.press('enter')
        print("Enter按下")
    elif cmdType.value == 5.0:
        # 键入“enter”
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        print("Tab按下")

def main(sheet1):
    nrows = sheet1.nrows
    for i in range(1,nrows):
        try:
            cmdType = sheet1.row(i)[0]
            cmdValue = sheet1.row(i)[1].value
            cmd_action(cmdType, cmdValue)
            print("第", i, "行指令执行完毕", cmdType.value)
        except:
            print("第", i, "行指令执行失败\n")
            break



if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
    print(base_dir)
    file = os.path.join(base_dir, 'cmd.xls')
    # file = 'cmd.xls'
    # file = 'test.xls'
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    # print(sheet1.row().value)
    main(sheet1)
