from pynput import keyboard
from PIL import ImageGrab, Image
import win32api
import win32gui
import win32con
import compare_image
import numpy
from queue import Queue
import pyautogui
import time
import tkinter as tk

############################## 全局参数 ##############################
window_title = "QQ游戏 - 连连看角色版"  # 游戏窗口名称
imgpath = "C:/Users/luadmin/Desktop/llk/testimg/"  # 截图保存位置
window_width = 800 #游戏窗口宽度
window_height = 600 #游戏窗口高度
game_width = 589  # 游戏区域宽度
game_height = 385  # 游戏区域高度
num_row = 11  # 11行
num_col = 19  # 19列
grid_width = 31  # 方块宽度
grid_height = 35  # 方块高度
color_empty_grid = (48, 76, 112) # 空方块RGB
screen_width = win32api.GetSystemMetrics(0)  # 屏幕分辨率宽
screen_height = win32api.GetSystemMetrics(1)  # 屏幕分辨率高
dx = [0, 1, 0, -1] # bfs方向
dy = [1, 0, -1, 0]
di = [1, 2, 3, 4] # 标记四个方向的序号
resq = Queue() # 消块队列
game_left = 0 # 游戏区域参数
game_top = 0
game_right = 0
game_bottom = 0
chonglie_x = 0 # 重列道具坐标
chonglie_y = 0
click_interval = 0 # 点击频率
############################## 全局参数 ##############################

def start():
    global game_left
    global game_top
    global game_right
    global game_bottom
    global chonglie_x
    global chonglie_y
    global click_interval
    resq.queue.clear()
    try:
        click_interval = float(entry.get())
    except:
        win32api.MessageBox(0, "输入有误,置为0", "ERROR", win32con.MB_OK)
        click_interval = 0
    hwnd = win32gui.FindWindow(0, window_title)  # 获取游戏窗口句柄
    if hwnd == 0:
        print("no find game hwnd")
    else:
        window_left, window_top, window_right, window_bottom = win32gui.GetWindowRect(hwnd)  # 获取游戏窗口坐标
        #如果游戏窗口超出屏幕范围，将窗口移动至屏幕中间
        if window_left < 0 or window_top < 0 or window_right > screen_width or window_bottom > screen_height:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, int(screen_width / 2 - window_width / 2), int(screen_height / 2 - window_height / 2), 100, 100, win32con.SWP_NOSIZE)
            window_left, window_top, window_right, window_bottom = win32gui.GetWindowRect(hwnd)
        # win32gui.SetForegroundWindow(hwnd)  # 激活窗口至前端
        
        #点击激活游戏窗口
        pyautogui.click(int(window_left + 0.5 * window_width), window_top + 20, button="left", interval=0)

        #计算游戏棋盘区域
        game_left = window_left + 14
        game_top = window_top + 181
        game_right = game_left + game_width
        game_bottom = game_top + game_height

        #计算重列道具位置
        chonglie_x = window_left + 656
        chonglie_y = window_top + 201

        #截取游戏区域图片
        img = getGameImage()

        #将游戏区域图片分成11 * 19并得到最终棋盘方块矩阵信息
        matrix = gameImageToMatrix(img)

        solve(matrix)

def solve(matrix):
    last_solve = -1
    solvable = True
    while True:
        solve = 0
        for i in range(num_row):
            for j in range(num_col):
                solve += bfs(matrix, i, j)
        if solve == num_row * num_col:
            break
        elif last_solve != -1 and solve == last_solve:  # 如果一轮遍历找不到能解决的方块，则无解，考虑使用重列道具
            solvable = False
            break
        else:
            last_solve = solve
    print(solve, resq.qsize())
    
    while not resq.empty():
        solveOneStep()
    if solvable == False:
        unsolvable()

def unsolvable():
    pyautogui.click(chonglie_x, chonglie_y, button="left", interval = 0) # 使用重列道具

    #等待游戏界面的无解提示消失，如果减少sleep时间将影响图像识别导致部分方块无法消除
    time.sleep(8)

    #截取游戏区域图片
    img = getGameImage()

    #将游戏区域图片分成11 * 19并得到最终棋盘方块矩阵信息
    matrix = gameImageToMatrix(img)

    solve(matrix)

def solveOneStep():
    if not resq.empty():
        cur = resq.get()
        print(cur)
        from_x = game_left + (cur[1] + 0.5) * grid_width
        from_y = game_top + (cur[0] + 0.5) * grid_height
        to_x = game_left + (cur[3] + 0.5) * grid_width
        to_y = game_top + (cur[2] + 0.5) * grid_height

        pyautogui.click(from_x, from_y, button = "left", interval = click_interval)
        pyautogui.click(to_x, to_y, button = "left", interval = click_interval)
    else:
        print("over")

def judge(x,y,count,visit,matrix,res):
    if x < 0 or y < 0 or x >= num_row or y >= num_col:  # 边界
        return False
    if count > 2: # 转向次数超过两次 剪枝
        return False
    if visit[x][y] == 1:
        return False
    if matrix[x][y] != 0 and matrix[x][y] != res:
        return False
    return True

def bfs(matrix,x,y):
    if matrix[x][y] == 0:# 如果坐标上不是方块则跳过
        return 1
    visit = numpy.zeros((num_row,num_col))
    visit[x][y] = 1
    q = Queue()
    q.put([x,y,-1,0,visit]) # 1.方块坐标x 2.方块坐标y 3.上一步方向 -1为起始 4.转向次数 5.坐标访问记录数组
    while not q.empty():
        cur = q.get()
        if (cur[0] != x or cur[1] != y) and matrix[cur[0]][cur[1]] == matrix[x][y]:
            matrix[x][y] = 0 # 如果不需要打印过程，可以将此处修改为 0，并将 printMatrix()删除
            matrix[cur[0]][cur[1]] = 0 # 如果不需要打印过程，可以将此处修改为 0，并将 printMatrix()删除
            resq.put([x,y,cur[0],cur[1]])
            return 0
        for i in range(4):
            nx = cur[0] + dx[i]
            ny = cur[1] + dy[i]
            nd = di[i]
            ncount = cur[3]
            nvisit = cur[4].copy()
            if cur[2] != -1 and cur[2] != nd: # 新方向与旧方向不同 增加转向次数
                ncount += 1
            if judge(nx, ny, ncount, nvisit, matrix, matrix[x][y]):
                nvisit[nx][ny] = 1
                q.put([nx,ny,nd,ncount,nvisit])
    return 0

def gameImageToMatrix(img):
    image_matrix = {}
    image_matrix_id = {}
    unknown_type_image = []

    for row in range(num_row): # 11
        image_matrix[row] = {}
        for col in range(num_col): # 19
            left = col * grid_width
            top = row * grid_height
            right = left + grid_width
            bottom = top + grid_height

            # 分割 存储 方块 image
            grid_image = img.crop((left,top,right,bottom))
            image_matrix[row][col] = grid_image
            
    for row in range(num_row):
        image_matrix_id[row] = {}
        for col in range(num_col):
            this_img = image_matrix[row][col]
            if isEmptyGrid(this_img):
                image_matrix_id[row][col] = 0 # 无方块标记为0
                continue
            found = False
            for i in range(len(unknown_type_image)):
                # 与已遍历过的未知方块对比 判断是否为一对方块
                if isSameGrid(this_img, unknown_type_image[i]):
                    found = True
                    id = i + 1
                    image_matrix_id[row][col] = id
                    break
            if not found:
                # 将未知方块加入列表 等待下次对比
                unknown_type_image.append(this_img)
                id = len(unknown_type_image)
                image_matrix_id[row][col] = id
    return image_matrix_id

def isSameGrid(img_a, img_b):
    numpy_array_a = numpy.array(img_a)
    numpy_array_b = numpy.array(img_b)
    if 0.95 < compare_image.classify_hist_with_split(numpy_array_a, numpy_array_b, size=img_a.size):
        return True
    return False

def isEmptyGrid(img):
    # 将方块图片64等分，取中心4份，相当于将图片缩放十六分之一
    center = img.resize(size=(int(img.width / 8.0), int(img.height / 8.0)),
                        box=(int(img.width / 8.0 * 3.0), int(img.height / 8.0 * 3.0), int(img.width / 8.0 * 5.0), int(img.height / 8.0 * 5.0)))
    # 根据中心部分图像RGB判断是否为背景颜色以判断是否为方块
    for color in center.getdata():
        if color != color_empty_grid:
            return False
    return True

def getGameImage():
    try:
        size = (game_left, game_top, game_right, game_bottom)
        img = ImageGrab.grab(size)
        return img
    except Exception as e:
        print("error:", e)
        exit(-1)


GUI = tk.Tk()
GUI.title("连连看自动玩")

GUI_width = 250
GUI_height = 100
print(type(screen_width))
alignstr = '%dx%d+%d+%d' % (GUI_width, GUI_height, (screen_width - GUI_width) / 8, (screen_height - GUI_height) / 4)
GUI.geometry(alignstr)
GUI.resizable(width = False, height = False)

tk.Button(GUI, text="0.0", font=("宋体", 12), command=lambda:setEntryValue(0.0)).place(width=30, height=30, x=10, y=55)
tk.Button(GUI, text="0.1", font=("宋体", 12), command=lambda:setEntryValue(0.1)).place(width=30, height=30, x=10, y=15)
tk.Button(GUI, text="0.2", font=("宋体", 12), command=lambda:setEntryValue(0.2)).place(width=30, height=30, x=40, y=15)
tk.Button(GUI, text="0.3", font=("宋体", 12), command=lambda:setEntryValue(0.3)).place(width=30, height=30, x=70, y=15)
tk.Button(GUI, text="0.4", font=("宋体", 12), command=lambda:setEntryValue(0.4)).place(width=30, height=30, x=100, y=15)

tk.Label(GUI, text="间隔", font=("宋体", 12)).place(width=40, height=30, x=45, y=55)
value = tk.StringVar(value="0.0")
entry = tk.Entry(GUI, show=None, font=("宋体", 12),insertontime=1000, textvariable=value)
entry.place(width=40, height=30, x=85, y=55)
tk.Button(GUI, text="开始", font=("宋体", 12), command=start).place(width=100, height=70, x=140, y=15)

def setEntryValue(v):
    entry.delete(0, "end")
    entry.insert(0, v)

GUI.mainloop()
