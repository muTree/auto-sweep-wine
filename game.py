from PIL import ImageGrab, Image
import time
import win32gui, win32api, win32con
import random
from win_mouse import *
debug = False
from copy import deepcopy

class sl:
    class_name = "TMain"
    title_name = "Minesweeper Arbiter "

    blocks_img = 0
    left, top, right, bottom = 0, 0, 0, 0
    x, y = 0, 0
    imgs = [0]*13
    gameover = False

    pivots = [[(8, 1),(192, 192, 192)], #0
              [(8, 8),(0, 0, 255)],     #1
              [(8, 8),(0, 128, 0)],     #2
              [(8, 8),(255, 0, 0)],     #3
              [(8, 8),(0, 0, 128)],     #4
              [(8, 8),(128, 0, 0)],     #5
              [(8, 8),(0, 128, 128)],   #6
              [(3, 3),(0, 0, 0)],       #7
              [(8, 8),(128, 128, 128)], #8
              [(4, 11),(0, 0, 0)],      #9
              [(8, 1),(255, 255, 255)], #10
              [(0, 0),(0, 0, 0)],       #11
              [(3, 3),(255, 0, 0)],     #12
              ]
    
    def __init__(self):
        pass
        #print("Intializing...")
        #for i in range(13):
            #self.imgs[i] = Image.open("sources/" + str(i) + ".bmp")
    def getWindowInfo(self):
        if debug:
            print("get window info")
        hwnd = win32gui.FindWindow(self.class_name, self.title_name)
        if hwnd:
            self.left, self.top, self.right, self.bottom = win32gui.GetWindowRect(hwnd)

        self.left += 15
        self.top += 101
        self.right -= 15
        self.bottom -=43
        self.x, self.y = int((self.right-self.left)/16), int((self.bottom-self.top)/16)
        if debug:
            print("x: %d" % self.x)
            print("y: %d" % self.y)
        self.blocks_img = [[0 for i in range(self.y)] for i in range(self.x)]
        return True

    def getcolor(self, img, xy, dxy):
        return img.getpixel((xy[0]*16+dxy[0], xy[1]*16+dxy[1]))
    def analyze_block(self):
        if debug:
            print("analyze...")
        screen = ImageGrab.grab((self.left, self.top, self.right, self.bottom))
        for j in range(self.y):
            for i in range(self.x):
                stat = False
                for index in (1, 2, 3, 4, 5, 6, 7, 8, 12):
                    if self.getcolor(screen, (i, j), self.pivots[index][0]) == self.pivots[index][1]:
                        self.blocks_img[i][j] = index
                        stat = True
                        break
                if not stat:
                    if self.getcolor(screen, (i, j), (8, 0)) == (128, 128, 128) \
                        and self.getcolor(screen, (i, j), (8, 15)) == (192, 192, 192):
                        self.blocks_img[i][j] = 0
                    elif self.getcolor(screen, (i, j), (1, 1)) == (255, 255, 255) \
                        and self.getcolor(screen, (i, j), (7, 7)) == (255, 0, 0):
                        self.blocks_img[i][j] = 9
                    elif self.getcolor(screen, (i, j), (8, 0)) == (255, 255, 255) \
                        and self.getcolor(screen, (i, j), (8, 15)) == (128, 128, 128):
                        self.blocks_img[i][j] = 10
                    elif self.getcolor(screen, (i, j), (1, 1)) == (192, 192, 192) \
                        and self.getcolor(screen, (i, j), (7, 7)) == (255, 255, 255):
                        self.blocks_img[i][j] = 11      

    def left_click(self, xy):
        if 0:
            print("left click: ", xy)
        if isinstance(xy, tuple):
            x = xy[0] * 16 + self.left + 8
            y = xy[1] * 16 + self.top + 8
            mouse_left_click(x, y)
            return True
        for (x, y) in xy:
            x = x * 16 + self.left + 8
            y = y * 16 + self.top + 8
            mouse_left_click(x, y)
    def right_click(self, xy):
        if 0:
            print("right click: ", xy)
        if isinstance(xy, tuple):
            x = xy[0] * 16 + self.left + 8
            y = xy[1] * 16 + self.top + 8
            mouse_right_click(x, y)
            return True
        for (x, y) in xy:
            x = x * 16 + self.left + 8
            y = y * 16 + self.top + 8
            mouse_right_click(x, y)
    def clear_click(self, xy):
        if 0:
            print("clear click: ", xy)
        if isinstance(xy, tuple):
            x = xy[0] * 16 + self.left + 8
            y = xy[1] * 16 + self.top + 8
            mouse_left_right_click(x, y)
            return True
        for (x, y) in xy:
            x = x * 16 + self.left + 8
            y = y * 16 + self.top + 8
            mouse_left_right_click(x, y)
        
    
    def printf(self, im):
        for j in range(self.y):
            for i in range(self.x):
                print(im[i][j], ' ', end='')
            print("")
    def printf2(self, im):
        for j in range(self.y):
            for i in range(self.x):
                print(im[i][j], ' ', end='')
            print("")

    def gameOver(self):
        if self.gameover:
            print("Over")
            return True
        for i in self.blocks_img:
            for j in i:
                if j == 11 or j == 12:
                    print("Over")
                    return True
        return False

    def around(self, x, y, bimg):
        b9 = []
        b10 = []
        for i in (-1, 0, 1):
            for j in (-1, 0, 1):
                xx = x + i
                yy = y + j
                if not 0 <= xx <= self.x-1:
                    continue
                if not 0 <= yy <= self.y-1:
                    continue
                if bimg[xx][yy] == 9:
                    b9.append((xx, yy))
                if bimg[xx][yy] == 10:
                    b10.append((xx, yy))
        return b9, b10

    def getallb10(self):
        tmp = []
        for i in range(self.x):
            for j in range(self.y):
                if self.blocks_img[i][j] == 10:
                    tmp.append((i, j))
        return tmp

    def chongtu(self, im):
        img = deepcopy(im)
        for j in range(self.y):
            for i in range(self.x):
                bnum = img[i][j]
                if bnum in range(1, 9):
                    blocks9, blocks10 = self.around(i, j, img)
                    l9, l10 = len(blocks9), len(blocks10)
                    if l9 > bnum or l9 + l10 < bnum:
                        if 0:
                            self.printf2(img)
                            print("chongtu ", i, j)
                            print("bnum:", bnum, "l9:", l9, "l10:", l10)
                            print(blocks9, blocks10)
                        return True
                    elif l9 == bnum and l10 > 0:
                        for xy in blocks10:
                            img[xy[0]][xy[1]] = 0
                        return self.chongtu(img)
                    elif l9 + l10 == bnum and l10 > 0:
                        for xy in blocks10:
                            img[xy[0]][xy[1]] = 9
                        return self.chongtu(img)
        return False


    def guessBlock(self):
        for i in range(self.x):
            for j in range(self.y):
                if self.blocks_img[i][j] == 10:
                    blocks_img_copy = deepcopy(self.blocks_img)
                    #print(blocks_img_copy)
                    if debug:
                        print("假设", i, j, "无雷")
                    blocks_img_copy[i][j] = 0
                    if self.chongtu(blocks_img_copy):
                        if debug:
                            print("冲突 right click", i, j)
                        self.right_click((i, j))
                        return True
                    if debug:
                        print("假设", i, j, "有雷")
                    blocks_img_copy = deepcopy(self.blocks_img)
                    #print(blocks_img_copy)
                    blocks_img_copy[i][j] = 9
                    #print(blocks_img_copy)
                    if self.chongtu(blocks_img_copy):
                        if debug:
                            print("冲突 left click", i, j)
                        self.left_click((i, j))
                        return True
                    if debug:
                        print("不冲突")
        return False

        
    def whatsNext(self):
        for j in range(self.y):
            for i in range(self.x):
                bnum = self.blocks_img[i][j]
                if bnum == 11 or bnum == 12:
                    print("Game Over")
                    gameover = True
                    return False
                if bnum in range(1, 9):
                    blocks9, blocks10 = self.around(i, j, self.blocks_img)
                    l9, l10 = len(blocks9), len(blocks10)
                    if l9 == bnum:
                        if l10 == 0:
                            continue
                        else:
                            self.clear_click((i, j))
                            return True
                    elif l9 >0:
                        if l9+l10 == bnum:
                            self.right_click(blocks10)
                            return True
                        else:
                            continue
                    elif l9 == 0:
                        if l10 == bnum:
                            self.right_click(blocks10)
                            return True
                        else:
                            continue
        if self.guessBlock():
            return True
        #随机选择
        allb10 = self.getallb10()
        length = len(allb10)
        pos = allb10[random.randint(0, length-1)]
        print("random click: ", pos)
        self.left_click(pos)
        return True
                        
    def play(self):
        self.getWindowInfo()
        self.analyze_block()
        while not self.gameOver():
            self.whatsNext()
            self.analyze_block()
if __name__ == '__main__':
    obj = sl()
    obj.play()
    if False:
        for j in range(obj.y):
            for i in range(obj.x):
                print(obj.blocks_img[j][i], ' ', end="")
            print("")
