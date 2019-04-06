import cv2
import numpy as np
from time import time, sleep
from mss import mss
import win32api
import win32con
from win32gui import GetWindowText, GetForegroundWindow, GetWindowRect

def image_search(image, templ, precision):
    res = cv2.matchTemplate(image, templ, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val < precision:
        return [-1, -1]
    return max_loc
    '''
    y, x = np.unravel_index(res.argmax(), res.shape)
    return x, y
    '''

def image_click(image, templ, precision=0.8):
    loc = image_search(image, templ, precision)
    if loc == [-1, -1]:
        return False
    height, width = templ.shape
    x, y = loc
    x += width // 2
    y += height // 2
    mouse_click(x, y)
    return True

def mouse_click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

if __name__ == '__main__':
    # cv2.namedWindow('Frame', cv2.WINDOW_NORMAL)
    # cv2.resizeWindow('Frame',  1366, 768)
    # cv2.moveWindow('Frame', 100, 100)
    with mss() as sct:
        while True:
            try:
                # hWnd = GetForegroundWindow()
                # text, rect = GetWindowText(hWnd), GetWindowRect(hWnd)
                # pywintypes.error: (1400, 'GetWindowRect', '잘못된 창 핸들입니다.')
                '''
                x, y = rect[:2]
                w, h = [y - x for x, y in zip(rect[:2], rect[2:])]
                print('Window({})'.format(text))
                print('Location: ({}, {})'.format(x, y))
                print('Size: ({}, {})'.format(w, h))
                '''
                # if text == 'Rainbow Six':
                #     pass
                frame = np.asarray(sct.grab(sct.monitors[1]))
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                
                for image_index in range(1, 6):
                    target = cv2.imread('./resources/{}.jpg'.format(image_index))
                    target = cv2.cvtColor(target, cv2.COLOR_BGR2GRAY)
                    result = image_click(frame, target)
                    if result:
                        # cv2.rectangle(img, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)
                        print('{}번째 이미지를 성공적으로 클릭했습니다.'.format(image_index))
                    
                # cv2.imshow('Frame', frame)
                sleep(0.01)
                # if cv2.waitKey(25) & 0xFF == ord('q'):
                #     cv2.destroyAllWindows()
                #     break
            except KeyboardInterrupt:
                break
