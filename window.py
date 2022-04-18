import win32gui, win32ui, win32con
from PIL import Image
import numpy as np

class Window():

    def __init__(self, hwnd):
        self.hwnd = hwnd
        self.left, self.top, self.right, self.bot = win32gui.GetWindowRect(self.hwnd)
        
        self.top += 140
        self.right -= 40
        self.left += 15
        self.bot -= 10

        self.w = self.right - self.left
        self.h = self.bot - self.top

    def get_screenshot(self):
        hdesktop = win32gui.GetDesktopWindow()
        hwndDC = win32gui.GetWindowDC(hdesktop)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, self.w, self.h)

        saveDC.SelectObject(saveBitMap)

        result = saveDC.BitBlt((0, 0), (self.w, self.h), mfcDC, (self.left, self.top), win32con.SRCCOPY)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        im = np.ascontiguousarray(im)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hdesktop, hwndDC)
        return im