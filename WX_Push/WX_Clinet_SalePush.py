import time
import sys
import os
import stomp
import json
import urllib3
import os
import time
from ctypes import *
from PIL import Image
import win32gui
import win32con
import win32clipboard
import win32api

import logging
import logging.handlers
from queue import Queue
import win32clipboard as w
from  stomp import ConnectionListener
from MyLocalMQ2 import Message_Contrl
import threading
import copy
from concurrent.futures import ThreadPoolExecutor

user = os.getenv("ACTIVEMQ_USER") or "admin"
password = os.getenv("ACTIVEMQ_PASSWORD") or "password"
host = os.getenv("ACTIVEMQ_HOST") or "192.168.10.212"#192.168.10.212localhost
port = os.getenv("ACTIVEMQ_PORT") or 61613
destination = sys.argv[1:2] or ["/topic/Event"]
destination = destination[0]

class MyListener2(ConnectionListener):

  def __init__(self, conn,message_contrller):
    self.conn = conn
    self.count = 0
    self._message_contrller = message_contrller
    self.start = time.time()

  def on_error(self, headers, message):
    print('received an error %s' % message)

  def on_message(self, headers, message):

    _headers = headers
    print(headers)

    # self.message_contrller.write(result_final["FILE.NAME"])

    if  _headers["msgtype"] == "blod":

      print(_headers["Url"]+_headers["message-id"].replace(':','_'))

      url=headers["Url"]+_headers["message-id"].replace(':','_')
      filename=headers["filename"]

      http = urllib3.PoolManager()
      r = http.request('GET', url.replace('localhost','localhost'),timeout=urllib3.Timeout(connect=5.0, read=5.0))#192.168.10.212
      if r.status == 200:
        file1='c:/'+filename+'.bmp'
        with open('c:/'+filename+'.bmp', "wb") as code:
          code.write(r.data)
      http.request('DELETE', url.replace('localhost','localhost'),timeout=urllib3.Timeout(connect=5.0, read=5.0))#192.168.10.212
      self._message_contrller.write(file1,headers['department'],'blod')

    if _headers["msgtype"] == "text":
        self._message_contrller.write(message, headers['department'], 'text')





  def on_disconnected(self):
      print("disconnect")
      self.conn.disconnect()
      print("reconnect")
      self.connect_and_subscribe(self.conn)

  def connect_and_subscribe(self, conn):
      conn.start()
      conn.connect(login=user, passcode=password, headers={"client-id": "multi_test_C"})
      conn.subscribe(destination=destination, id="multi_test01", ack='auto',
                     headers={"activemq.subscriptionName": "multi_testC01", "client-id": "multi_test_C01"})
      conn.subscribe(destination="/topic/SaleInfo", id="multi_test02", ack='auto',
                     headers={"activemq.subscriptionName": "multi_testC02", "client-id": "multi_test_C02"})

class RometeMQReceiver:
    def __init__(self,message_contrller):
        self._message_contrller=message_contrller

    def recevie(self):
        conn = stomp.Connection(host_and_ports=[(host, port)])
        conn.set_listener('', MyListener2(conn, self._message_contrller))
        conn.start()
        conn.connect(login=user, passcode=password, headers={"client-id": "multi_test_C"})
        conn.subscribe(destination=destination, id="multi_test01", ack='auto',
                       headers={"activemq.subscriptionName": "multi_testC01", "client-id": "multi_test_C01"})
        conn.subscribe(destination="/topic/SaleInfo", id="multi_test02", ack='auto',
                       headers={"activemq.subscriptionName": "multi_testC02", "client-id": "multi_test_C02"})
        print("Waiting for messages...")
        while 1:
            time.sleep(3)



class Task_Managerment:
    def __init__(self,message_contrller):
        self._message_contrller = message_contrller
        self._flag = False
        self._cv = threading.Condition()
        self.mode=0

        print ('Thread_Managerment init')


    def tesk4sendMessage(self):
        '''
        new version for messange task
        :return:
        '''
        print("read task running now")

        while True:
            time.sleep(0.5)
            try:
                msg = self._message_contrller.read()
                self.sentMessage(msg)
            except Exception:
                print("Wroing in sendmessange")
                logger.info("Wroing in sendmessange")
                # self._oldday = -1
            # finally:
            #     self._flag = False




    def sentMessage(self, msg):
        try:
            hWndList = []
            win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)
            self.check_and_send(msg, hWndList)
        except Exception:
            logger.info("error in sendmessage")
            print("error in sendmessage")


    def check_and_send(self, msg, hwnds):
            try:
                print(msg[0], msg[1])
                for hwnd in hwnds:
                    # print (type(hwnd))
                    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                        if msg[1] in win32gui.GetWindowText(hwnd):
                            #set the aim window top
                            self.prepare_for_send(hwnd)
                            if msg[2]== "text":
                                temple=msg[0]       #self._datastruction.read()
                                all=temple.split('\\n')

                                for submes in all:
                                    if submes!='':
                                        print(submes)

                                        w.OpenClipboard()
                                        w.EmptyClipboard()
                                        w.SetClipboardData(win32con.CF_UNICODETEXT, submes)
                                        w.CloseClipboard()
                                        self.function_Ctrl_V()
                                        if self.mode==0:
                                            self.new_Line(hwnd)
                                        else:
                                            self.common_for_send(hwnd)
                            if msg[2]== "blod":
                                aString = windll.user32.LoadImageW(0, msg[0], win32con.IMAGE_BITMAP,
                                                                   0, 0, win32con.LR_LOADFROMFILE)
                                if aString != 0:
                                    w.OpenClipboard()
                                    w.EmptyClipboard()
                                    w.SetClipboardData(win32con.CF_BITMAP, aString)
                                    w.CloseClipboard()
                                    print("复制完成")
                                    self.function_Ctrl_V()
                                    if self.mode == 0:
                                        self.new_Line(hwnd)
                                    else:
                                        self.common_for_send(hwnd)
                                    #os.remove(array_all[0])#删除文件
                                else:
                                    print('接受图片有问题')

                            if self.mode == 0:
                                self.common_for_send(hwnd)
                            else:
                                self.new_Line(hwnd)

            except Exception:
                print("xxxxwroing")
                logger.info("send wroing core")
    def new_Line(self,hwnd):
         '''
         :param hwnd:
         :return: new line
         '''
         time.sleep(0.5)
         win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
         win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
         time.sleep(0.5)

    def function_Ctrl_V(self):
        win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
        time.sleep(0.1)
        win32api.keybd_event(86, 0, 0, 0)  # V键位码是86
        win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)


    def common_for_send(self,hwnd):
        '''

        :param hwnd:
        :return:use ctl + enter for send
        '''
        time.sleep(1)
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)

        time.sleep(0.1)
        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        # win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)





    def prepare_for_send(self,hwnd):
        title = win32gui.GetWindowText(hwnd)
        cunrrenthandle = win32gui.GetForegroundWindow()
        curenttitle = win32gui.GetWindowText(cunrrenthandle)
        print(title, curenttitle)
        if title != curenttitle:
            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
            win32gui.SetForegroundWindow(hwnd)

    def check(self,array):
        sum=0
        for i in range(0,len(array)):
            sum+=array[i]
        if sum==len(array):
            print("this complet")
            for i in range(0, len(array)):
                array[i]=0

        # print("array",array)









if __name__ == '__main__':

    LOG_FILE = 'logging.log'

    handler = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=1024 * 1024, backupCount=5)  # 实例化handler
    fmt = '%(asctime)s - %(filename)s:%(lineno)s - %(name)s - %(message)s'

    formatter = logging.Formatter(fmt)  # 实例化formatter
    handler.setFormatter(formatter)  # 为handler添加formatter

    logger = logging.getLogger('tst')  # 获取名为tst的logger
    logger.addHandler(handler)  # 为logger添加handler
    logger.setLevel(logging.DEBUG)


    message_contrller=Message_Contrl()

    receiver=RometeMQReceiver(message_contrller)

    threadmanager = Task_Managerment(message_contrller)
    pool=ThreadPoolExecutor(2)
    # pool.submit(threadmanager.task4read_and_send,)

    pool.submit(threadmanager.tesk4sendMessage,)

    pool.submit(receiver.recevie())

    #pyinstaller -F E:\WX_PUSH_3\WX_Push\WX_Clinet_SalePush.py







