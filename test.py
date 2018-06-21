from pynput import mouse,keyboard
import pyautogui
import time
mouse.Button
keyboard.Key
keyboard.Controller
controller=mouse.Controller()
print(pyautogui.size())

for i in range(100):
    time.sleep(4)
    print(str(i)+"              *****************                 **")
    print(controller.position)
    # pyautogui.press("space")
