import pyautogui
import time

print("Move your mouse to the point you want the bot to click.")
print("The coordinates will appear below.")
print("Press Ctrl+C when you have the correct position.")

while True:
    x, y = pyautogui.position()
    print(f"X = {x}, Y = {y}", end="\r")
    time.sleep(0.2)
