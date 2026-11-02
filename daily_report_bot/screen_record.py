import cv2
import numpy as np
import pyautogui
import time

OUTPUT = "pyautogui_project_demo.mp4"

FPS = 10

screen_width, screen_height = pyautogui.size()

fourcc = cv2.VideoWriter_fourcc(*"mp4v")

video = cv2.VideoWriter(
    OUTPUT,
    fourcc,
    FPS,
    (screen_width, screen_height)
)

print("Recording started...")
print("Press Ctrl+C to stop.")

try:
    while True:

        screenshot = pyautogui.screenshot()

        frame = np.array(screenshot)

        frame = cv2.cvtColor(
            frame,
            cv2.COLOR_RGB2BGR
        )

        video.write(frame)

        time.sleep(1 / FPS)

except KeyboardInterrupt:

    print("Recording stopped.")

finally:

    video.release()

    cv2.destroyAllWindows()

    print("Video saved:", OUTPUT)