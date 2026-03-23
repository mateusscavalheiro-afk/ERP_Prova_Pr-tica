import pyautogui
import time

print("Program will pause for 5 seconds. Move your mouse to the desired position.")
time.sleep(15)  # Pause for 5 seconds
x, y = pyautogui.position() # Get current mouse position
print(f"Mouse position after 15 seconds: X: {x}, Y: {y}")