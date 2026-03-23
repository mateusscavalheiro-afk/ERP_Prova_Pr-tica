import pyautogui
import time

print("Program will pause for 15 seconds. Move your mouse to the desired position.")
time.sleep(5)  # Pause for 5 seconds
x, y = pyautogui.position() # Get current mouse position
print(f"Mouse position after 15 seconds: X: {x}, Y: {y}")