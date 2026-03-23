import pyautogui
import sys
import time

# Enable failsafe (it's True by default, but good practice to be explicit)
pyautogui.FAILSAFE = True 

# You can adjust the pause duration to give more or less time to trigger the failsafe
pyautogui.PAUSE = 0.1 

print("Program started. Move mouse to the top-left corner to stop.")

try:
    while True:
        # Example automation code - replace with your actual tasks
        # These calls have a built-in check for the top-left corner
        pyautogui.moveTo(100, 100, duration=0.25)
        pyautogui.moveTo(200, 200, duration=0.25)
        # Add more pyautogui functions here...

        # You can also manually check the position at any time
        x, y = pyautogui.position()
        if x == 0 and y == 0:
            print("Fail-safe triggered manually by position check.")
            sys.exit()

except pyautogui.FailSafeException:
    print("Fail-safe exception raised. Program aborted.")
    sys.exit()
except KeyboardInterrupt:
    print("Program stopped by user (Ctrl+C).")
    sys.exit()
