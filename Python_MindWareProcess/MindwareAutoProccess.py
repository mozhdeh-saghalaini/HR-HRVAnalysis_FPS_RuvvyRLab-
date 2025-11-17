"""
Created on Wed Oct 22 14:08:55 2025

@author: Mozhdeh 
"""

# Importing libraries
import pyautogui        
import time            
import subprocess       
import os              
from PIL import Image   # Image processing 


# Defining MindWare HRV software path and acquisition file name
# ------------------------------------------------------------
mindware_path = r"C:\Program Files (x86)\MindWare\HRV 3.2.13\MindWare HRV Analysis 3.2.13.exe"
acq_file_name = "ECG IMP EDA ACC Demo.mwi"

# Launching MindWare HRV software
# ------------------------------------------------------------
print("Starting MindWare HRV...")
subprocess.Popen(mindware_path)   
time.sleep(20)                   

# Clicks 'Demo Mode' button
# ------------------------------------------------------------
screen_width, screen_height = pyautogui.size()
demo_x = screen_width // 2 + 150
demo_y = screen_height // 2 + 150
pyautogui.click(demo_x, demo_y)
print("Clicked Demo Mode")
time.sleep(2)

# Clicks 'Continue' button
# ------------------------------------------------------------
continue_x = 966
continue_y = 636
pyautogui.click(continue_x, continue_y)
print("Clicked Continue")
time.sleep(7)

# Loading acquisition file
# ------------------------------------------------------------
filename_field_x = screen_width // 2
filename_field_y = screen_height // 2 + 100
pyautogui.click(filename_field_x, filename_field_y)
time.sleep(1)

# Clear existing text
pyautogui.hotkey('ctrl', 'a')
time.sleep(0.5)
pyautogui.press('delete')
time.sleep(0.5)

# Type filename and press Enter
pyautogui.write(acq_file_name)
time.sleep(1)
pyautogui.press('enter')
time.sleep(2)
print("File opened successfully")
time.sleep(3)

# Confirming ECG channel selection (Channel Map window)
# ------------------------------------------------------------
ok_x = 873
ok_y = 699
pyautogui.click(ok_x, ok_y)
print("Clicked OK on Channel Map window")
time.sleep(2)

# Setting Segment Time to 60 seconds
# ------------------------------------------------------------
pyautogui.click(929, 533)         
time.sleep(1)
pyautogui.doubleClick(929, 533)   
time.sleep(0.5)
pyautogui.write('60')            
time.sleep(1)
pyautogui.press('enter')
print("Set segment time to 60 seconds")
time.sleep(5)

# Opening HRV Calibration Settings
# ------------------------------------------------------------
pyautogui.click(817, 227)         # Clicks HRV Calibration tab
time.sleep(1)
pyautogui.click(750, 445)         # Clicks Calculation Method (Entire)
time.sleep(1)
pyautogui.click(750, 445)         # Clicks Calculation Method (Entire)
time.sleep(1)

# Setting Frequency Filter values (LF & HF)
# ------------------------------------------------------------
# LF filter
pyautogui.click(999, 623)
time.sleep(1)
pyautogui.doubleClick(999, 623)
time.sleep(0.5)
pyautogui.write('0.15')
time.sleep(1)
pyautogui.press('enter')
print("Set LF filter to 0.15")
time.sleep(5)

# HF filter
pyautogui.click(1084, 623)
time.sleep(1)
pyautogui.doubleClick(1084, 623)
time.sleep(0.5)
pyautogui.write('0.15')
time.sleep(1)
pyautogui.press('enter')
print("Set HF filter to 0.15")
time.sleep(5)

# Navigating to R Peak & Artifact Settings tab
# ------------------------------------------------------------
pyautogui.click(1200, 227)
time.sleep(1)

# Navigating to Additional Settings tab
# ------------------------------------------------------------
pyautogui.click(1400, 227)
time.sleep(1)

# Toggles 'Use Default Directory' option
pyautogui.click(580, 760)
time.sleep(1)
pyautogui.click(580, 760)
time.sleep(1)

# Setting output folder directory
# ------------------------------------------------------------
pyautogui.click(1038, 300)   # Click folder input field
time.sleep(1)
pyautogui.write(r'C:\Users\user\Downloads\ECG-60-Hz-Noise')
time.sleep(2)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('enter')
pyautogui.click(1186, 807)
print("Set output folder")
time.sleep(5)

# Run analysis
# ------------------------------------------------------------
pyautogui.click(781, 957)    
time.sleep(1)
print("Analysis started successfully")













# # Mouse Location Code

# import pyautogui
# import time

# print("Move your mouse to the desired position and press Ctrl+C to stop.")
# try:
#     while True:
#         x, y = pyautogui.position()
#         print(f'X: {x}, Y: {y}', end='\r')
#         time.sleep(0.1)
# except KeyboardInterrupt:
#     print(f'\nFinal coordinates: X: {x}, Y: {y}')