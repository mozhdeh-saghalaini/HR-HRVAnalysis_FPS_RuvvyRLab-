"""
@author: Mozhdeh
"""
 
# Importing libraries
# ------------------------------------------------------------
import pyautogui
import time
import subprocess
import os
import pygetwindow as gw   # For window management
import pytesseract
from PIL import Image


# Defining MindWare HRV software path and acquisition file name
# ------------------------------------------------------------
mindware_path = r"C:\Program Files (x86)\MindWare\HRV 3.2.13\MindWare HRV Analysis 3.2.13.exe"
acq_file_name = "ECG IMP EDA ACC Demo.mwi"

# Utility functions
# ------------------------------------------------------------

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def verify_channel_map(expected_texts):
    """
    Captures the Channel Map window and verifies that only ECG is set.
    expected_texts: dict with keys 'ECG', 'Z0', 'dZdt', 'Resp' and expected values.
    """
    screenshot_path = "channel_map_check.png"
    pyautogui.screenshot(screenshot_path, region=(300, 300, 700, 400))  
    extracted_text = pytesseract.image_to_string(Image.open(screenshot_path))

    print("Extracted Channel Map Text:")
    print(extracted_text)

    for label, expected in expected_texts.items():
        if expected not in extracted_text:
            print(f"ERROR: {label} does not match expected value '{expected}'")
            return False
    print("Channel Map verified successfully.")
    return True


def wait_and_click(image_path, confidence=0.9, timeout=15):
    """
    Waits until the given image appears on screen, then clicks it.
    Uses pyautogui.locateOnScreen for robust detection.
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
        if location:
            pyautogui.click(location)
            print(f"Clicked {image_path}")
            return True
        time.sleep(0.5)
    print(f"ERROR: Could not find {image_path} within {timeout} seconds.")
    return False

def type_text(text, delay=0.1):
    """Types text safely with a small delay between keystrokes."""
    pyautogui.write(text, interval=delay)

# ------------------------------------------------------------
# Step 1: Launch MindWare HRV software
# ------------------------------------------------------------
print("Starting MindWare HRV...")
subprocess.Popen(mindware_path)
time.sleep(20)

# Force window to a known position (top-left, fixed size)
try:
    app_window = gw.getWindowsWithTitle("MindWare HRV Analysis")[0]
    app_window.moveTo(0, 0)
    app_window.resizeTo(1280, 800)
    print("Window repositioned for consistent automation.")
except Exception as e:
    print("WARNING: Could not reposition window:", e)

# ------------------------------------------------------------
# Step 2: Navigate through startup dialogs
# ------------------------------------------------------------
wait_and_click("demo_mode.png")      # Screenshot of 'Demo Mode' button
time.sleep(2)

wait_and_click("continue.png")       # Screenshot of 'Continue' button
time.sleep(7)

# ------------------------------------------------------------
# Step 3: Load acquisition file
# ------------------------------------------------------------
wait_and_click("filename_field.png") # Screenshot of filename input field
time.sleep(1)

pyautogui.hotkey('ctrl', 'a')
time.sleep(0.5)
pyautogui.press('delete')
time.sleep(0.5)

type_text(acq_file_name)
time.sleep(1)
pyautogui.press('enter')
print("File opened successfully")
time.sleep(3)

# ------------------------------------------------------------
# Step 4: Confirm ECG channel selection
# ------------------------------------------------------------
# Wait for Channel Map window to appear
time.sleep(3)

# Verify that only ECG is set
expected_channels = {
    "ECG": "ECG",      # Should contain "ECG"
    "Z0": "",          # Should be empty
    "dZdt": "",        # Should be empty
    "Resp": ""         # Should be empty
}

if verify_channel_map(expected_channels):
    wait_and_click("ok_channel_map.png")
    time.sleep(2)
else:
    print("Channel Map verification failed. Aborting.")
    exit()


# ------------------------------------------------------------
# Step 5: Set Segment Time to 60 seconds
# ------------------------------------------------------------
wait_and_click("segment_time_field.png")
time.sleep(1)
pyautogui.doubleClick()
time.sleep(0.5)
type_text("60")
pyautogui.press('enter')
print("Set segment time to 60 seconds")
time.sleep(5)

# ------------------------------------------------------------
# Step 6: HRV Calibration Settings
# ------------------------------------------------------------
wait_and_click("hrv_calibration_tab.png")
time.sleep(1)
wait_and_click("calculation_entire.png")
time.sleep(1)

# LF filter
wait_and_click("lf_field.png")
pyautogui.doubleClick()
type_text("0.15")
pyautogui.press('enter')
print("Set LF filter to 0.15")
time.sleep(5)

# HF filter
wait_and_click("hf_field.png")
pyautogui.doubleClick()
type_text("0.15")
pyautogui.press('enter')
print("Set HF filter to 0.15")
time.sleep(5)

# ------------------------------------------------------------
# Step 7: Additional Settings
# ------------------------------------------------------------
wait_and_click("rpeak_tab.png")
time.sleep(1)

wait_and_click("additional_settings_tab.png")
time.sleep(1)

wait_and_click("use_default_directory.png")
wait_and_click("use_default_directory.png")
time.sleep(1)

wait_and_click("folder_field.png")
type_text(r"C:\Users\user\Downloads\ECG-60-Hz-Noise")
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('enter')

print("Set output folder")
time.sleep(5)

# ------------------------------------------------------------
# Step 8: Run analysis
# ------------------------------------------------------------
wait_and_click("analyze_button.png")
print("Analysis started successfully")
