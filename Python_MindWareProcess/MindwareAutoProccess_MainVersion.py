"""
Last Update: 11/28/2025

@author: Mozhdeh Saghalaini - email:m.saghalaini@gmail.com
This code is for Automating MindWare HRV (Main Version) workflow
"""

# %% ---------------------------------------------------------
# Importing libraries
# ------------------------------------------------------------

import pyautogui
import pygetwindow as gw   
import time
import subprocess
import os
import pytesseract
from PIL import Image
import winsound

# %% ---------------------------------------------------------
# Asking user for settings via console
# ------------------------------------------------------------

print("Please provide the following paths and settings:")

## Option1: Paths are predefined
# MindWare HRV executable 
mindware_path = r"C:\Program Files (x86)\MindWare\HRV 3.2.13\MindWare HRV Analysis 3.2.13.exe"

# Folder containing data files
acq_folder = r"C:\Users\user\OneDrive\Documents\MindWare\HRV Examples"

# Output folder 
output_folder = r"C:\Users\user\Downloads\ECG-60-Hz-Noise"

## Option2: Paths are provided by user
# mindware_path = input(
#     "Enter full path to MindWare HRV executable (.exe)\n"
#     "Example: C:\Program Files (x86)\MindWare\HRV 3.2.13\MindWare HRV Analysis 3.2.13.exe\n> "
# ).strip()

# acq_folder = input(
#     "Enter folder path containing data files\n"
#     "Example: C:\Users\user\OneDrive\Documents\MindWare\HRV Examples\n> "
# ).strip()

# output_folder = input(
#     "Enter output folder path for the processed data\n"
#     "Example: C:\Users\user\Downloads\ECG-60-Hz-Noise\n> "
# ).strip()

segment_time = int(input("Enter segment time in seconds (default 60): ") or 60)
lf_filter = float(input("Enter LF filter value (default 0.15 Hz): ") or 0.15)
hf_filter = float(input("Enter HF filter value (default 0.15 Hz): ") or 0.15)
max_number_of_seg = int(input("Enter The Maximum Number of Segments in Each File (default 20): ") or 20)

print("\n Settings chosen:")
print(f"MindWare path: {mindware_path}")
print(f"Acquisition folder: {acq_folder}")
print(f"Output folder: {output_folder}")
print(f"Segment time: {segment_time}")
print(f"LF filter: {lf_filter}")
print(f"HF filter: {hf_filter}")
print(f"Maximum Number of Segments: {max_number_of_seg}")

# %% ---------------------------------------------------------
# Utility functions 
# ------------------------------------------------------------
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # This needs to be set based on the system as well

# Checking whether the Channel Map window is set up correctly
def verify_channel_map(expected_texts):
    screenshot_path = "channel_map_check.png"
    pyautogui.screenshot(screenshot_path, region=(806, 496, 1114, 673))  
    extracted_text = pytesseract.image_to_string(Image.open(screenshot_path))
    print("Extracted Channel Map Text:")
    print(extracted_text)
    for label, expected in expected_texts.items():
        if expected not in extracted_text:
            print(f"ERROR: {label} does not match expected value '{expected}'")
            return False
    print("Channel Map verified successfully.")
    return True

# ------------------------------------------------------------
# Utility helper
def wait_and_click(image_path, confidence=0.9, timeout=15):
    start_time = time.time()
    while time.time() - start_time < timeout:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
        if location:
            pyautogui.click(location)
            print(f"Clicked {image_path}")
            return True
        time.sleep(0.5)
    # If not found then beep + wait for user
    winsound.MessageBeep()
    print(f"Could not find {image_path}. Please fix manually, then press ENTER to continue...")
    input()
    return False


# ------------------------------------------------------------
# Typing text
def type_text(text, delay=0.1):
    pyautogui.write(text, interval=delay)

# ------------------------------------------------------------
# Visual checker for detecting problematic R peaks
def segment_has_yellow_peaks(region=(403, 274, 700, 128), threshold=0.0001):
    time.sleep(1)
    screenshot_path = "ecg_segment.png"
    pyautogui.screenshot(screenshot_path, region=region)
    img = Image.open(screenshot_path).convert("RGB")
    yellow_count = 0
    total_pixels = img.width * img.height
    for r, g, b in img.getdata():
        if r > 180 and g > 180 and b < 100:
            yellow_count += 1
    yellow_ratio = yellow_count / total_pixels
    print(f"Yellow pixel ratio: {yellow_ratio:.6f}")
    return yellow_ratio > threshold

# ------------------------------------------------------------
# Utility helper
def check_all_segments(max_segments=max_number_of_seg):
    for i in range(max_segments):
        print(f"\nChecking segment {i+1}...")
        if segment_has_yellow_peaks():
            print("Yellow peaks detected. Clicking 'Edit R’s'...")
            winsound.MessageBeep()
            wait_and_click("edit_rs_button.png")
            time.sleep(2)
            input(">>> Fix R-peaks manually, close the Edit window, then press Enter in the console two times to continue...")
        else:
            print("Segment is clean. No action needed.")
        pyautogui.click(x=458, y=206)  
        time.sleep(3)

# ------------------------------------------------------------
# Safe Action (Error handling)
def safe_action(action_fn, *args, **kwargs):
    """
    Run an action safely: if it fails, beep, ask user to fix manually,
    and wait for Enter before continuing.
    """
    try:
        return action_fn(*args, **kwargs)
    except Exception as e:
        winsound.MessageBeep()
        print(f"ERROR: {e}")
        print("Please fix the issue manually, then press ENTER in the console to continue...")
        input()
        return None

# %% ---------------------------------------------------------
# Launching MindWare HRV software
# ------------------------------------------------------------
print("Starting MindWare HRV...")
subprocess.Popen(mindware_path)
time.sleep(25)

# Force window to a known position
try:
    app_window = gw.getWindowsWithTitle("MindWare HRV Analysis")[0]
    app_window.moveTo(0, 0)
    app_window.resizeTo(1280, 800)
    print("Window repositioned for consistent automation.")
except Exception as e:
    print("WARNING: Could not reposition window:", e)

# %% ---------------------------------------------------------
# # Navigating through startup dialogs (specifically for the Demo Version)
# # ------------------------------------------------------------
# wait_and_click("demo_mode.png")      # 'Demo Mode' button
# time.sleep(2)

# wait_and_click("continue.png")       # 'Continue' button
# time.sleep(7)

# %% ---------------------------------------------------------
# Detecting .acq files
# ------------------------------------------------------------

files = [f for f in os.listdir(acq_folder) if f.lower().endswith(".acq")]
if not files:
    print(" No .acq files found in folder!")
    # Alt+F4+Fn
    pyautogui.hotkey('alt', 'fn', 'f4')
    time.sleep(2)
    
# Print how many files were detected
print(f" Found {len(files)} .acq files to process:")

# %% ---------------------------------------------------------
# Looping through each files and processing them 
# ------------------------------------------------------------
 
for acq_file_name in files:
    full_path = os.path.join(acq_folder, acq_file_name)
    print(f"\n Starting analysis for file: {acq_file_name}")
    
    # Clicking folder path field and typing the foldername
    safe_action(wait_and_click, "folder_path_field.png")   
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(0.5)
    type_text(acq_folder)  
    time.sleep(1)
    pyautogui.press('enter')
    
    # Clicking filename field and typing the file name
    safe_action(wait_and_click, "filename_field.png")     
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


    # Confirming ECG channel selection
    time.sleep(3)
    expected_channels = {"ECG": "ECG", "Z0": "", "dZdt": "", "Resp": ""}
    if verify_channel_map(expected_channels):
        safe_action(wait_and_click, "ok_channel_map.png")
        time.sleep(5)
    else:
        winsound.MessageBeep()
        print(" Channel Map verification failed.")
        print("Please refine the Channel Map manually, then press ENTER in the console to continue...")
        input()
    
        safe_action(wait_and_click, "ok_channel_map.png", confidence=0.7)
        time.sleep(5)
        continue

    # Adding Digital Event Channel
    print("Adding Digital Event Channel...")
    
    safe_action(wait_and_click, "add_button.png") # add_button for digital event
    time.sleep(5)
    
    screenshot_path = "digital_event_check.png"
    pyautogui.screenshot(screenshot_path, region=(791, 479, 1130, 642))
    extracted_text = pytesseract.image_to_string(Image.open(screenshot_path))
    print(extracted_text)
    
    if "Event Channel" in extracted_text:
        print("Digital Event Channel already set.")

        safe_action(wait_and_click, "event_ok.png")
        time.sleep(2)
        safe_action(wait_and_click, "event_ok.png")
        time.sleep(5)
    else:

        winsound.MessageBeep()
        print("Digital Event Channel not set.")
        print("Please refine the Channel Map manually, then press ENTER in the console to continue...")
        input()
    
        safe_action(wait_and_click, "event_ok.png")
        time.sleep(2)
        safe_action(wait_and_click, "event_ok.png")
        time.sleep(5)
    
    # Handle possible pop-up (Continue button)
    if safe_action(wait_and_click, "continue_button.png", timeout=5, confidence=0.6):
        print("Pop-up detected: pressed Continue.")
    else:
        print("No pop-up detected, continuing workflow.")

    # Setting Segment Time
    safe_action(wait_and_click, "segment_time_field.png")
    time.sleep(1)
    pyautogui.doubleClick()
    time.sleep(0.5)
    type_text(str(segment_time))
    pyautogui.press('enter')
    print(f"Set segment time to {segment_time} seconds")
    time.sleep(5)
        
    # HRV Calibration Settings
    safe_action(wait_and_click, "hrv_calibration_tab.png")
    time.sleep(1)
    safe_action(wait_and_click, "calculation_entire.png")
    time.sleep(1)

    safe_action(wait_and_click, "lf_field.png")
    pyautogui.doubleClick()
    type_text(str(lf_filter))
    pyautogui.press('enter')
    print(f"Set LF Band filter to {lf_filter} Hz")
    time.sleep(5)

    safe_action(wait_and_click, "hf_field.png")
    pyautogui.doubleClick()
    type_text(str(hf_filter))
    pyautogui.press('enter')
    print(f"Set HF/RSA Band filter to {hf_filter} Hz")
    time.sleep(5)

    # R peak and additional setting tabs 
    safe_action(wait_and_click, "rpeak_tab.png")
    time.sleep(1)
    safe_action(wait_and_click, "additional_settings_tab.png")
    time.sleep(1)
    safe_action(wait_and_click, "use_default_directory.png")
    safe_action(wait_and_click, "use_default_directory.png")
    time.sleep(1)
    safe_action(wait_and_click, "folder_field.png")
    type_text(output_folder)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    print("Set output folder")
    time.sleep(5)


    # Running analysis  
    safe_action(wait_and_click, "analyze_button.png")
    print("Analysis started successfully")

    # segment checks
    time.sleep(10)
    check_all_segments()

    # Exporting results
    print("\nAll segments checked. Exporting results...")
    pyautogui.hotkey('ctrl', 'shift', 'w')
    time.sleep(6)
    
    # Step A: Click into the folder path field
    safe_action(wait_and_click, "output_folder_field.png")  
    time.sleep(1)
    
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(0.5)
    
    type_text(output_folder)  
    time.sleep(1)

    pyautogui.press('enter')  
    time.sleep(1)
    pyautogui.press('enter')  
    time.sleep(10)
    
    pyautogui.press('enter')

    print(f" Export complete for {acq_file_name}")

    # Exiting the Analyze window
    # Alt+F4+Fn
    pyautogui.hotkey('alt', 'fn', 'f4')
    time.sleep(2)

    pyautogui.hotkey('ctrl', 'o')
    time.sleep(0.5)
    
    
print("\n All files processed. Workflow finished.")
