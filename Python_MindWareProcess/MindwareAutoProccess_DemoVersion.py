"""
Last Update: 11/27/2025

@author: Mozhdeh Saghalaini - email:m.saghalaini@gmail.com
This code is for Automating MindWare HRV (Demo Version) workflow
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

print("\n Settings chosen:")
print(f"MindWare path: {mindware_path}")
print(f"Acquisition folder: {acq_folder}")
print(f"Output folder: {output_folder}")
print(f"Segment time: {segment_time}")
print(f"LF filter: {lf_filter}")
print(f"HF filter: {hf_filter}")

# %% ---------------------------------------------------------
# Utility functions 
# ------------------------------------------------------------
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

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
    print(f"ERROR: Could not find {image_path} within {timeout} seconds.")
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
def check_all_segments(max_segments=8):
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
# Navigating through startup dialogs (specifically for the Demo Version)
# ------------------------------------------------------------
wait_and_click("demo_mode.png")      # 'Demo Mode' button
time.sleep(2)

wait_and_click("continue.png")       # 'Continue' button
time.sleep(7)

# %% ---------------------------------------------------------
# Detecting .mwi files
# ------------------------------------------------------------

files = [f for f in os.listdir(acq_folder) if f.lower().endswith(".mwi")]
if not files:
    print("❌ No .mwi files found in folder. Exiting.")
    exit()

# %% ---------------------------------------------------------
# Looping through each files and processing them 
# ------------------------------------------------------------
 
for acq_file_name in files:
    full_path = os.path.join(acq_folder, acq_file_name)
    print(f"\n Starting analysis for file: {acq_file_name}")
    
    # Clicking folder path field and typing the foldername
    wait_and_click("folder_path_field.png")   
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(0.5)
    type_text(acq_folder)  
    time.sleep(1)
    pyautogui.press('enter')
    
    # Clicking filename field and typing the file name
    wait_and_click("filename_field.png")     
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
        wait_and_click("ok_channel_map.png")
        time.sleep(5)
    else:
        winsound.MessageBeep()
        print(" Channel Map verification failed.")
        print("Please refine the Channel Map manually, then press ENTER in the console to continue...")
        input()
    
        wait_and_click("ok_channel_map.png", confidence=0.7)
        time.sleep(5)
        continue

    # Adding Digital Event Channel
    print("➕ Adding Digital Event Channel...")
    
    wait_and_click("add_button.png") # add_button for digital event
    time.sleep(5)
    
    screenshot_path = "digital_event_check.png"
    pyautogui.screenshot(screenshot_path, region=(791, 479, 1130, 642))
    extracted_text = pytesseract.image_to_string(Image.open(screenshot_path))
    print(extracted_text)
    
    if "Event Channel" in extracted_text:
        print("Digital Event Channel already set.")

        wait_and_click("event_ok.png")
        time.sleep(2)
        wait_and_click("event2_ok.png")
        time.sleep(5)
    else:

        winsound.MessageBeep()
        print("Digital Event Channel not set.")
        print("Please refine the Channel Map manually, then press ENTER in the console to continue...")
        input()
    
        wait_and_click("event_ok.png")
        time.sleep(2)
        wait_and_click("event2_ok.png")
        time.sleep(5)
    
    # Handle possible pop-up (Continue button)
    if wait_and_click("continue_button.png", timeout=5, confidence=0.6):
        print("Pop-up detected → pressed Continue.")
    else:
        print("No pop-up detected, continuing workflow.")

    # Setting Segment Time
    wait_and_click("segment_time_field.png")
    time.sleep(1)
    pyautogui.doubleClick()
    time.sleep(0.5)
    type_text(str(segment_time))
    pyautogui.press('enter')
    print(f"Set segment time to {segment_time} seconds")
    time.sleep(5)
        
    # HRV Calibration Settings
    wait_and_click("hrv_calibration_tab.png")
    time.sleep(1)
    wait_and_click("calculation_entire.png")
    time.sleep(1)

    wait_and_click("lf_field.png")
    pyautogui.doubleClick()
    type_text(str(lf_filter))
    pyautogui.press('enter')
    print(f"Set LF Band filter to {lf_filter} Hz")
    time.sleep(5)

    wait_and_click("hf_field.png")
    pyautogui.doubleClick()
    type_text(str(hf_filter))
    pyautogui.press('enter')
    print(f"Set HF/RSA Band filter to {hf_filter} Hz")
    time.sleep(5)

    # R peak and additional setting tabs 
    wait_and_click("rpeak_tab.png")
    time.sleep(1)
    wait_and_click("additional_settings_tab.png")
    time.sleep(1)
    wait_and_click("use_default_directory.png")
    wait_and_click("use_default_directory.png")
    time.sleep(1)
    wait_and_click("folder_field.png")
    type_text(output_folder)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    print("Set output folder")
    time.sleep(5)

    # Runing analysis  
    wait_and_click("analyze_button.png")
    print("Analysis started successfully")

    # segment checks
    time.sleep(10)
    check_all_segments()

    # Exporting results
    print("\nAll segments checked. Exporting results...")
    pyautogui.hotkey('ctrl', 'shift', 'w')
    time.sleep(6)
    
    # Step A: Click into the folder path field
    wait_and_click("output_folder_field.png")  
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




# # Importing libraries
# # ------------------------------------------------------------
# import pyautogui
# import time
# import subprocess
# import os
# import pygetwindow as gw   # For window management
# import pytesseract
# from PIL import Image


# # Defining MindWare HRV software path and acquisition file name
# # ------------------------------------------------------------
# mindware_path = r"C:\Program Files (x86)\MindWare\HRV 3.2.13\MindWare HRV Analysis 3.2.13.exe"
# acq_file_name = "ECG IMP EDA ACC Demo.mwi"

# # Utility functions
# # ------------------------------------------------------------

# # Set Tesseract path
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# def verify_channel_map(expected_texts):
#     """
#     Captures the Channel Map window and verifies that only ECG is set.
#     expected_texts: dict with keys 'ECG', 'Z0', 'dZdt', 'Resp' and expected values.
#     """
#     screenshot_path = "channel_map_check.png"
#     pyautogui.screenshot(screenshot_path, region=(300, 300, 700, 400))  
#     extracted_text = pytesseract.image_to_string(Image.open(screenshot_path))

#     print("Extracted Channel Map Text:")
#     print(extracted_text)

#     for label, expected in expected_texts.items():
#         if expected not in extracted_text:
#             print(f"ERROR: {label} does not match expected value '{expected}'")
#             return False
#     print("Channel Map verified successfully.")
#     return True


# def wait_and_click(image_path, confidence=0.9, timeout=15):
#     """
#     Waits until the given image appears on screen, then clicks it.
#     Uses pyautogui.locateOnScreen for robust detection.
#     """
#     start_time = time.time()
#     while time.time() - start_time < timeout:
#         location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
#         if location:
#             pyautogui.click(location)
#             print(f"Clicked {image_path}")
#             return True
#         time.sleep(0.5)
#     print(f"ERROR: Could not find {image_path} within {timeout} seconds.")
#     return False

# def type_text(text, delay=0.1):
#     """Types text safely with a small delay between keystrokes."""
#     pyautogui.write(text, interval=delay)


# def segment_has_yellow_peaks(region=(403, 274, 700, 128), threshold=0.0001):
#     """
#     Checks if the ECG graph contains yellow R-peaks.
#     Returns True if yellow pixels are detected above threshold.
#     """
#     time.sleep(1)  # allow UI to settle
#     screenshot_path = "ecg_segment.png"
#     pyautogui.screenshot(screenshot_path, region=region)
#     img = Image.open(screenshot_path).convert("RGB")

#     yellow_count = 0
#     total_pixels = img.width * img.height

#     for r, g, b in img.getdata():
#         # broader definition of yellow
#         if r > 180 and g > 180 and b < 100:  # stricter yellow
#             yellow_count += 1


#     yellow_ratio = yellow_count / total_pixels
#     print(f"Yellow pixel ratio: {yellow_ratio:.6f}")
#     return yellow_ratio > threshold


# def check_all_segments(max_segments=8):
#     """
#     Loops through each segment and checks for yellow peaks.
#     If yellow peaks are detected, clicks 'Edit R’s' and waits
#     for user to press Enter before continuing.
#     """
#     for i in range(max_segments):
#         print(f"\nChecking segment {i+1}...")

#         if segment_has_yellow_peaks():
#             print("Yellow peaks detected. Clicking 'Edit R’s'...")
#             wait_and_click("edit_rs_button.png")
#             time.sleep(2)

#             # Pause here until you confirm you're done
#             input(">>> Fix R-peaks manually, close the Edit window, then press Enter to continue...")

#         else:
#             print("Segment is clean. No action needed.")

#         # Move to next segment
#         pyautogui.click(x=458, y=206)  
#         time.sleep(3)



# # ------------------------------------------------------------
# # Step 1: Launch MindWare HRV software
# # ------------------------------------------------------------
# print("Starting MindWare HRV...")
# subprocess.Popen(mindware_path)
# time.sleep(20)

# # Force window to a known position (top-left, fixed size)
# try:
#     app_window = gw.getWindowsWithTitle("MindWare HRV Analysis")[0]
#     app_window.moveTo(0, 0)
#     app_window.resizeTo(1280, 800)
#     print("Window repositioned for consistent automation.")
# except Exception as e:
#     print("WARNING: Could not reposition window:", e)

# # ------------------------------------------------------------
# # Step 2: Navigate through startup dialogs
# # ------------------------------------------------------------
# wait_and_click("demo_mode.png")      # Screenshot of 'Demo Mode' button
# time.sleep(2)

# wait_and_click("continue.png")       # Screenshot of 'Continue' button
# time.sleep(7)

# # ------------------------------------------------------------
# # Step 3: Load acquisition file
# # ------------------------------------------------------------
# wait_and_click("filename_field.png") # Screenshot of filename input field
# time.sleep(1)

# pyautogui.hotkey('ctrl', 'a')
# time.sleep(0.5)
# pyautogui.press('delete')
# time.sleep(0.5)

# type_text(acq_file_name)
# time.sleep(1)
# pyautogui.press('enter')
# print("File opened successfully")
# time.sleep(3)

# # ------------------------------------------------------------
# # Step 4: Confirm ECG channel selection
# # ------------------------------------------------------------
# # Wait for Channel Map window to appear
# time.sleep(3)

# # Verify that only ECG is set
# expected_channels = {
#     "ECG": "ECG",      # Should contain "ECG"
#     "Z0": "",          # Should be empty
#     "dZdt": "",        # Should be empty
#     "Resp": ""         # Should be empty
# }

# if verify_channel_map(expected_channels):
#     wait_and_click("ok_channel_map.png")
#     time.sleep(2)
# else:
#     print("Channel Map verification failed. Aborting.")
#     exit()


# # ------------------------------------------------------------
# # Step 5: Set Segment Time to 60 seconds
# # ------------------------------------------------------------
# wait_and_click("segment_time_field.png")
# time.sleep(1)
# pyautogui.doubleClick()
# time.sleep(0.5)
# type_text("60")
# pyautogui.press('enter')
# print("Set segment time to 60 seconds")
# time.sleep(5)

# # ------------------------------------------------------------
# # Step 6: HRV Calibration Settings
# # ------------------------------------------------------------
# wait_and_click("hrv_calibration_tab.png")
# time.sleep(1)
# wait_and_click("calculation_entire.png")
# time.sleep(1)

# # LF filter
# wait_and_click("lf_field.png")
# pyautogui.doubleClick()
# type_text("0.15")
# pyautogui.press('enter')
# print("Set LF filter to 0.15")
# time.sleep(5)

# # HF filter
# wait_and_click("hf_field.png")
# pyautogui.doubleClick()
# type_text("0.15")
# pyautogui.press('enter')
# print("Set HF filter to 0.15")
# time.sleep(5)

# # ------------------------------------------------------------
# # Step 7: Additional Settings
# # ------------------------------------------------------------
# wait_and_click("rpeak_tab.png")
# time.sleep(1)

# wait_and_click("additional_settings_tab.png")
# time.sleep(1)

# wait_and_click("use_default_directory.png")
# wait_and_click("use_default_directory.png")
# time.sleep(1)

# wait_and_click("folder_field.png")
# type_text(r"C:\Users\user\Downloads\ECG-60-Hz-Noise")
# pyautogui.press('enter')
# time.sleep(1)
# pyautogui.press('enter')

# print("Set output folder")
# time.sleep(5)

# # ------------------------------------------------------------
# # Step 8: Run analysis
# # ------------------------------------------------------------
# wait_and_click("analyze_button.png")
# print("Analysis started successfully")

# # ------------------------------------------------------------
# # Step 9: Post-analysis segment check
# # ------------------------------------------------------------
# time.sleep(10)  # give the first segment time to fully render
# check_all_segments()


# # ------------------------------------------------------------
# # Step 10: Export results
# # ------------------------------------------------------------
# print("\nAll segments checked. Exporting results...")

# # Trigger Write All Segments via keyboard shortcut
# pyautogui.hotkey('ctrl', 'shift', 'w')
# time.sleep(6)

# # # In the pop-up, select folder
# # wait_and_click("export_folder_field.png")
# # type_text(r"C:\Users\user\Downloads\ECG-60-Hz-Noise")
# pyautogui.press('enter')
# time.sleep(1)


# # Press OK
# pyautogui.press('enter')
# time.sleep(2)

# print("Export complete. Workflow finished.")


