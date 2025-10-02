# ZoomMate

A lightweight AutoIt (AU3) script that automates common Zoom tasks for meetings of Jehovah’s Witnesses. Designed to let AV operators focus less on Zoom controls, this script handles routine actions like opening meetings, managing host tools, and streamlining the virtual meeting experience.

## How To Use

1. **Configure your meeting settings:**
   - Run the script or compiled app to use launch the configuration GUI.
   - Set your Meeting ID, meeting times, and language.

2. **Run the script:**
   - Double-click `ZoomMate.au3` if you have AutoIt installed.
   - Or compile it to an executable as described below.

3. **Compile to EXE (optional):**
   - Use the AutoIt compiler `Aut2exe.exe` to create a standalone executable.
   - Example command line:

    ```powershell
     Aut2exe.exe /in "ZoomMate.au3" /out "ZoomMate.exe" /icon "zoommate.ico"
     ```

   - This will generate `ZoomMate.exe` with your custom icon.

4. **Start ZoomMate:**
   - Run the compiled EXE or the AU3 script.
   - The tray icon will appear; click it to access configuration.

5. **Automation:**
   - ZoomMate will automatically launch Zoom and manage meeting controls based on your schedule.

## Requirements

- Windows OS
- [AutoIt](https://www.autoitscript.com/site/autoit/downloads/) installed for running `.au3` scripts or compiling to `.exe`
- Zoom desktop client installed
