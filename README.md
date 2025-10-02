# ZoomMate

A lightweight AutoIt (AU3) script that automates common Zoom tasks for meetings of Jehovah's Witnesses. Designed to let AV operators focus less on Zoom controls, this script handles routine actions like opening meetings, managing host tools, and streamlining the virtual meeting experience.

## What It Does

ZoomMate automatically:

- Launches Zoom meetings at scheduled times
- Configures meeting security settings before and after meetings
- Applies meeting-specific settings when meetings start
- Manages host controls and participant settings
- Runs in the system tray for easy access and configuration

## Requirements

- Windows OS
- [AutoIt](https://www.autoitscript.com/site/autoit/downloads/) installed for running `.au3` scripts or compiling to `.exe`
- Zoom desktop client installed

## Installation & Setup

1. **Download and extract** the ZoomMate files to your desired location.

2. **Configure your settings:**
   - Run `ZoomMate.au3` (double-click if AutoIt is installed)
   - The configuration GUI will launch automatically
   - Set your Meeting ID, meeting times, and language preferences
   - Configure any other settings as needed

3. **Start using ZoomMate:**
   - The script will run in your system tray
   - Click the tray icon to access settings and configuration
   - ZoomMate will automatically manage your Zoom meetings based on your schedule

## Optional: Compile to Standalone Executable

If you prefer not to install AutoIt or want a standalone executable:

```powershell
Aut2exe.exe /in "ZoomMate.au3" /out "ZoomMate.exe" /icon "zoommate.ico"
```

This creates `ZoomMate.exe` with your custom icon that can be run without AutoIt installed.

## Usage

Once configured, ZoomMate runs automatically in the background:

- It monitors your scheduled meeting times
- Launches Zoom and applies your configured settings
- Manages meeting controls throughout the session
- Access the tray icon anytime to modify settings or view status

The script is designed to be unobtrusive and requires minimal interaction once initially configured.
