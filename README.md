
# Motion Detection Application

## Overview

This Python-based motion detection application is designed to detect motion through a connected webcam and save footage whenever motion is detected. The clips are saved to the currently logged-on user's OneDrive `Videos` folder. The application runs a scheduled motion detection by default but can also be manually overridden to run outside of scheduled hours.

The application is built using:
- OpenCV for capturing and processing video streams.
- Tkinter for the graphical user interface (GUI).
- Additional libraries for process and file management.

## Features
1. **Motion Detection**: Detects motion and saves the footage to the user's OneDrive `Videos` folder.
2. **Scheduled Mode**: The application runs between 9 AM and 6 PM on weekdays (Monday to Friday).
3. **Override Mode**: Can manually start motion detection outside of scheduled hours.
4. **Motion Count**: Tracks and displays how many motions were detected each day.
5. **Auto-Delete Old Videos**: Automatically deletes videos older than 3 days.
6. **Disk Space Monitoring**: Stops recording if the disk space is below 1 GB to prevent issues.

## Requirements

### Python Libraries:
Ensure you have the following libraries installed:

```
pip install opencv-python pillow pywin32 psutil
```

### Running as Administrator
The application may need administrative privileges to access the camera and manage system processes.

To ensure the application runs with sufficient permissions:
1. Open **Visual Studio Code**.
2. Right-click the **Visual Studio Code** shortcut and select **Run as administrator**.
3. Open the project folder containing the Python script in Visual Studio Code.

### Running the Application in Visual Studio Code:
1. Open Visual Studio Code.
2. Load the project folder.
3. Ensure all required Python libraries are installed.
4. Open the script file (e.g., `motion_detector_app.py`).
5. Press `F5` or click **Run** > **Run Without Debugging** to execute the program.

## How It Works

1. **Webcam Setup**: 
   - When the application is started, it attempts to connect to the default webcam. If multiple webcams are available, you can select one from the dropdown menu in the GUI.
   - If the webcam is in use by another application, the program detects it and asks if you want to close the conflicting applications.

2. **Motion Detection**:
   - The application uses OpenCV to process frames from the webcam. It detects motion by comparing consecutive frames.
   - If motion is detected, the program begins recording and saves the footage to the user's OneDrive `Videos` folder.

3. **Saving Footage**:
   - All detected motion footage is saved to the current user's OneDrive `Videos` folder, located at:
     ```
     C:\Users\<CurrentUser>\OneDrive\Videos
     ```
   - The footage is saved with filenames in the format `temp_motion_YYYY-MM-DD-HH-MM-SS.avi`.

4. **Scheduling**:
   - By default, the motion detection is active from 9 AM to 6 PM on weekdays (Monday to Friday). You can start or stop this mode by using the GUI buttons.

5. **Override Mode**:
   - You can manually override the scheduled mode to start motion detection at any time by clicking the **Start Override** button.

6. **Motion Count**:
   - The GUI displays a count of the number of motions detected each day. You can reset this counter using the **Reset Counter** button.

7. **Automatic Cleanup**:
   - Videos older than 3 days are automatically deleted from the OneDrive `Videos` folder.

## Controls and GUI

- **Start Schedule**: Start the scheduled motion detection mode (active from 9 AM to 6 PM).
- **Stop Schedule**: Stop the scheduled mode.
- **Start Override**: Start motion detection manually, outside of scheduled hours.
- **Stop Override**: Stop manual override mode.
- **Reset Counter**: Reset the daily motion detection counter.
- **Exit**: Close the application.

## Troubleshooting

1. **Webcam Not Detected**: 
   - Ensure the webcam is connected properly and not being used by another application.
   - If the webcam is in use, close any conflicting applications.

2. **Videos Not Saving**:
   - Ensure that OneDrive is set up and accessible on your system.
   - Verify that the application has write permissions to the OneDrive `Videos` folder.

3. **Administrator Access**:
   - If the application fails to access the webcam or manage processes, try running it as an administrator by launching Visual Studio Code with admin privileges.

## Notes

- Ensure that OneDrive is syncing properly if you want to back up the motion-detected videos to the cloud.
- If the application is run on a system without OneDrive, the footage will not be saved to the cloud.

## License

This project is open-source and can be modified or distributed as per the requirements of the user.
