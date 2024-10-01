import tkinter as tk
from tkinter import messagebox
import cv2
import time
import os
from collections import deque
from datetime import datetime, timedelta
from PIL import Image, ImageTk
import shutil
import win32com.client
import psutil  # For process management

class MotionDetectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Motion Detector")

        # Variables
        self.scheduled_mode = False
        self.override_mode = False
        self.motion_counter = 0
        self.last_reset_date = datetime.now().date()
        self.selected_camera_index = 0
        self.camera_info = []

        # Create UI elements
        self.create_widgets()

        # Initialize the camera
        self.cap = None
        self.initialize_camera()

        # Initialize motion detection variables
        self.first_frame = None
        self.motion_detected = False
        self.video_writer = None
        self.buffer_size = 60
        self.frame_buffer = deque(maxlen=self.buffer_size)
        self.last_flush_time = time.time()
        self.flush_interval = 10
        self.temp_filename = None

        # Set the OneDrive Videos folder path using the current logged-on user
        self.onedrive_videos_folder = os.path.join(os.environ['OneDrive'], "Videos")

        # Ensure the directory exists
        if not os.path.exists(self.onedrive_videos_folder):
            os.makedirs(self.onedrive_videos_folder)

        self.min_free_space_gb = 1

        # Schedule the first frame processing
        self.process_frame()

        # Schedule periodic tasks
        self.schedule_delete_old_videos()
        self.reset_first_frame()

    def create_widgets(self):
        # Create buttons and counter label
        self.start_schedule_button = tk.Button(self.root, text="Start Schedule", command=self.start_schedule)
        self.start_schedule_button.grid(row=0, column=0, padx=5, pady=5)

        self.stop_schedule_button = tk.Button(self.root, text="Stop Schedule", command=self.stop_schedule)
        self.stop_schedule_button.grid(row=0, column=1, padx=5, pady=5)

        self.start_override_button = tk.Button(self.root, text="Start Override", command=self.start_override)
        self.start_override_button.grid(row=1, column=0, padx=5, pady=5)

        self.stop_override_button = tk.Button(self.root, text="Stop Override", command=self.stop_override)
        self.stop_override_button.grid(row=1, column=1, padx=5, pady=5)

        self.exit_button = tk.Button(self.root, text="Exit", command=self.exit_application)
        self.exit_button.grid(row=2, column=0, padx=5, pady=5)

        self.reset_counter_button = tk.Button(self.root, text="Reset Counter", command=self.reset_counter)
        self.reset_counter_button.grid(row=2, column=1, padx=5, pady=5)

        self.counter_label = tk.Label(self.root, text=f"Motion Detections Today: {self.motion_counter}")
        self.counter_label.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

        # Webcam selection
        self.webcam_label = tk.Label(self.root, text="Available Webcams:")
        self.webcam_label.grid(row=4, column=0, padx=5, pady=5)

        self.camera_var = tk.StringVar()
        self.camera_dropdown = tk.OptionMenu(self.root, self.camera_var, [])
        self.camera_dropdown.grid(row=4, column=1, padx=5, pady=5)

        # Video display label
        self.video_label = tk.Label(self.root)
        self.video_label.grid(row=5, column=0, columnspan=2)

    def initialize_camera(self):
        self.detect_cameras()
        if not self.camera_info:
            messagebox.showerror("Error", "No webcams detected. Please connect a webcam and restart the application.")
            # Do not exit the application, allow user to connect a webcam
            return

        # Update the dropdown menu with available webcams
        camera_names = [info['name'] for info in self.camera_info]
        self.camera_var.set(camera_names[0])  # Set default selection
        menu = self.camera_dropdown["menu"]
        menu.delete(0, "end")
        for name in camera_names:
            menu.add_command(label=name, command=lambda value=name: self.camera_var.set(value))

        # Bind the selection event
        self.camera_var.trace('w', self.on_camera_selection)

        # Open the default camera
        self.selected_camera_index = self.camera_info[0]['index']
        self.open_camera(self.selected_camera_index)

    def detect_cameras(self):
        self.camera_info = self.get_cameras_windows()

    def get_cameras_windows(self):
        camera_info = []
        try:
            wmi = win32com.client.GetObject("winmgmts:")
            devices = wmi.InstancesOf("Win32_PnPEntity")
            idx = 0
            for device in devices:
                device_name = getattr(device, 'Name', None)
                if device_name and ('Camera' in device_name or 'Video' in device_name):
                    camera_info.append({'index': idx, 'name': device_name})
                    idx += 1
            if not camera_info:
                # Fallback to OpenCV detection
                camera_info = self.get_cameras_opencv()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get camera info: {e}")
            camera_info = self.get_cameras_opencv()
        return camera_info

    def get_cameras_opencv(self, max_tested=5):
        camera_info = []
        for i in range(max_tested):
            cap = cv2.VideoCapture(i)
            if cap.isOpened():
                ret, _ = cap.read()
                if ret:
                    camera_info.append({'index': i, 'name': f"Camera {i}"})
                cap.release()
        return camera_info

    def open_camera(self, index):
        if self.cap is not None:
            self.cap.release()
        self.cap = cv2.VideoCapture(index)
        if not self.cap.isOpened():
            # Check for applications using the webcam
            self.cap.release()
            self.cap = None
            conflicting_apps = self.find_conflicting_apps()
            if conflicting_apps:
                app_list = "\n".join(conflicting_apps)
                response = messagebox.askyesno(
                    "Webcam In Use",
                    f"The webcam is currently in use by the following applications:\n\n{app_list}\n\n"
                    "Do you want to close these applications?"
                )
                if response:
                    self.close_conflicting_apps(conflicting_apps)
                    # Try to open the camera again
                    self.cap = cv2.VideoCapture(index)
                    if not self.cap.isOpened():
                        messagebox.showerror("Error", "Could not open camera after closing applications.")
                        self.cap.release()
                        self.cap = None
                else:
                    messagebox.showinfo("Info", "Please close the applications manually and try again.")
            else:
                messagebox.showerror(
                    "Error",
                    f"Could not open camera {index}. Please check if the camera is connected and permissions are granted."
                )
        if self.cap and self.cap.isOpened():
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

    def find_conflicting_apps(self):
        # List of common applications that may use the webcam
        common_apps = [
            'Teams.exe',
            'Zoom.exe',
            'Skype.exe',
            'Discord.exe',
            'chrome.exe',
            'firefox.exe',
            'opera.exe',
            'obs64.exe',
            'CameraApp.exe',
            'ManyCam.exe'
        ]
        conflicting_apps = []
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] in common_apps:
                    conflicting_apps.append(proc.info['name'])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return conflicting_apps

    def close_conflicting_apps(self, apps):
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] in apps:
                    proc.terminate()
                    proc.wait(timeout=5)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                continue

    def on_camera_selection(self, *args):
        selected_name = self.camera_var.get()
        for cam in self.camera_info:
            if cam['name'] == selected_name:
                self.selected_camera_index = cam['index']
                break
        self.open_camera(self.selected_camera_index)

    def start_schedule(self):
        self.scheduled_mode = True
        self.override_mode = False
        messagebox.showinfo("Schedule", "Scheduled motion detection started.")

    def stop_schedule(self):
        self.scheduled_mode = False
        messagebox.showinfo("Schedule", "Scheduled motion detection stopped.")

    def start_override(self):
        self.override_mode = True
        self.scheduled_mode = False
        messagebox.showinfo("Override", "Override motion detection started.")

    def stop_override(self):
        self.override_mode = False
        messagebox.showinfo("Override", "Override motion detection stopped.")

    def reset_counter(self):
        self.motion_counter = 0
        self.counter_label.config(text=f"Motion Detections Today: {self.motion_counter}")
        messagebox.showinfo("Counter", "Motion detection counter reset.")

    def exit_application(self):
        if self.cap is not None:
            self.cap.release()
        cv2.destroyAllWindows()
        self.root.quit()

    def process_frame(self):
        # Schedule next frame processing
        self.root.after(30, self.process_frame)  # Adjusted delay for better performance

        # Check if we should be running
        if not self.should_run():
            return

        # Check if camera is available
        if self.cap is None:
            # Optionally, you can display a message or update the GUI to indicate no camera is selected
            return

        # Capture frame
        ret, frame = self.cap.read()
        if not ret:
            messagebox.showerror("Error", "Failed to capture image. Please ensure the camera is connected and permissions are granted.")
            if self.cap is not None:
                self.cap.release()
            self.cap = None
            return  # Do not exit application, allow user to select a different camera

        # Convert the frame to grayscale and apply GaussianBlur
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (21, 21), 0)

        # Initialize the first frame
        if self.first_frame is None:
            self.first_frame = gray
            return

        # Compute the absolute difference between the current frame and the first frame
        delta_frame = cv2.absdiff(self.first_frame, gray)
        thresh_frame = cv2.threshold(delta_frame, 30, 255, cv2.THRESH_BINARY)[1]
        thresh_frame = cv2.dilate(thresh_frame, None, iterations=2)

        # Find contours
        contours, _ = cv2.findContours(thresh_frame.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # Add the current frame to the buffer
        self.frame_buffer.append(frame.copy())

        motion_detected_in_frame = False

        # Check for motion
        for contour in contours:
            if cv2.contourArea(contour) < 10000:  # Adjust sensitivity as needed
                continue
            (x, y, w, h) = cv2.boundingRect(contour)
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
            motion_detected_in_frame = True

        # Start recording if motion is detected
        if motion_detected_in_frame:
            if self.video_writer is None:
                self.start_recording()
            self.save_buffer()  # Save the buffered frames
            self.video_writer.write(frame)  # Write the current frame
            self.motion_detected = True
            self.increment_counter()
        else:
            # Reset motion detection if no motion was detected
            if self.video_writer is not None:
                self.video_writer.release()
                self.video_writer = None
                self.save_to_onedrive()  # Save the detected motion video to OneDrive
                self.temp_filename = None

        # Check disk space
        if time.time() - self.last_flush_time >= self.flush_interval:
            if self.check_disk_space() < self.min_free_space_gb:
                messagebox.showwarning("Disk Space", "Low disk space. Stopping recording to prevent issues.")
                # Instead of exiting, we can stop recording
                if self.video_writer is not None:
                    self.video_writer.release()
                    self.video_writer = None
                return
            self.last_flush_time = time.time()

        # Display the video frame in the GUI
        cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(cv2image)
        imgtk = ImageTk.PhotoImage(image=img)
        self.video_label.imgtk = imgtk
        self.video_label.configure(image=imgtk)

    def should_run(self):
        if self.override_mode:
            return True
        elif self.scheduled_mode and self.is_within_active_hours():
            return True
        else:
            return False

    def is_within_active_hours(self):
        # Active between 9AM and 6PM on weekdays
        start_hour = 9
        end_hour = 18
        active_days = [0, 1, 2, 3, 4]  # Monday to Friday
        current_time = datetime.now()
        current_hour = current_time.hour
        current_day = current_time.weekday()
        return start_hour <= current_hour < end_hour and current_day in active_days

    def start_recording(self):
        fourcc = cv2.VideoWriter_fourcc(*'XVID')
        current_time_str = time.strftime("%Y-%m-%d-%H-%M-%S")
        self.temp_filename = f'temp_motion_{current_time_str}.avi'
        self.video_writer = cv2.VideoWriter(self.temp_filename, fourcc, 20.0, (640, 480))

    def save_buffer(self):
        if self.video_writer is not None:
            while self.frame_buffer:
                frame = self.frame_buffer.popleft()
                self.video_writer.write(frame)

    def save_to_onedrive(self):
        if self.temp_filename is not None and os.path.exists(self.temp_filename):
            destination_path = os.path.join(self.onedrive_videos_folder, os.path.basename(self.temp_filename))
            try:
                os.rename(self.temp_filename, destination_path)
            except OSError:
                # If os.rename fails (e.g., moving across different file systems), use shutil.move
                shutil.move(self.temp_filename, destination_path)
            print(f"Saved detected motion video to: {destination_path}")
        self.temp_filename = None

    def delete_old_videos(self):
        now = time.time()
        three_days_ago = now - (3 * 24 * 60 * 60)  # Timestamp for 3 days ago
        for filename in os.listdir(self.onedrive_videos_folder):
            file_path = os.path.join(self.onedrive_videos_folder, filename)
            if os.path.isfile(file_path):
                file_creation_time = os.path.getctime(file_path)
                if file_creation_time < three_days_ago:
                    os.remove(file_path)
                    print(f"Deleted old video: {file_path}")

    def schedule_delete_old_videos(self):
        self.delete_old_videos()
        # Schedule to run again after 5 minutes (300,000 ms)
        self.root.after(300000, self.schedule_delete_old_videos)

    def reset_first_frame(self):
        self.first_frame = None
        # Schedule to run again after 5 minutes
        self.root.after(300000, self.reset_first_frame)

    def check_disk_space(self):
        total, used, free = shutil.disk_usage('C:\\')
        free_space_gb = free / (1024 ** 3)
        return free_space_gb

    def increment_counter(self):
        # Reset counter daily
        today = datetime.now().date()
        if today != self.last_reset_date:
            self.motion_counter = 0
            self.last_reset_date = today
        self.motion_counter += 1
        self.counter_label.config(text=f"Motion Detections Today: {self.motion_counter}")

if __name__ == "__main__":
    root = tk.Tk()
    app = MotionDetectorApp(root)
    root.mainloop()
