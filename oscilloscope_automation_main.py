"""
Keysight Oscilloscope Automation Application
Professional automation interface for capturing screenshots, exporting data, and generating graphs
Built using the instrument_control package
"""

# These lines import necessary libraries for the program to work.

import sys  # Used for system-level operations, like exiting the program.
import logging  # Used for logging events and errors.
from pathlib import Path  # Helps in handling file paths in a way that works on any operating system.
from typing import Optional, Dict, Any, List, Tuple  # Provides type hints for better code readability.
import tkinter as tk  # This is the main library for creating the graphical user interface (GUI).
from tkinter import ttk, messagebox, filedialog, scrolledtext  # These are additional components for the GUI.
import pandas as pd  # A powerful library for data manipulation, used here to create CSV files.
import matplotlib.pyplot as plt  # Used for creating plots and graphs.
import numpy as np  # A library for numerical operations, especially with arrays of data.
from datetime import datetime  # Used to get the current date and time, for timestamps in filenames.
import threading  # Allows running multiple tasks at the same time, so the GUI doesn't freeze.
import queue  # A tool to safely pass messages between different threads.

# This section tries to import the custom modules needed to control the oscilloscope.
# If these modules are not found, it prints an error and exits the program.
try:
    # Imports the main oscilloscope control class and a custom error class.
    from instrument_control.keysight_oscilloscope import KeysightDSOX6004A, KeysightDSOX6004AError
    # Imports a helper class that sends commands to the instrument.
    from instrument_control.scpi_wrapper import SCPIWrapper
except ImportError as e:
    # This code runs if the imports fail.
    print(f"Error importing instrument control modules: {e}")
    print("Please ensure the instrument_control package is in your Python path")
    sys.exit(1)  # Exits the program.


# This class handles all tasks related to getting data from the oscilloscope.
class OscilloscopeDataAcquisition:
    """Manages acquiring, exporting, and plotting waveform data from the oscilloscope."""

    # This is the constructor method. It runs when a new OscilloscopeDataAcquisition object is created.
    def __init__(self, oscilloscope_instance):
        # Stores the oscilloscope object so it can be used by other methods in this class.
        self.scope = oscilloscope_instance
        # Sets up a logger to record events and errors for this class.
        self._logger = logging.getLogger(f'{self.__class__.__name__}')

    # This method gets the waveform (the signal shape) from a specific channel on the oscilloscope.
    def acquire_waveform_data(self, channel: int, max_points: int = 62500) -> Optional[Dict[str, Any]]:
        """
        Acquires waveform data from the specified channel.
        Returns a dictionary containing the time and voltage data, plus metadata.
        """
        # First, check if the oscilloscope is connected. If not, log an error and stop.
        if not self.scope.is_connected:
            self._logger.error("Cannot acquire data: oscilloscope not connected")
            return None

        try:
            # These commands configure the oscilloscope to send the waveform data.
            self.scope._scpi_wrapper.write(f":WAVeform:SOURce CHANnel{channel}")  # Select the channel.
            self.scope._scpi_wrapper.write(":WAVeform:FORMat BYTE")  # Set the data format.
            self.scope._scpi_wrapper.write(":WAVeform:POINts:MODE RAW")  # Get the raw data points.
            self.scope._scpi_wrapper.write(f":WAVeform:POINts {max_points}")  # Set how many data points to get.

            # This command gets metadata (the "preamble") about the waveform, which is needed to correctly interpret the data.
            preamble = self.scope._scpi_wrapper.query(":WAVeform:PREamble?")
            preamble_parts = preamble.split(',')  # The preamble is a comma-separated string.

            # These lines extract scaling factors from the preamble to convert raw data into actual voltage and time values.
            y_increment = float(preamble_parts[7])  # Voltage step per data point.
            y_origin = float(preamble_parts[8])     # Voltage offset.
            y_reference = float(preamble_parts[9])  # Reference voltage level.
            x_increment = float(preamble_parts[4])  # Time step per data point.
            x_origin = float(preamble_parts[5])     # Time offset.

            # This command requests the actual waveform data from the oscilloscope.
            raw_data = self.scope._scpi_wrapper.query_binary_values(":WAVeform:DATA?", datatype='B')

            # This line converts the raw data from the scope into meaningful voltage values using the scaling factors.
            voltage_data = [(value - y_reference) * y_increment + y_origin for value in raw_data]

            # This line creates a corresponding time value for each voltage point.
            time_data = [x_origin + (i * x_increment) for i in range(len(voltage_data))]

            self._logger.info(f"Successfully acquired {len(voltage_data)} points from channel {channel}")

            # The method returns a dictionary containing all the useful information.
            return {
                'channel': channel,
                'time': time_data,  # The array of time values.
                'voltage': voltage_data,  # The array of voltage values.
                'sample_rate': 1.0 / x_increment,  # How many samples were taken per second.
                'time_increment': x_increment,  # The time between each sample.
                'voltage_increment': y_increment,  # The voltage resolution.
                'points_count': len(voltage_data),  # The total number of data points.
                'acquisition_time': datetime.now().isoformat()  # When the data was acquired.
            }

        except Exception as e:
            # If anything goes wrong, log the error and return nothing.
            self._logger.error(f"Failed to acquire waveform data from channel {channel}: {e}")
            return None

    # This method saves the acquired waveform data to a CSV (Comma-Separated Values) file.
    def export_to_csv(self, waveform_data: Dict[str, Any], filename: Optional[str] = None) -> Optional[str]:
        """Exports the provided waveform data into a CSV file with metadata."""
        # If there's no data to export, log an error and stop.
        if not waveform_data:
            self._logger.error("No waveform data to export")
            return None

        try:
            # Make sure the output directory for data files exists.
            self.scope.setup_output_directories()

            # If no filename is provided, create one automatically with a timestamp.
            if filename is None:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"waveform_ch{waveform_data['channel']}_{timestamp}.csv"

            # Ensure the filename ends with .csv.
            if not filename.endswith('.csv'):
                filename += '.csv'

            # Create the full path to the file.
            filepath = self.scope.data_dir / filename

            # Use the pandas library to organize the data into a table (DataFrame).
            df = pd.DataFrame({
                'Time (s)': waveform_data['time'],
                'Voltage (V)': waveform_data['voltage']
            })

            # Open the file and write the metadata as comments at the top.
            # The '#' makes these lines comments, so they won't be read as data by spreadsheet programs.
            with open(filepath, 'w') as f:
                f.write(f"# Oscilloscope Waveform Data\n")
                f.write(f"# Channel: {waveform_data['channel']}\n")
                f.write(f"# Acquisition Time: {waveform_data['acquisition_time']}\n")
                f.write(f"# Sample Rate: {waveform_data['sample_rate']:.2e} Hz\n")
                f.write(f"# Points Count: {waveform_data['points_count']}\n")
                f.write(f"# Time Increment: {waveform_data['time_increment']:.2e} s\n")
                f.write(f"# Voltage Increment: {waveform_data['voltage_increment']:.2e} V\n")
                f.write("\n")

            # Append the actual data to the file, after the metadata comments.
            df.to_csv(filepath, mode='a', index=False)

            self._logger.info(f"CSV exported successfully: {filepath}")
            return str(filepath)  # Return the path of the created file.

        except Exception as e:
            # If anything goes wrong, log the error and return nothing.
            self._logger.error(f"Failed to export CSV: {e}")
            return None

    # This method creates a plot of the waveform data and saves it as an image file (e.g., .png).
    def generate_waveform_plot(self, waveform_data: Dict[str, Any], filename: Optional[str] = None, 
                             plot_title: Optional[str] = None) -> Optional[str]:
        """Generates and saves a plot of the waveform data."""
        # If there's no data, log an error and stop.
        if not waveform_data:
            self._logger.error("No waveform data to plot")
            return None

        try:
            # Ensure the output directory for graphs exists.
            self.scope.setup_output_directories()

            # If no filename is given, create one with a timestamp.
            if filename is None:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"waveform_plot_ch{waveform_data['channel']}_{timestamp}.png"

            # Make sure the filename has a valid image extension.
            if not filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                filename += '.png'

            # Create the full path for the plot image.
            filepath = self.scope.graph_dir / filename

            # Create the plot using matplotlib.
            plt.figure(figsize=(12, 8))  # Set the size of the plot.
            plt.plot(waveform_data['time'], waveform_data['voltage'], 'b-', linewidth=1)  # Plot time vs. voltage.

            # If no title is provided, create a default one.
            if plot_title is None:
                plot_title = f"Oscilloscope Waveform - Channel {waveform_data['channel']}"

            # Set the title and labels for the plot.
            plt.title(plot_title, fontsize=14, fontweight='bold')
            plt.xlabel('Time (s)', fontsize=12)
            plt.ylabel('Voltage (V)', fontsize=12)
            plt.grid(True, alpha=0.3)  # Add a grid to the background.

            # Calculate some basic statistics about the waveform.
            voltage_array = np.array(waveform_data['voltage'])
            stats_text = f"""Statistics:
Max: {np.max(voltage_array):.3f} V
Min: {np.min(voltage_array):.3f} V
Mean: {np.mean(voltage_array):.3f} V
RMS: {np.sqrt(np.mean(voltage_array**2)):.3f} V
Std Dev: {np.std(voltage_array):.3f} V
Points: {len(voltage_array)}"""

            # Add the statistics as a text box onto the plot.
            plt.text(0.02, 0.98, stats_text, transform=plt.gca().transAxes, 
                    fontsize=10, verticalalignment='top',
                    bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))

            plt.tight_layout()  # Adjust the plot to fit everything nicely.
            plt.savefig(filepath, dpi=300, bbox_inches='tight')  # Save the plot to the file.
            plt.close()  # Close the plot to free up memory.

            self._logger.info(f"Plot saved successfully: {filepath}")
            return str(filepath)  # Return the path of the saved plot.

        except Exception as e:
            # If anything goes wrong, log the error and return nothing.
            self._logger.error(f"Failed to generate plot: {e}")
            return None


# This class defines the entire Graphical User Interface (GUI) for the application.
class AutomationGUI:
    """Manages the graphical user interface for the oscilloscope automation tool."""

    # The constructor method, which runs when the GUI is created.
    def __init__(self):
        self.root = tk.Tk()  # Creates the main window of the application.
        self.oscilloscope = None  # A variable to hold the oscilloscope object once connected.
        self.data_acquisition = None  # A variable to hold the data acquisition object.
        self.setup_logging()  # Sets up the logging system.
        self.setup_gui()  # Calls the method to build all the GUI components.
        self.status_queue = queue.Queue()  # Creates a queue for thread-safe communication.
        self.check_status_updates()  # Starts a loop to check for updates from other threads.

    # This method configures the logging system for the application.
    def setup_logging(self):
        """Configures the format and level for application-wide logging."""
        logging.basicConfig(
            level=logging.INFO,  # Set the minimum level of messages to log (e.g., INFO, WARNING, ERROR).
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'  # Define the format for log messages.
        )
        self.logger = logging.getLogger('OscilloscopeAutomation')  # Get a logger instance for the GUI.

    # This method sets up the main structure of the GUI.
    def setup_gui(self):
        """Initializes and arranges all the main components of the GUI window."""
        self.root.title("Keysight Oscilloscope Automation Suite")  # Set the title of the window.
        self.root.geometry("900x700")  # Set the initial size of the window.
        self.root.configure(bg='#f0f0f0')  # Set the background color.

        # Configure the visual style of the GUI elements.
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use a modern theme.
        self.configure_styles()  # Apply custom styles for buttons and labels.

        # Create a main container frame for all other widgets.
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Add a title label at the top of the window.
        title_label = ttk.Label(main_frame, text="Oscilloscope Automation Control Panel", 
                               style='Title.TLabel')
        title_label.pack(pady=(0, 20))

        # Create the different sections of the GUI by calling other methods.
        self.create_connection_frame(main_frame)  # Section for connection settings.
        self.create_operations_frame(main_frame)  # Section for action buttons.
        self.create_status_frame(main_frame)  # Section for status messages and logs.

    # This method defines custom styles for various GUI widgets to make them look better.
    def configure_styles(self):
        """Defines custom fonts and colors for GUI elements like titles and buttons."""
        # Style for the main title label.
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'), 
                           foreground='#2c5aa0')

        # Style for buttons that indicate a successful or primary action.
        self.style.configure('Success.TButton', font=('Arial', 10, 'bold'))
        self.style.map('Success.TButton', 
                      background=[('active', '#28a745'), ('!active', '#198754')])

        # Style for buttons that indicate a warning or destructive action.
        self.style.configure('Warning.TButton', font=('Arial', 10, 'bold'))
        self.style.map('Warning.TButton',
                      background=[('active', '#fd7e14'), ('!active', '#e36209')])

        # Style for informational buttons.
        self.style.configure('Info.TButton', font=('Arial', 10, 'bold'))
        self.style.map('Info.TButton',
                      background=[('active', '#17a2b8'), ('!active', '#138496')])

    # This method creates the section of the GUI for managing the connection to the oscilloscope.
    def create_connection_frame(self, parent):
        """Creates the UI elements for setting the VISA address and connecting/disconnecting."""
        conn_frame = ttk.LabelFrame(parent, text="Connection Settings", padding="15")
        conn_frame.pack(fill=tk.X, pady=(0, 20))

        # Create a frame for the VISA address input field.
        address_frame = ttk.Frame(conn_frame)
        address_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(address_frame, text="VISA Address:", font=('Arial', 10)).pack(side=tk.LEFT)
        # A special variable to hold the text of the entry box.
        self.visa_address_var = tk.StringVar(value="TCPIP0::192.168.1.100::inst0::INSTR")
        self.visa_entry = ttk.Entry(address_frame, textvariable=self.visa_address_var, width=40)
        self.visa_entry.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)

        # Create a frame for the Connect and Disconnect buttons.
        button_frame = ttk.Frame(conn_frame)
        button_frame.pack(fill=tk.X)

        # The 'Connect' button. When clicked, it calls the `connect_oscilloscope` method.
        self.connect_btn = ttk.Button(button_frame, text="Connect", 
                                    command=self.connect_oscilloscope, style='Success.TButton')
        self.connect_btn.pack(side=tk.LEFT, padx=(0, 10))

        # The 'Disconnect' button. It's disabled by default.
        self.disconnect_btn = ttk.Button(button_frame, text="Disconnect", 
                                       command=self.disconnect_oscilloscope, 
                                       style='Warning.TButton', state='disabled')
        self.disconnect_btn.pack(side=tk.LEFT)

        # A label to show the current connection status (e.g., "Connected" or "Disconnected").
        self.conn_status_var = tk.StringVar(value="Disconnected")
        self.conn_status_label = ttk.Label(button_frame, textvariable=self.conn_status_var, 
                                         font=('Arial', 10, 'bold'), foreground='red')
        self.conn_status_label.pack(side=tk.RIGHT)

    # This method creates the main control buttons for the application.
    def create_operations_frame(self, parent):
        """Creates the UI elements for selecting a channel and running operations."""
        ops_frame = ttk.LabelFrame(parent, text="Operations", padding="15")
        ops_frame.pack(fill=tk.X, pady=(0, 20))

        # Create a dropdown menu for selecting the oscilloscope channel.
        channel_frame = ttk.Frame(ops_frame)
        channel_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(channel_frame, text="Channel:", font=('Arial', 10)).pack(side=tk.LEFT)
        self.channel_var = tk.IntVar(value=1)  # Variable to hold the selected channel number.
        channel_combo = ttk.Combobox(channel_frame, textvariable=self.channel_var, 
                                   values=[1, 2, 3, 4], width=10, state='readonly')
        channel_combo.pack(side=tk.LEFT, padx=(10, 0))

        # Create a grid to organize the operation buttons.
        button_grid = ttk.Frame(ops_frame)
        button_grid.pack(fill=tk.X)

        # --- First row of buttons ---
        self.screenshot_btn = ttk.Button(button_grid, text="Capture Screenshot", 
                                       command=self.capture_screenshot, style='Info.TButton')
        self.screenshot_btn.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

        self.acquire_data_btn = ttk.Button(button_grid, text="Acquire Waveform Data", 
                                         command=self.acquire_data, style='Info.TButton')
        self.acquire_data_btn.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        self.export_csv_btn = ttk.Button(button_grid, text="Export to CSV", 
                                       command=self.export_csv, style='Success.TButton')
        self.export_csv_btn.grid(row=0, column=2, padx=5, pady=5, sticky='ew')

        # --- Second row of buttons ---
        self.generate_plot_btn = ttk.Button(button_grid, text="Generate Plot", 
                                          command=self.generate_plot, style='Success.TButton')
        self.generate_plot_btn.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

        self.full_automation_btn = ttk.Button(button_grid, text="Full Automation", 
                                            command=self.run_full_automation, 
                                            style='Success.TButton')
        self.full_automation_btn.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

        self.open_folder_btn = ttk.Button(button_grid, text="Open Output Folder", 
                                        command=self.open_output_folder, style='Info.TButton')
        self.open_folder_btn.grid(row=1, column=2, padx=5, pady=5, sticky='ew')

        # Make the buttons expand to fill the available space.
        for i in range(3):
            button_grid.columnconfigure(i, weight=1)

        # Disable all these buttons until a connection is made.
        self.disable_operation_buttons()

    # This method creates the area at the bottom of the GUI for showing status and log messages.
    def create_status_frame(self, parent):
        """Creates the UI elements for displaying the current status and activity log."""
        status_frame = ttk.LabelFrame(parent, text="Status & Logs", padding="15")
        status_frame.pack(fill=tk.BOTH, expand=True)

        # A label to show the status of the current operation (e.g., "Connecting...", "Ready").
        self.current_operation_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, text="Current Operation:", font=('Arial', 10)).pack(anchor='w')
        operation_label = ttk.Label(status_frame, textvariable=self.current_operation_var, 
                                  font=('Arial', 10, 'bold'), foreground='#2c5aa0')
        operation_label.pack(anchor='w', pady=(0, 10))

        # A text box with a scrollbar to display a running log of all actions and errors.
        ttk.Label(status_frame, text="Activity Log:", font=('Arial', 10)).pack(anchor='w')
        self.log_text = scrolledtext.ScrolledText(status_frame, height=10, font=('Courier', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # A button to clear all messages from the log box.
        ttk.Button(status_frame, text="Clear Log", command=self.clear_log).pack(anchor='e', pady=(5, 0))

    def log_message(self, message: str, level: str = "INFO"):
        """Add message to log display"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}\n"

        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)

        # Color coding for different levels
        if level == "ERROR":
            self.log_text.tag_add("error", f"end-{len(log_entry)}c", "end-1c")
            self.log_text.tag_config("error", foreground="red")
        elif level == "SUCCESS":
            self.log_text.tag_add("success", f"end-{len(log_entry)}c", "end-1c")
            self.log_text.tag_config("success", foreground="green")

    def clear_log(self):
        """Clear the log display"""
        self.log_text.delete(1.0, tk.END)

    def update_status(self, status: str):
        """Update current operation status"""
        self.current_operation_var.set(status)
        self.root.update_idletasks()

    def disable_operation_buttons(self):
        """Disable all operation buttons"""
        buttons = [self.screenshot_btn, self.acquire_data_btn, self.export_csv_btn, 
                  self.generate_plot_btn, self.full_automation_btn, self.open_folder_btn]
        for btn in buttons:
            btn.configure(state='disabled')

    def enable_operation_buttons(self):
        """Enable all operation buttons"""
        buttons = [self.screenshot_btn, self.acquire_data_btn, self.export_csv_btn, 
                  self.generate_plot_btn, self.full_automation_btn, self.open_folder_btn]
        for btn in buttons:
            btn.configure(state='normal')

    def connect_oscilloscope(self):
        """Connect to the oscilloscope"""
        def connect_thread():
            try:
                self.update_status("Connecting...")
                self.log_message("Attempting to connect to oscilloscope...")

                visa_address = self.visa_address_var.get().strip()
                if not visa_address:
                    raise ValueError("VISA address cannot be empty")

                self.oscilloscope = KeysightDSOX6004A(visa_address)

                if self.oscilloscope.connect():
                    self.data_acquisition = OscilloscopeDataAcquisition(self.oscilloscope)

                    # Get instrument info
                    info = self.oscilloscope.get_instrument_info()
                    if info:
                        self.log_message(f"Connected to {info['manufacturer']} {info['model']}", "SUCCESS")
                        self.log_message(f"Serial: {info['serial_number']}, Firmware: {info['firmware_version']}")

                    self.status_queue.put(("connected", None))
                else:
                    raise Exception("Failed to establish connection")

            except Exception as e:
                self.status_queue.put(("error", f"Connection failed: {str(e)}"))

        threading.Thread(target=connect_thread, daemon=True).start()

    def disconnect_oscilloscope(self):
        """Disconnect from the oscilloscope"""
        try:
            if self.oscilloscope:
                self.oscilloscope.disconnect()
                self.oscilloscope = None
                self.data_acquisition = None

            self.conn_status_var.set("Disconnected")
            self.conn_status_label.configure(foreground='red')
            self.connect_btn.configure(state='normal')
            self.disconnect_btn.configure(state='disabled')
            self.disable_operation_buttons()

            self.update_status("Disconnected")
            self.log_message("Disconnected from oscilloscope", "SUCCESS")

        except Exception as e:
            self.log_message(f"Error during disconnection: {e}", "ERROR")

    def capture_screenshot(self):
        """Capture oscilloscope screenshot"""
        def screenshot_thread():
            try:
                self.update_status("Capturing screenshot...")
                self.log_message("Starting screenshot capture...")

                filename = self.oscilloscope.capture_screenshot()
                if filename:
                    self.status_queue.put(("screenshot_success", filename))
                else:
                    self.status_queue.put(("error", "Screenshot capture failed"))

            except Exception as e:
                self.status_queue.put(("error", f"Screenshot error: {str(e)}"))

        if self.oscilloscope and self.oscilloscope.is_connected:
            threading.Thread(target=screenshot_thread, daemon=True).start()
        else:
            messagebox.showerror("Error", "Oscilloscope not connected")

    def acquire_data(self):
        """Acquire waveform data"""
        def acquire_thread():
            try:
                self.update_status("Acquiring waveform data...")
                self.log_message(f"Acquiring data from channel {self.channel_var.get()}...")

                data = self.data_acquisition.acquire_waveform_data(self.channel_var.get())
                if data:
                    self.status_queue.put(("data_acquired", data))
                else:
                    self.status_queue.put(("error", "Data acquisition failed"))

            except Exception as e:
                self.status_queue.put(("error", f"Data acquisition error: {str(e)}"))

        if self.data_acquisition:
            threading.Thread(target=acquire_thread, daemon=True).start()
        else:
            messagebox.showerror("Error", "Oscilloscope not connected")

    def export_csv(self):
        """Export last acquired data to CSV"""
        if not hasattr(self, 'last_acquired_data') or not self.last_acquired_data:
            messagebox.showwarning("Warning", "No data available. Please acquire data first.")
            return

        def export_thread():
            try:
                self.update_status("Exporting to CSV...")
                self.log_message("Exporting waveform data to CSV...")

                filename = self.data_acquisition.export_to_csv(self.last_acquired_data)
                if filename:
                    self.status_queue.put(("csv_exported", filename))
                else:
                    self.status_queue.put(("error", "CSV export failed"))

            except Exception as e:
                self.status_queue.put(("error", f"CSV export error: {str(e)}"))

        threading.Thread(target=export_thread, daemon=True).start()

    def generate_plot(self):
        """Generate plot from last acquired data"""
        if not hasattr(self, 'last_acquired_data') or not self.last_acquired_data:
            messagebox.showwarning("Warning", "No data available. Please acquire data first.")
            return

        def plot_thread():
            try:
                self.update_status("Generating plot...")
                self.log_message("Generating waveform plot...")

                filename = self.data_acquisition.generate_waveform_plot(self.last_acquired_data)
                if filename:
                    self.status_queue.put(("plot_generated", filename))
                else:
                    self.status_queue.put(("error", "Plot generation failed"))

            except Exception as e:
                self.status_queue.put(("error", f"Plot generation error: {str(e)}"))

        threading.Thread(target=plot_thread, daemon=True).start()

    def run_full_automation(self):
        """Run complete automation sequence"""
        def full_automation_thread():
            try:
                channel = self.channel_var.get()

                # Step 1: Capture screenshot
                self.update_status("Full Automation: Capturing screenshot...")
                self.log_message("Starting full automation sequence...")
                self.log_message("Step 1/4: Capturing screenshot...")

                screenshot_file = self.oscilloscope.capture_screenshot()
                if not screenshot_file:
                    raise Exception("Screenshot capture failed")

                # Step 2: Acquire data
                self.update_status("Full Automation: Acquiring data...")
                self.log_message("Step 2/4: Acquiring waveform data...")

                data = self.data_acquisition.acquire_waveform_data(channel)
                if not data:
                    raise Exception("Data acquisition failed")

                # Step 3: Export CSV
                self.update_status("Full Automation: Exporting CSV...")
                self.log_message("Step 3/4: Exporting to CSV...")

                csv_file = self.data_acquisition.export_to_csv(data)
                if not csv_file:
                    raise Exception("CSV export failed")

                # Step 4: Generate plot
                self.update_status("Full Automation: Generating plot...")
                self.log_message("Step 4/4: Generating plot...")

                plot_file = self.data_acquisition.generate_waveform_plot(data)
                if not plot_file:
                    raise Exception("Plot generation failed")

                results = {
                    'screenshot': screenshot_file,
                    'csv': csv_file,
                    'plot': plot_file,
                    'data': data
                }

                self.status_queue.put(("full_automation_complete", results))

            except Exception as e:
                self.status_queue.put(("error", f"Full automation error: {str(e)}"))

        if self.data_acquisition:
            threading.Thread(target=full_automation_thread, daemon=True).start()
        else:
            messagebox.showerror("Error", "Oscilloscope not connected")

    def open_output_folder(self):
        """Open the output folder in file explorer"""
        try:
            if self.oscilloscope:
                self.oscilloscope.setup_output_directories()
                import subprocess
                import platform

                base_path = Path.cwd()

                if platform.system() == "Windows":
                    subprocess.run(['explorer', str(base_path)], check=True)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(['open', str(base_path)], check=True)
                else:  # Linux
                    subprocess.run(['xdg-open', str(base_path)], check=True)

                self.log_message("Opened output folder")
            else:
                messagebox.showinfo("Info", "Connect to oscilloscope first to create output folders")

        except Exception as e:
            self.log_message(f"Error opening folder: {e}", "ERROR")

    def check_status_updates(self):
        """Check for status updates from background threads"""
        try:
            while True:
                status_type, data = self.status_queue.get_nowait()

                if status_type == "connected":
                    self.conn_status_var.set("Connected")
                    self.conn_status_label.configure(foreground='green')
                    self.connect_btn.configure(state='disabled')
                    self.disconnect_btn.configure(state='normal')
                    self.enable_operation_buttons()
                    self.update_status("Connected - Ready for operations")

                elif status_type == "error":
                    self.log_message(data, "ERROR")
                    self.update_status("Error occurred")

                elif status_type == "screenshot_success":
                    self.log_message(f"Screenshot saved: {data}", "SUCCESS")
                    self.update_status("Screenshot captured successfully")

                elif status_type == "data_acquired":
                    self.last_acquired_data = data
                    self.log_message(f"Acquired {data['points_count']} data points from channel {data['channel']}", "SUCCESS")
                    self.update_status("Data acquisition completed")

                elif status_type == "csv_exported":
                    self.log_message(f"CSV exported: {data}", "SUCCESS")
                    self.update_status("CSV export completed")

                elif status_type == "plot_generated":
                    self.log_message(f"Plot generated: {data}", "SUCCESS")
                    self.update_status("Plot generation completed")

                elif status_type == "full_automation_complete":
                    self.last_acquired_data = data['data']
                    self.log_message("Full automation sequence completed successfully!", "SUCCESS")
                    self.log_message(f"Files created:", "SUCCESS")
                    self.log_message(f"  Screenshot: {data['screenshot']}", "SUCCESS")
                    self.log_message(f"  CSV Data: {data['csv']}", "SUCCESS") 
                    self.log_message(f"  Plot: {data['plot']}", "SUCCESS")
                    self.update_status("Full automation completed")

                    messagebox.showinfo("Success", 
                                      f"Full automation completed successfully!\n\n"
                                      f"Files created:\n"
                                      f"• Screenshot: {Path(data['screenshot']).name}\n"
                                      f"• CSV Data: {Path(data['csv']).name}\n"
                                      f"• Plot: {Path(data['plot']).name}")

        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_status_updates)

    # This method starts the GUI application.
    def run(self):
        """Starts the main event loop of the Tkinter application."""
        try:
            self.log_message("Oscilloscope Automation Suite started")
            # This line starts the GUI, makes it visible, and waits for user interaction (like button clicks).
            self.root.mainloop()
        except KeyboardInterrupt:
            # This part is for gracefully handling when the user closes the program with Ctrl+C.
            self.log_message("Application interrupted by user")
        finally:
            # This code will run when the application is closing, for any reason.
            if self.oscilloscope:
                # Ensure the oscilloscope is properly disconnected to avoid leaving it in a bad state.
                self.oscilloscope.disconnect()


# This is the main function that starts the entire application.
def main():
    """Main application entry point."""
    try:
        # Create an instance of our GUI class.
        app = AutomationGUI()
        # Run the application.
        app.run()
    except Exception as e:
        # If a major, unhandled error occurs, print it to the console.
        print(f"Application error: {e}")
        input("Press Enter to exit...")


# This is a standard Python construct. It ensures that the `main()` function is called only when the script is executed directly.
if __name__ == "__main__":
    main()
