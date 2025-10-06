"""
Keysight Oscilloscope Automation Application

This module provides a comprehensive GUI application for automating Keysight oscilloscope operations.
It includes features for data acquisition, waveform analysis, screenshot capture, and file export.

Key Features:
- Responsive GUI layout that adapts to window resizing
- VISA-based instrument communication
- Multi-threaded operations to prevent UI freezing
- Customizable save locations for data, graphs, and screenshots
- Real-time logging with color-coded messages
- Full automation workflow (screenshot + data + CSV + plot)

Author: Professional Instrumentation Control System
Version: 2.0
"""

# Standard library imports for system operations and data structures
import sys
import logging
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
import os
import threading
import queue
from datetime import datetime

# Third-party imports for GUI, data processing, and visualization
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Import custom instrument control modules
# These modules handle low-level SCPI communication with the oscilloscope
try:
    from instrument_control.keysight_oscilloscope import KeysightDSOX6004A, KeysightDSOX6004AError
    from instrument_control.scpi_wrapper import SCPIWrapper
except ImportError as e:
    print(f"Error importing instrument control modules: {e}")
    print("Please ensure the instrument_control package is in your Python path")
    sys.exit(1)

class OscilloscopeDataAcquisition:
    """
    Data Acquisition Handler for Oscilloscope Operations
    
    This class manages all data acquisition, export, and visualization operations
    for the oscilloscope. It provides methods to:
    - Acquire raw waveform data from oscilloscope channels
    - Export data to CSV format with metadata
    - Generate publication-quality plots with statistics
    
    The class maintains default directories for organizing output files and supports
    custom save locations for flexibility.
    """

    def __init__(self, oscilloscope_instance):
        """
        Initialize the data acquisition handler.
        
        Args:
            oscilloscope_instance: Connected KeysightDSOX6004A oscilloscope object
        """
        # Store reference to the oscilloscope instance for SCPI communication
        self.scope = oscilloscope_instance
        
        # Create a logger instance for this class to track operations
        self._logger = logging.getLogger(f'{self.__class__.__name__}')
        
        # Define default output directories relative to current working directory
        self.default_data_dir = Path.cwd() / "data"
        self.default_graph_dir = Path.cwd() / "graphs"
        self.default_screenshot_dir = Path.cwd() / "screenshots"

    def acquire_waveform_data(self, channel: int, max_points: int = 62500) -> Optional[Dict[str, Any]]:
        """
        Acquire waveform data from a specified oscilloscope channel.
        
        This method configures the oscilloscope for data acquisition, retrieves the
        raw waveform data, and converts it to voltage and time arrays using the
        oscilloscope's scaling parameters.
        
        Args:
            channel: Channel number to acquire data from (1-4)
            max_points: Maximum number of data points to acquire (default: 62500)
            
        Returns:
            Dictionary containing waveform data and metadata, or None if acquisition fails.
            Dictionary keys:
                - 'channel': Channel number
                - 'time': List of time values in seconds
                - 'voltage': List of voltage values in volts
                - 'sample_rate': Sampling rate in Hz
                - 'time_increment': Time between samples in seconds
                - 'voltage_increment': Voltage resolution in volts
                - 'points_count': Number of data points acquired
                - 'acquisition_time': ISO format timestamp of acquisition
        """
        if not self.scope.is_connected:
            self._logger.error("Cannot acquire data: oscilloscope not connected")
            return None

        try:
            # Configure oscilloscope waveform source to the specified channel
            self.scope._scpi_wrapper.write(f":WAVeform:SOURce CHANnel{channel}")
            
            # Set data format to BYTE (8-bit unsigned integers) for efficient transfer
            self.scope._scpi_wrapper.write(":WAVeform:FORMat BYTE")
            
            # Use RAW mode to get all available data points from acquisition memory
            self.scope._scpi_wrapper.write(":WAVeform:POINts:MODE RAW")
            
            # Set the maximum number of points to retrieve
            self.scope._scpi_wrapper.write(f":WAVeform:POINts {max_points}")

            # Query the waveform preamble which contains scaling information
            # The preamble is a comma-separated list of 10 values describing the waveform
            preamble = self.scope._scpi_wrapper.query(":WAVeform:PREamble?")
            preamble_parts = preamble.split(',')

            # Extract scaling factors from preamble for converting raw data to real units
            # y_increment: voltage value per ADC count
            y_increment = float(preamble_parts[7])
            # y_origin: voltage at ADC reference point
            y_origin = float(preamble_parts[8])
            # y_reference: ADC reference value (typically 0)
            y_reference = float(preamble_parts[9])
            # x_increment: time between consecutive data points
            x_increment = float(preamble_parts[4])
            # x_origin: time of first data point
            x_origin = float(preamble_parts[5])

            # Retrieve the actual waveform data as binary values (unsigned bytes)
            raw_data = self.scope._scpi_wrapper.query_binary_values(":WAVeform:DATA?", datatype='B')

            # Convert raw ADC values to actual voltage using the formula:
            # Voltage = (ADC_value - reference) * increment + origin
            voltage_data = [(value - y_reference) * y_increment + y_origin for value in raw_data]
            
            # Generate time array: each point's time = origin + (index * increment)
            time_data = [x_origin + (i * x_increment) for i in range(len(voltage_data))]

            self._logger.info(f"Successfully acquired {len(voltage_data)} points from channel {channel}")

            # Return structured dictionary with all waveform data and metadata
            return {
                'channel': channel,
                'time': time_data,
                'voltage': voltage_data,
                'sample_rate': 1.0 / x_increment,  # Calculate sampling rate from time increment
                'time_increment': x_increment,
                'voltage_increment': y_increment,
                'points_count': len(voltage_data),
                'acquisition_time': datetime.now().isoformat()
            }

        except Exception as e:
            self._logger.error(f"Failed to acquire waveform data from channel {channel}: {e}")
            return None

    def export_to_csv(self, waveform_data: Dict[str, Any], custom_path: Optional[str] = None, 
                      filename: Optional[str] = None) -> Optional[str]:
        """
        Export waveform data to CSV file with metadata header.
        
        Creates a CSV file containing time and voltage data with a comprehensive
        metadata header that includes acquisition parameters. This format is suitable
        for post-processing in spreadsheet applications or data analysis tools.
        
        Args:
            waveform_data: Dictionary containing waveform data from acquire_waveform_data()
            custom_path: Optional custom directory path for saving the file
            filename: Optional custom filename (will auto-generate if not provided)
            
        Returns:
            Full path to the saved CSV file as a string, or None if export fails
        """
        if not waveform_data:
            self._logger.error("No waveform data to export")
            return None

        try:
            # Determine the save directory: use custom path if provided, otherwise use default
            if custom_path:
                save_dir = Path(custom_path)
            else:
                save_dir = self.default_data_dir
                # Ensure default output directories exist
                self.scope.setup_output_directories()

            # Create the directory if it doesn't exist (including parent directories)
            save_dir.mkdir(parents=True, exist_ok=True)

            # Generate a timestamped filename if not provided by user
            if filename is None:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"waveform_ch{waveform_data['channel']}_{timestamp}.csv"

            # Ensure filename has .csv extension
            if not filename.endswith('.csv'):
                filename += '.csv'

            # Construct full file path
            filepath = save_dir / filename

            # Create a pandas DataFrame with labeled columns for time and voltage
            df = pd.DataFrame({
                'Time (s)': waveform_data['time'],
                'Voltage (V)': waveform_data['voltage']
            })

            # Write metadata header as comments (lines starting with #)
            # This preserves important acquisition parameters with the data
            with open(filepath, 'w') as f:
                f.write(f"# Oscilloscope Waveform Data\n")
                f.write(f"# Channel: {waveform_data['channel']}\n")
                f.write(f"# Acquisition Time: {waveform_data['acquisition_time']}\n")
                f.write(f"# Sample Rate: {waveform_data['sample_rate']:.2e} Hz\n")
                f.write(f"# Points Count: {waveform_data['points_count']}\n")
                f.write(f"# Time Increment: {waveform_data['time_increment']:.2e} s\n")
                f.write(f"# Voltage Increment: {waveform_data['voltage_increment']:.2e} V\n")
                f.write("\n")

            # Append the actual data to the file (mode='a' for append)
            df.to_csv(filepath, mode='a', index=False)
            self._logger.info(f"CSV exported successfully: {filepath}")
            return str(filepath)

        except Exception as e:
            self._logger.error(f"Failed to export CSV: {e}")
            return None

    def generate_waveform_plot(self, waveform_data: Dict[str, Any], custom_path: Optional[str] = None,
                              filename: Optional[str] = None, plot_title: Optional[str] = None) -> Optional[str]:
        """
        Generate a publication-quality plot of waveform data with statistics.
        
        Creates a matplotlib figure showing the waveform with an embedded statistics box.
        The plot includes voltage vs time, grid lines, and calculated statistics
        (max, min, mean, RMS, standard deviation).
        
        Args:
            waveform_data: Dictionary containing waveform data from acquire_waveform_data()
            custom_path: Optional custom directory path for saving the plot
            filename: Optional custom filename (will auto-generate if not provided)
            plot_title: Optional custom title for the plot (auto-generates if not provided)
            
        Returns:
            Full path to the saved plot file as a string, or None if generation fails
        """
        if not waveform_data:
            self._logger.error("No waveform data to plot")
            return None

        try:
            # Determine the save directory: use custom path if provided, otherwise use default
            if custom_path:
                save_dir = Path(custom_path)
            else:
                save_dir = self.default_graph_dir
                # Ensure default output directories exist
                self.scope.setup_output_directories()

            # Create the directory if it doesn't exist (including parent directories)
            save_dir.mkdir(parents=True, exist_ok=True)

            # Generate a timestamped filename if not provided by user
            if filename is None:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                filename = f"waveform_plot_ch{waveform_data['channel']}_{timestamp}.png"

            # Ensure filename has an image extension
            if not filename.lower().endswith(('.png', '.jpg', '.jpeg')):
                filename += '.png'

            # Construct full file path
            filepath = save_dir / filename

            # Create a matplotlib figure with specified size (12x8 inches)
            plt.figure(figsize=(12, 8))
            
            # Plot the waveform data: blue line with 1-pixel width
            plt.plot(waveform_data['time'], waveform_data['voltage'], 'b-', linewidth=1)

            # Set the plot title: use custom title if provided, otherwise auto-generate
            if plot_title is None:
                plot_title = f"Oscilloscope Waveform - Channel {waveform_data['channel']}"

            # Configure plot labels and styling
            plt.title(plot_title, fontsize=14, fontweight='bold')
            plt.xlabel('Time (s)', fontsize=12)
            plt.ylabel('Voltage (V)', fontsize=12)
            plt.grid(True, alpha=0.3)  # Add semi-transparent grid for readability

            # Calculate statistical measures of the waveform
            voltage_array = np.array(waveform_data['voltage'])
            stats_text = f"""Statistics:
Max: {np.max(voltage_array):.3f} V
Min: {np.min(voltage_array):.3f} V
Mean: {np.mean(voltage_array):.3f} V
RMS: {np.sqrt(np.mean(voltage_array**2)):.3f} V
Std Dev: {np.std(voltage_array):.3f} V
Points: {len(voltage_array)}"""

            # Add statistics text box to the plot (top-left corner)
            # transform=plt.gca().transAxes uses axes coordinates (0-1 range)
            plt.text(0.02, 0.98, stats_text, transform=plt.gca().transAxes,
                    fontsize=10, verticalalignment='top',
                    bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))

            # Adjust layout to prevent label cutoff
            plt.tight_layout()
            
            # Save the figure at high resolution (300 DPI) for publication quality
            plt.savefig(filepath, dpi=300, bbox_inches='tight')
            
            # Close the figure to free memory
            plt.close()

            self._logger.info(f"Plot saved successfully: {filepath}")
            return str(filepath)

        except Exception as e:
            self._logger.error(f"Failed to generate plot: {e}")
            return None

class TrulyResponsiveAutomationGUI:
    """
    Main GUI Application for Oscilloscope Automation
    
    This class implements a professional Tkinter-based graphical user interface
    for controlling and automating Keysight oscilloscope operations. The GUI
    features a responsive layout that adapts to window resizing and provides
    comprehensive controls for all oscilloscope functions.
    
    Key Features:
    - Responsive grid-based layout that fills available space
    - Multi-threaded operations to prevent UI freezing
    - Real-time color-coded logging
    - Customizable save locations for all output files
    - Full automation workflow combining multiple operations
    - Connection management with status indicators
    
    Architecture:
    - Uses queue-based communication between worker threads and GUI thread
    - Implements proper thread safety for UI updates
    - Maintains separation between UI logic and instrument control
    """

    def __init__(self):
        """
        Initialize the GUI application and all its components.
        
        Sets up the main window, initializes instance variables, configures
        logging, creates the GUI layout, and starts the status update loop.
        """
        # Create the main Tkinter window
        self.root = tk.Tk()
        
        # Initialize oscilloscope-related objects (will be set when connected)
        self.oscilloscope = None  # KeysightDSOX6004A instance
        self.data_acquisition = None  # OscilloscopeDataAcquisition instance
        self.last_acquired_data = None  # Stores most recent waveform data
        
        # User preferences for save locations (can be customized via GUI)
        self.save_locations = {
            'data': str(Path.cwd() / "data"),
            'graphs': str(Path.cwd() / "graphs"),
            'screenshots': str(Path.cwd() / "screenshots")
        }

        # Configure application logging
        self.setup_logging()
        
        # Build the GUI layout
        self.setup_gui()
        
        # Create queue for thread-safe communication between worker threads and GUI
        self.status_queue = queue.Queue()
        
        # Start the periodic status update checker
        self.check_status_updates()

    def setup_logging(self):
        """
        Configure the logging system for the application.
        
        Sets up basic logging configuration with INFO level and a standard
        format that includes timestamp, logger name, level, and message.
        """
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger('OscilloscopeAutomation')
    
    def setup_gui(self):
        """
        Initialize the responsive GUI layout with all controls and widgets.
        
        Creates a grid-based layout that properly scales with window resizing.
        The layout uses a hierarchical structure where the main frame contains
        multiple sub-frames for different functional areas. Grid weights are
        carefully configured to ensure the log area expands while control areas
        remain fixed size.
        """
        # Set window title
        self.root.title("Keysight Oscilloscope Automation")

        # Set initial window size and background color
        self.root.geometry("1000x650")  # Width x Height in pixels
        self.root.configure(bg='#f5f5f5')  # Light gray background
        self.root.minsize(800, 500)  # Minimum window dimensions

        # Configure root window grid to allow expansion
        # weight=1 means the column/row will expand to fill available space
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Configure ttk styling for professional appearance
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use 'clam' theme for modern look
        self.configure_styles()

        # Create main container frame that fills the entire window
        # padding="3" adds 3 pixels of internal spacing
        # sticky='nsew' makes the frame expand in all directions (North, South, East, West)
        main_frame = ttk.Frame(self.root, padding="3")
        main_frame.grid(row=0, column=0, sticky='nsew', padx=3, pady=3)

        # Configure main frame grid layout
        # Column 0 will expand horizontally (weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Configure row weights to control vertical space distribution
        # weight=0: Fixed size rows (controls)
        # weight=1: Expandable row (log area)
        main_frame.rowconfigure(0, weight=0)  # Title - fixed size
        main_frame.rowconfigure(1, weight=0)  # Connection controls - fixed size  
        main_frame.rowconfigure(2, weight=0)  # Channel configuration - fixed size
        main_frame.rowconfigure(3, weight=0)  # File preferences - fixed size
        main_frame.rowconfigure(4, weight=0)  # Operation buttons - fixed size
        main_frame.rowconfigure(5, weight=1)  # Status/Log area - EXPANDABLE

        # Create title label at the top
        title_label = ttk.Label(main_frame, 
                               text="Keysight Oscilloscope Automation",
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, pady=(0, 3), sticky='ew')

        # Create all functional sections in order
        # Each section is placed in its designated row
        self.create_connection_frame(main_frame, row=1)
        self.create_channel_config_frame(main_frame, row=2)
        self.create_file_preferences_frame(main_frame, row=3)
        self.create_operations_frame(main_frame, row=4)
        self.create_status_frame(main_frame, row=5)  # This expands to fill remaining space

        # Store reference to main frame for potential future use
        self.main_frame = main_frame

    def configure_styles(self):
        """
        Configure custom ttk widget styles for consistent appearance.
        
        Defines custom styles for labels and buttons with specific fonts and colors.
        These styles can be applied to widgets using the style parameter.
        """
        # Title label style: Large bold font with dark blue color
        self.style.configure('Title.TLabel', 
                            font=('Times New Roman', 12, 'bold'), 
                            foreground='#1a365d',  # Dark blue text
                            background='#f5f5f5')  # Light gray background
        
        # Button styles with compact fonts for space efficiency
        self.style.configure('Success.TButton', font=('Arial', 8))  # For successful operations
        self.style.configure('Warning.TButton', font=('Arial', 8))  # For warning operations
        self.style.configure('Info.TButton', font=('Arial', 8))     # For informational operations
        self.style.configure('Primary.TButton', font=('Arial', 8))  # For primary actions

    def create_connection_frame(self, parent, row):
        """Create ULTRA-COMPACT connection controls"""
        conn_frame = ttk.LabelFrame(parent, text="Connection", padding="3")
        conn_frame.grid(row=row, column=0, sticky='ew', pady=(0, 2))
        conn_frame.columnconfigure(1, weight=1)  # Entry field expands

        # SINGLE ROW layout
        ttk.Label(conn_frame, text="VISA:", font=('Calibri', 8)).grid(row=0, column=0, sticky='w', padx=(0, 3))

        self.visa_address_var = tk.StringVar(value="USB0::0x0957::0x1780::MY65220169::INSTR")
        self.visa_entry = ttk.Entry(conn_frame, textvariable=self.visa_address_var, font=('Arial', 8))
        self.visa_entry.grid(row=0, column=1, sticky='ew', padx=(0, 3))

        self.connect_btn = ttk.Button(conn_frame, text="Connect", width=8, command=self.connect_oscilloscope, style='Success.TButton')
        self.connect_btn.grid(row=0, column=2, padx=1)

        self.disconnect_btn = ttk.Button(conn_frame, text="Disc", width=6, command=self.disconnect_oscilloscope, style='Warning.TButton', state='disabled')
        self.disconnect_btn.grid(row=0, column=3, padx=1)

        self.test_btn = ttk.Button(conn_frame, text="Test", width=5, command=self.test_connection, style='Info.TButton', state='disabled')
        self.test_btn.grid(row=0, column=4, padx=1)

        self.conn_status_var = tk.StringVar(value="Disconnected")
        self.conn_status_label = ttk.Label(conn_frame, textvariable=self.conn_status_var, font=('Arial', 8, 'bold'), foreground='#e53e3e')
        self.conn_status_label.grid(row=0, column=5, sticky='e', padx=(5, 0))

        # Ultra-compact instrument info
        self.info_text = tk.Text(conn_frame, height=1, font=('Courier', 7), state='disabled', bg='#f8f9fa', relief='flat', borderwidth=0)
        self.info_text.grid(row=1, column=0, columnspan=6, sticky='ew', pady=(2, 0))

    def create_channel_config_frame(self, parent, row):
        """
        Create channel configuration frame with multi-channel selection support.
        
        Allows users to select multiple channels for simultaneous data acquisition.
        Each channel can be independently enabled/disabled using checkboxes.
        Checked channels will be used for both acquisition and configuration.
        """
        config_frame = ttk.LabelFrame(parent, text="Channel Config", padding="3")
        config_frame.grid(row=row, column=0, sticky='ew', pady=(0, 2))

        # Single row with all controls
        col = 0

        # Multi-channel selection with checkboxes
        ttk.Label(config_frame, text="Select:", font=('Arial', 8, 'bold')).grid(row=0, column=col, sticky='w')
        col += 1
        
        # Create checkbox variables for each channel (1-4)
        # Each checkbox will control channel display on/off on the oscilloscope
        self.channel_enable_vars = {}
        for ch in [1, 2, 3, 4]:
            var = tk.BooleanVar(value=(ch == 1))  # Channel 1 enabled by default
            self.channel_enable_vars[ch] = var
            # Add callback to toggle channel display when checkbox changes
            var.trace_add('write', lambda *args, channel=ch: self.toggle_channel_display(channel))
            ttk.Checkbutton(config_frame, text=f"Ch{ch}", variable=var, 
                          style='TCheckbutton').grid(row=0, column=col, padx=2)
            col += 1
        
        # Add separator
        ttk.Separator(config_frame, orient='vertical').grid(row=0, column=col, sticky='ns', padx=8)
        col += 1

        ttk.Label(config_frame, text="V/div:", font=('Arial', 8)).grid(row=0, column=col, sticky='w')
        col += 1
        self.v_scale_var = tk.DoubleVar(value=1.0)
        ttk.Combobox(config_frame, textvariable=self.v_scale_var, values=[0.001, 0.01, 0.1, 1.0, 10.0], width=6, state='readonly', font=('Arial', 8)).grid(row=0, column=col, padx=(0, 8))
        col += 1

        ttk.Label(config_frame, text="Offset:", font=('Arial', 8)).grid(row=0, column=col, sticky='w')
        col += 1
        self.v_offset_var = tk.DoubleVar(value=0.0)
        ttk.Entry(config_frame, textvariable=self.v_offset_var, width=6, font=('Arial', 8)).grid(row=0, column=col, padx=(0, 8))
        col += 1

        ttk.Label(config_frame, text="Coup:", font=('Arial', 8)).grid(row=0, column=col, sticky='w')
        col += 1
        self.coupling_var = tk.StringVar(value="DC")
        ttk.Combobox(config_frame, textvariable=self.coupling_var, values=["AC", "DC"], width=4, state='readonly', font=('Arial', 8)).grid(row=0, column=col, padx=(0, 8))
        col += 1

        ttk.Label(config_frame, text="Probe:", font=('Arial', 8)).grid(row=0, column=col, sticky='w')
        col += 1
        self.probe_var = tk.DoubleVar(value=1.0)
        ttk.Combobox(config_frame, textvariable=self.probe_var, values=[1.0, 10.0, 100.0], width=5, state='readonly', font=('Arial', 8)).grid(row=0, column=col, padx=(0, 8))
        col += 1

        self.config_channel_btn = ttk.Button(config_frame, text="Configure", command=self.configure_channel, style='Primary.TButton', state='disabled')
        self.config_channel_btn.grid(row=0, column=col, sticky='ew')

        # Allow last column to expand
        config_frame.columnconfigure(col, weight=1)

    def create_file_preferences_frame(self, parent, row):
        """Create ULTRA-COMPACT file preferences"""
        pref_frame = ttk.LabelFrame(parent, text="File Preferences", padding="3")
        pref_frame.grid(row=row, column=0, sticky='ew', pady=(0, 2))

        # Configure columns for proper expansion
        pref_frame.columnconfigure(1, weight=1)
        pref_frame.columnconfigure(4, weight=1) 
        pref_frame.columnconfigure(7, weight=1)

        # Row 1: All three folder types
        ttk.Label(pref_frame, text="Data:", font=('Arial', 8, 'bold')).grid(row=0, column=0, sticky='w')
        self.data_path_var = tk.StringVar(value=str(Path.cwd() / "data"))
        ttk.Entry(pref_frame, textvariable=self.data_path_var, font=('Arial', 7)).grid(row=0, column=1, sticky='ew', padx=(2, 2))
        ttk.Button(pref_frame, text="...", command=lambda: self.browse_folder('data'), width=3).grid(row=0, column=2, padx=(0, 8))

        ttk.Label(pref_frame, text="Graphs:", font=('Arial', 8, 'bold')).grid(row=0, column=3, sticky='w')
        self.graph_path_var = tk.StringVar(value=str(Path.cwd() / "graphs"))
        ttk.Entry(pref_frame, textvariable=self.graph_path_var, font=('Arial', 7)).grid(row=0, column=4, sticky='ew', padx=(2, 2))
        ttk.Button(pref_frame, text="...", command=lambda: self.browse_folder('graphs'), width=3).grid(row=0, column=5, padx=(0, 8))

        ttk.Label(pref_frame, text="Screenshots:", font=('Arial', 8, 'bold')).grid(row=0, column=6, sticky='w')
        self.screenshot_path_var = tk.StringVar(value=str(Path.cwd() / "screenshots"))
        ttk.Entry(pref_frame, textvariable=self.screenshot_path_var, font=('Arial', 7)).grid(row=0, column=7, sticky='ew', padx=(2, 2))
        ttk.Button(pref_frame, text="...", command=lambda: self.browse_folder('screenshots'), width=3).grid(row=0, column=8)

        # Row 2: Graph title
        ttk.Label(pref_frame, text="Title:", font=('Arial', 8, 'bold')).grid(row=1, column=0, sticky='w', pady=(2, 0))
        self.graph_title_var = tk.StringVar(value="")
        ttk.Entry(pref_frame, textvariable=self.graph_title_var, font=('Arial', 8)).grid(row=1, column=1, columnspan=7, sticky='ew', padx=(2, 2), pady=(2, 0))
        ttk.Label(pref_frame, text="(auto)", font=('Arial', 7, 'italic'), foreground='gray').grid(row=1, column=8, pady=(2, 0))

    def create_operations_frame(self, parent, row):
        """Create ULTRA-COMPACT operations controls"""
        ops_frame = ttk.LabelFrame(parent, text="Operations", padding="3")
        ops_frame.grid(row=row, column=0, sticky='ew', pady=(0, 2))

        # Configure columns to distribute evenly
        for i in range(6):
            ops_frame.columnconfigure(i, weight=1)

        # Single row with all buttons - compact text
        self.screenshot_btn = ttk.Button(ops_frame, text="Screenshot", command=self.capture_screenshot, style='Info.TButton')
        self.screenshot_btn.grid(row=0, column=0, sticky='ew', padx=1)

        self.acquire_data_btn = ttk.Button(ops_frame, text="Acquire Data", command=self.acquire_data, style='Primary.TButton')
        self.acquire_data_btn.grid(row=0, column=1, sticky='ew', padx=1)

        self.export_csv_btn = ttk.Button(ops_frame, text="Export CSV", command=self.export_csv, style='Success.TButton')
        self.export_csv_btn.grid(row=0, column=2, sticky='ew', padx=1)

        self.generate_plot_btn = ttk.Button(ops_frame, text="Generate Plot", command=self.generate_plot, style='Success.TButton')
        self.generate_plot_btn.grid(row=0, column=3, sticky='ew', padx=1)

        self.full_automation_btn = ttk.Button(ops_frame, text="Full Auto", command=self.run_full_automation, style='Primary.TButton')
        self.full_automation_btn.grid(row=0, column=4, sticky='ew', padx=1)

        self.open_folder_btn = ttk.Button(ops_frame, text="Open Folder", command=self.open_output_folder, style='Info.TButton')
        self.open_folder_btn.grid(row=0, column=5, sticky='ew', padx=1)

        # Disable initially
        self.disable_operation_buttons()

    def create_status_frame(self, parent, row):
        """Create status frame that EXPANDS to fill ALL remaining space"""
        status_frame = ttk.LabelFrame(parent, text="Status & Activity Log", padding="3")
        status_frame.grid(row=row, column=0, sticky='nsew')  # NSEW fills all available space!

        # Configure status frame for proper expansion
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=0)  # Status line - fixed
        status_frame.rowconfigure(1, weight=0)  # Controls line - fixed  
        status_frame.rowconfigure(2, weight=1)  # LOG TEXT - EXPANDS TO FILL ALL REMAINING SPACE!

        # Status line - compact
        self.current_operation_var = tk.StringVar(value="Ready - Connect to oscilloscope")
        status_label = ttk.Label(status_frame, textvariable=self.current_operation_var, font=('Arial', 9, 'bold'), foreground='#1a365d')
        status_label.grid(row=0, column=0, sticky='ew', pady=(0, 2))

        # Controls line - compact
        controls_frame = ttk.Frame(status_frame)
        controls_frame.grid(row=1, column=0, sticky='ew', pady=(0, 2))
        controls_frame.columnconfigure(1, weight=1)  # Spacer expands

        ttk.Label(controls_frame, text="Activity Log:", font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky='w')
        ttk.Frame(controls_frame).grid(row=0, column=1, sticky='ew')  # Spacer
        ttk.Button(controls_frame, text="Clear", command=self.clear_log, width=6).grid(row=0, column=2, padx=(0, 2))
        ttk.Button(controls_frame, text="Save", command=self.save_log, width=6).grid(row=0, column=3)

        # THE MAGIC: LOG TEXT EXPANDS TO FILL ALL REMAINING SPACE!
        self.log_text = scrolledtext.ScrolledText(status_frame, 
                                                 font=('Consolas', 8),
                                                 bg='#1a1a1a', 
                                                 fg='#00ff00', 
                                                 insertbackground='white',
                                                 wrap=tk.WORD)
        self.log_text.grid(row=2, column=0, sticky='nsew')  # NSEW = fills ALL remaining space

    def browse_folder(self, folder_type):
        """
        Open a folder browser dialog to select a custom save location.
        
        Allows the user to choose a directory for saving specific types of files
        (data, graphs, or screenshots). Updates the corresponding path variable
        and creates the directory if it doesn't exist.
        
        Args:
            folder_type: Type of folder to browse for ('data', 'graphs', or 'screenshots')
        """
        try:
            # Map folder types to their corresponding StringVar attribute names
            path_var_mapping = {
                'data': 'data_path_var',
                'graphs': 'graph_path_var',
                'screenshots': 'screenshot_path_var'
            }

            # Get the attribute name for this folder type
            var_name = path_var_mapping.get(folder_type)
            if not var_name:
                raise ValueError(f"Unknown folder type: {folder_type}")

            # Get the current path value, use current working directory as fallback
            current_path = getattr(self, var_name).get()
            if not current_path or not os.path.exists(current_path):
                current_path = str(Path.cwd())

            self.log_message(f"Opening folder dialog for {folder_type}...")

            # Open the folder selection dialog
            # mustexist=False allows creating new directories
            folder_path = filedialog.askdirectory(
                initialdir=current_path,
                title=f"Select {folder_type.title()} Save Location",
                mustexist=False
            )

            # If user selected a folder (didn't cancel)
            if folder_path and folder_path.strip():
                # Update the StringVar with the new path
                getattr(self, var_name).set(folder_path)
                # Update the save_locations dictionary
                self.save_locations[folder_type] = folder_path
                self.log_message(f"Updated {folder_type}: {folder_path}", "SUCCESS")
                # Create the directory if it doesn't exist
                Path(folder_path).mkdir(parents=True, exist_ok=True)
                self.log_message(f"Directory verified: {folder_path}")
            else:
                self.log_message(f"Folder selection cancelled for {folder_type}")

        except Exception as e:
            self.log_message(f"Error selecting {folder_type} folder: {str(e)}", "ERROR")
            messagebox.showerror("Folder Selection Error", f"Error: {str(e)}")
    
    def log_message(self, message: str, level: str = "INFO"):
        """
        Add a timestamped message to the activity log with color coding.
        
        Inserts a message into the scrolled text widget with appropriate color
        based on the message level. Automatically scrolls to show the latest message.
        
        Args:
            message: The message text to log
            level: Message severity level ("INFO", "SUCCESS", "WARNING", or "ERROR")
        """
        # Create timestamp and format log entry
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"

        try:
            # Insert the message at the end of the log
            self.log_text.insert(tk.END, log_entry)
            # Auto-scroll to show the latest message
            self.log_text.see(tk.END)

            # Apply color coding based on message level
            # Tag names are used to apply formatting to specific text ranges
            if level == "ERROR":
                # Red color for errors
                self.log_text.tag_add("error", f"end-{len(log_entry)}c", "end-1c")
                self.log_text.tag_config("error", foreground="#ff6b6b")
            elif level == "SUCCESS":
                # Green color for successful operations
                self.log_text.tag_add("success", f"end-{len(log_entry)}c", "end-1c")
                self.log_text.tag_config("success", foreground="#51cf66")
            elif level == "WARNING":
                # Yellow color for warnings
                self.log_text.tag_add("warning", f"end-{len(log_entry)}c", "end-1c")
                self.log_text.tag_config("warning", foreground="#ffd43b")
            else:
                # Light blue color for informational messages
                self.log_text.tag_add("info", f"end-{len(log_entry)}c", "end-1c")
                self.log_text.tag_config("info", foreground="#74c0fc")
        except Exception as e:
            # Fallback to console if GUI logging fails
            print(f"Log error: {e}")

    def clear_log(self):
        """Clear log"""
        try:
            self.log_text.delete(1.0, tk.END)
            self.log_message("Log cleared")
        except Exception as e:
            print(f"Clear error: {e}")

    def save_log(self):
        """Save log to file"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_content = self.log_text.get(1.0, tk.END)

            filename = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialname=f"oscilloscope_log_{timestamp}.txt"
            )

            if filename:
                with open(filename, 'w') as f:
                    f.write(f"Oscilloscope Automation Log - {datetime.now()}\n{'='*50}\n\n{log_content}")
                self.log_message(f"Log saved: {filename}", "SUCCESS")
        except Exception as e:
            self.log_message(f"Save error: {e}", "ERROR")

    def update_status(self, status: str):
        """Update status"""
        try:
            self.current_operation_var.set(status)
            self.root.update_idletasks()
        except:
            pass

    def get_selected_channels(self):
        """
        Get list of channels that are currently enabled for acquisition.
        
        Returns:
            List of channel numbers (1-4) that have their checkboxes checked.
            Returns empty list if no channels are selected.
        """
        selected = []
        for ch, var in self.channel_enable_vars.items():
            if var.get():  # If checkbox is checked
                selected.append(ch)
        return sorted(selected)  # Return in ascending order

    def toggle_channel_display(self, channel):
        """
        Toggle channel display on/off on the oscilloscope.
        
        When a channel checkbox is checked/unchecked, this automatically
        turns the channel display on/off on the oscilloscope screen.
        This ensures screenshots only show the channels you want.
        
        Args:
            channel: Channel number (1-4) to toggle
        """
        # Only execute if oscilloscope is connected
        if not self.oscilloscope or not self.oscilloscope.is_connected:
            return
        
        try:
            # Get the current state of the checkbox
            is_enabled = self.channel_enable_vars[channel].get()
            
            # Send SCPI command to turn channel display on or off
            # :CHANnel<n>:DISPlay <state> where state is 1 (ON) or 0 (OFF)
            state = 1 if is_enabled else 0
            command = f":CHANnel{channel}:DISPlay {state}"
            
            self.oscilloscope._scpi_wrapper.write(command)
            
            # Log the action
            action = "enabled" if is_enabled else "disabled"
            self.log_message(f"Ch{channel} display {action}")
            
        except Exception as e:
            self.log_message(f"Error toggling Ch{channel} display: {e}", "ERROR")

    def disable_operation_buttons(self):
        """Disable operation buttons"""
        buttons = [self.screenshot_btn, self.acquire_data_btn, self.export_csv_btn,
                  self.generate_plot_btn, self.full_automation_btn, self.open_folder_btn,
                  self.config_channel_btn, self.test_btn]
        for btn in buttons:
            try:
                btn.configure(state='disabled')
            except:
                pass

    def enable_operation_buttons(self):
        """Enable operation buttons"""
        buttons = [self.screenshot_btn, self.acquire_data_btn, self.export_csv_btn,
                  self.generate_plot_btn, self.full_automation_btn, self.open_folder_btn,
                  self.config_channel_btn, self.test_btn]
        for btn in buttons:
            try:
                btn.configure(state='normal')
            except:
                pass

    def connect_oscilloscope(self):
        """
        Establish connection to the oscilloscope via VISA.
        
        Creates a new thread to perform the connection operation without blocking
        the GUI. Validates the VISA address, creates the oscilloscope instance,
        and initializes the data acquisition handler upon successful connection.
        
        The connection status is communicated back to the GUI thread via the
        status queue to ensure thread-safe UI updates.
        """
        def connect_thread():
            """Worker thread function for oscilloscope connection."""
            try:
                # Update GUI status
                self.update_status("Connecting...")
                self.log_message("Connecting to Keysight oscilloscope...")

                # Get and validate VISA address from the entry field
                visa_address = self.visa_address_var.get().strip()
                if not visa_address:
                    raise ValueError("VISA address empty")

                # Create oscilloscope instance with the specified VISA address
                self.oscilloscope = KeysightDSOX6004A(visa_address)

                # Attempt to establish connection
                if self.oscilloscope.connect():
                    # Create data acquisition handler for this oscilloscope
                    self.data_acquisition = OscilloscopeDataAcquisition(self.oscilloscope)
                    
                    # Retrieve instrument information for display
                    info = self.oscilloscope.get_instrument_info()
                    if info:
                        self.log_message(f"Connected: {info['manufacturer']} {info['model']}", "SUCCESS")
                        # Send success status with instrument info to GUI thread
                        self.status_queue.put(("connected", info))
                    else:
                        # Connected but no info available
                        self.status_queue.put(("connected", None))
                else:
                    raise Exception("Connection failed")
            except Exception as e:
                # Send error status to GUI thread
                self.status_queue.put(("error", f"Connection failed: {str(e)}"))

        # Start connection in a separate daemon thread
        # daemon=True ensures thread terminates when main program exits
        threading.Thread(target=connect_thread, daemon=True).start()

    def disconnect_oscilloscope(self):
        """Disconnect oscilloscope"""
        try:
            if self.oscilloscope:
                self.oscilloscope.disconnect()
                self.oscilloscope = None
                self.data_acquisition = None
                self.last_acquired_data = None

                self.conn_status_var.set("Disconnected")
                self.conn_status_label.configure(foreground='#e53e3e')
                self.connect_btn.configure(state='normal')
                self.disconnect_btn.configure(state='disabled')
                self.disable_operation_buttons()

                self.info_text.configure(state='normal')
                self.info_text.delete(1.0, tk.END)
                self.info_text.configure(state='disabled')

                self.update_status("Disconnected")
                self.log_message("Disconnected", "SUCCESS")
        except Exception as e:
            self.log_message(f"Disconnect error: {e}", "ERROR")

    def test_connection(self):
        """Test connection"""
        try:
            if self.oscilloscope and self.oscilloscope.is_connected:
                self.log_message("Connection test: PASSED", "SUCCESS")
                self.update_status("Test passed")
                messagebox.showinfo("Test", "Connection OK!")
            else:
                self.log_message("Connection test: FAILED", "ERROR")
                messagebox.showerror("Test", "Not connected")
        except Exception as e:
            self.log_message(f"Test error: {e}", "ERROR")

    def configure_channel(self):
        """
        Configure all selected channels with the specified settings.
        
        Applies the V/div, Offset, Coupling, and Probe settings to all
        channels that are currently checked in the GUI.
        """
        def config_thread():
            try:
                # Get selected channels
                selected_channels = self.get_selected_channels()
                
                if not selected_channels:
                    self.status_queue.put(("error", "No channels selected. Please check at least one channel."))
                    return
                
                # Get configuration parameters
                v_scale = self.v_scale_var.get()
                v_offset = self.v_offset_var.get()
                coupling = self.coupling_var.get()
                probe = self.probe_var.get()

                self.update_status(f"Configuring {len(selected_channels)} channel(s)...")
                self.log_message(f"Configuring channels {selected_channels}: {v_scale}V/div, {v_offset}V, {coupling}, {probe}x")

                # Configure each selected channel
                success_count = 0
                for channel in selected_channels:
                    self.log_message(f"Configuring Ch{channel}...")
                    success = self.oscilloscope.configure_channel(
                        channel=channel, vertical_scale=v_scale, vertical_offset=v_offset,
                        coupling=coupling, probe_attenuation=probe)

                    if success:
                        success_count += 1
                        self.log_message(f"Ch{channel} configured successfully", "SUCCESS")
                    else:
                        self.log_message(f"Ch{channel} configuration failed", "ERROR")

                # Send completion status
                if success_count == len(selected_channels):
                    self.status_queue.put(("channel_configured", f"All {success_count} channel(s) configured"))
                elif success_count > 0:
                    self.status_queue.put(("channel_configured", f"{success_count}/{len(selected_channels)} channel(s) configured"))
                else:
                    self.status_queue.put(("error", "All channel configurations failed"))
                    
            except Exception as e:
                self.status_queue.put(("error", f"Config error: {str(e)}"))

        if self.oscilloscope and self.oscilloscope.is_connected:
            threading.Thread(target=config_thread, daemon=True).start()
        else:
            messagebox.showerror("Error", "Not connected")

    def capture_screenshot(self):
        """Capture screenshot"""
        def screenshot_thread():
            try:
                self.update_status("Capturing screenshot...")
                self.log_message("Capturing screenshot...")

                screenshot_dir = Path(self.screenshot_path_var.get())
                screenshot_dir.mkdir(parents=True, exist_ok=True)

                filename = self.oscilloscope.capture_screenshot()
                if filename:
                    original_path = Path(filename)
                    if original_path.parent != screenshot_dir:
                        new_path = screenshot_dir / original_path.name
                        import shutil
                        shutil.move(str(original_path), str(new_path))
                        filename = str(new_path)

                    self.status_queue.put(("screenshot_success", filename))
                else:
                    self.status_queue.put(("error", "Screenshot failed"))
            except Exception as e:
                self.status_queue.put(("error", f"Screenshot error: {str(e)}"))

        if self.oscilloscope and self.oscilloscope.is_connected:
            threading.Thread(target=screenshot_thread, daemon=True).start()
        else:
            messagebox.showerror("Error", "Not connected")

    def acquire_data(self):
        """
        Acquire waveform data from selected channels.
        
        Acquires data from all channels that are enabled (checked) in the GUI.
        If multiple channels are selected, data is acquired sequentially from each.
        All acquired data is stored in self.last_acquired_data as a dictionary
        keyed by channel number.
        """
        def acquire_thread():
            try:
                # Get list of selected channels
                selected_channels = self.get_selected_channels()
                
                if not selected_channels:
                    self.status_queue.put(("error", "No channels selected. Please check at least one channel."))
                    return
                
                self.update_status(f"Acquiring data from {len(selected_channels)} channel(s)...")
                self.log_message(f"Acquiring data from channels: {selected_channels}")
                
                # Dictionary to store data from all channels
                all_channel_data = {}
                
                # Acquire data from each selected channel
                for channel in selected_channels:
                    self.log_message(f"Acquiring Ch{channel}...")
                    data = self.data_acquisition.acquire_waveform_data(channel)
                    
                    if data:
                        all_channel_data[channel] = data
                        self.log_message(f"Ch{channel}: {data['points_count']} points acquired", "SUCCESS")
                    else:
                        self.log_message(f"Ch{channel}: Acquisition failed", "ERROR")
                
                # Send all acquired data to GUI thread
                if all_channel_data:
                    self.status_queue.put(("data_acquired", all_channel_data))
                else:
                    self.status_queue.put(("error", "Data acquisition failed for all channels"))
                    
            except Exception as e:
                self.status_queue.put(("error", f"Acquire error: {str(e)}"))

        if self.data_acquisition:
            threading.Thread(target=acquire_thread, daemon=True).start()
        else:
            messagebox.showerror("Error", "Not connected")

    def export_csv(self):
        """
        Export acquired waveform data to CSV files.
        
        If multiple channels were acquired, creates a separate CSV file for each channel.
        Each file is named with the channel number and timestamp.
        """
        if not hasattr(self, 'last_acquired_data') or not self.last_acquired_data:
            messagebox.showwarning("Warning", "No data. Acquire first.")
            return

        def export_thread():
            try:
                self.update_status("Exporting CSV...")
                self.log_message("Exporting CSV...")
                
                exported_files = []
                
                # Check if data is a dictionary (multi-channel) or single channel data
                if isinstance(self.last_acquired_data, dict) and 'channel' not in self.last_acquired_data:
                    # Multi-channel data: dictionary keyed by channel number
                    for channel, data in self.last_acquired_data.items():
                        filename = self.data_acquisition.export_to_csv(
                            data, custom_path=self.data_path_var.get())
                        if filename:
                            exported_files.append(filename)
                            self.log_message(f"Ch{channel} CSV exported: {Path(filename).name}", "SUCCESS")
                else:
                    # Single channel data (backward compatibility)
                    filename = self.data_acquisition.export_to_csv(
                        self.last_acquired_data, custom_path=self.data_path_var.get())
                    if filename:
                        exported_files.append(filename)

                if exported_files:
                    self.status_queue.put(("csv_exported", exported_files))
                else:
                    self.status_queue.put(("error", "CSV export failed"))
            except Exception as e:
                self.status_queue.put(("error", f"CSV error: {str(e)}"))

        threading.Thread(target=export_thread, daemon=True).start()

    def generate_plot(self):
        """
        Generate plots for acquired waveform data.
        
        If multiple channels were acquired, creates a separate plot for each channel.
        Each plot is saved with the channel number and timestamp in the filename.
        """
        if not hasattr(self, 'last_acquired_data') or not self.last_acquired_data:
            messagebox.showwarning("Warning", "No data. Acquire first.")
            return

        def plot_thread():
            try:
                self.update_status("Generating plot...")
                self.log_message("Generating plot...")
                
                generated_plots = []
                custom_title = self.graph_title_var.get().strip() or None

                # Check if data is a dictionary (multi-channel) or single channel data
                if isinstance(self.last_acquired_data, dict) and 'channel' not in self.last_acquired_data:
                    # Multi-channel data: dictionary keyed by channel number
                    for channel, data in self.last_acquired_data.items():
                        # Create custom title for each channel if base title provided
                        if custom_title:
                            channel_title = f"{custom_title} - Channel {channel}"
                        else:
                            channel_title = None
                            
                        filename = self.data_acquisition.generate_waveform_plot(
                            data, custom_path=self.graph_path_var.get(), plot_title=channel_title)
                        if filename:
                            generated_plots.append(filename)
                            self.log_message(f"Ch{channel} plot generated: {Path(filename).name}", "SUCCESS")
                else:
                    # Single channel data (backward compatibility)
                    filename = self.data_acquisition.generate_waveform_plot(
                        self.last_acquired_data, custom_path=self.graph_path_var.get(), plot_title=custom_title)
                    if filename:
                        generated_plots.append(filename)

                if generated_plots:
                    self.status_queue.put(("plot_generated", generated_plots))
                else:
                    self.status_queue.put(("error", "Plot failed"))
            except Exception as e:
                self.status_queue.put(("error", f"Plot error: {str(e)}"))

        threading.Thread(target=plot_thread, daemon=True).start()

    def run_full_automation(self):
        """
        Execute the complete automation workflow in sequence for all selected channels.
        
        Performs all oscilloscope operations in a single automated sequence:
        1. Capture screenshot of oscilloscope display
        2. Acquire waveform data from all selected channels
        3. Export data to CSV files (one per channel)
        4. Generate plots with statistics (one per channel)
        
        All operations run in a background thread to keep the GUI responsive.
        Results are saved to user-specified directories with timestamped filenames.
        """
        def full_automation_thread():
            """Worker thread function for full automation sequence."""
            try:
                # Get user-configured parameters
                selected_channels = self.get_selected_channels()
                if not selected_channels:
                    self.status_queue.put(("error", "No channels selected. Please check at least one channel."))
                    return
                    
                custom_title = self.graph_title_var.get().strip() or None

                self.log_message(f"Starting full automation for channels: {selected_channels}...")

                # Step 1: Capture screenshot of oscilloscope display
                self.update_status("Step 1/4: Screenshot...")
                self.log_message("Step 1/4: Screenshot...")
                screenshot_file = self.oscilloscope.capture_screenshot()
                if screenshot_file:
                    # Move screenshot to user-specified directory if needed
                    screenshot_dir = Path(self.screenshot_path_var.get())
                    screenshot_dir.mkdir(parents=True, exist_ok=True)
                    original_path = Path(screenshot_file)
                    if original_path.parent != screenshot_dir:
                        new_path = screenshot_dir / original_path.name
                        import shutil
                        shutil.move(str(original_path), str(new_path))
                        screenshot_file = str(new_path)

                # Step 2: Acquire waveform data from all selected channels
                self.update_status(f"Step 2/4: Acquiring data from {len(selected_channels)} channel(s)...")
                self.log_message(f"Step 2/4: Acquiring data from {len(selected_channels)} channel(s)...")
                
                all_channel_data = {}
                for channel in selected_channels:
                    self.log_message(f"Acquiring Ch{channel}...")
                    data = self.data_acquisition.acquire_waveform_data(channel)
                    if data:
                        all_channel_data[channel] = data
                        self.log_message(f"Ch{channel}: {data['points_count']} points acquired", "SUCCESS")
                    else:
                        self.log_message(f"Ch{channel}: Acquisition failed", "ERROR")
                
                if not all_channel_data:
                    raise Exception("Data acquisition failed for all channels")

                # Step 3: Export acquired data to CSV format (one file per channel)
                self.update_status("Step 3/4: Exporting CSV...")
                self.log_message("Step 3/4: Exporting CSV...")
                csv_files = []
                for channel, data in all_channel_data.items():
                    csv_file = self.data_acquisition.export_to_csv(data, custom_path=self.data_path_var.get())
                    if csv_file:
                        csv_files.append(csv_file)
                        self.log_message(f"Ch{channel} CSV exported", "SUCCESS")
                    else:
                        self.log_message(f"Ch{channel} CSV export failed", "ERROR")

                # Step 4: Generate plots with statistics (one plot per channel)
                self.update_status("Step 4/4: Generating plots...")
                self.log_message("Step 4/4: Generating plots...")
                plot_files = []
                for channel, data in all_channel_data.items():
                    # Create custom title for each channel if base title provided
                    if custom_title:
                        channel_title = f"{custom_title} - Channel {channel}"
                    else:
                        channel_title = None
                        
                    plot_file = self.data_acquisition.generate_waveform_plot(
                        data, custom_path=self.graph_path_var.get(), plot_title=channel_title)
                    if plot_file:
                        plot_files.append(plot_file)
                        self.log_message(f"Ch{channel} plot generated", "SUCCESS")
                    else:
                        self.log_message(f"Ch{channel} plot generation failed", "ERROR")

                # Package all results and send to GUI thread
                results = {
                    'screenshot': screenshot_file, 
                    'csv': csv_files, 
                    'plot': plot_files, 
                    'data': all_channel_data,
                    'channels': selected_channels
                }
                self.status_queue.put(("full_automation_complete", results))

            except Exception as e:
                self.status_queue.put(("error", f"Automation error: {str(e)}"))

        if self.data_acquisition:
            threading.Thread(target=full_automation_thread, daemon=True).start()
        else:
            messagebox.showerror("Error", "Not connected")

    def open_output_folder(self):
        """Open output folders"""
        try:
            import subprocess
            import platform

            folders = [("Data", self.data_path_var.get()), ("Graphs", self.graph_path_var.get()), ("Screenshots", self.screenshot_path_var.get())]

            for name, path in folders:
                try:
                    path_obj = Path(path)
                    path_obj.mkdir(parents=True, exist_ok=True)

                    if platform.system() == "Windows":
                        subprocess.run(['explorer', str(path_obj)], check=True)
                    elif platform.system() == "Darwin":
                        subprocess.run(['open', str(path_obj)], check=True)
                    else:
                        subprocess.run(['xdg-open', str(path_obj)], check=True)

                    self.log_message(f"Opened {name}")
                except Exception as e:
                    self.log_message(f"Failed to open {name}: {e}", "ERROR")
        except Exception as e:
            self.log_message(f"Folder error: {e}", "ERROR")

    def display_instrument_info(self, info):
        """Display instrument info"""
        try:
            self.info_text.configure(state='normal')
            self.info_text.delete(1.0, tk.END)

            if info:
                info_str = f"Connected: {info.get('manufacturer', 'N/A')} {info.get('model', 'N/A')} | S/N: {info.get('serial_number', 'N/A')} | FW: {info.get('firmware_version', 'N/A')}"
            else:
                info_str = "Connected but no instrument info available"

            self.info_text.insert(1.0, info_str)
            self.info_text.configure(state='disabled')
        except Exception as e:
            print(f"Info display error: {e}")

    def check_status_updates(self):
        """Check for status updates"""
        try:
            while True:
                status_type, data = self.status_queue.get_nowait()

                if status_type == "connected":
                    self.conn_status_var.set("Connected")
                    self.conn_status_label.configure(foreground='#2d7d32')
                    self.connect_btn.configure(state='disabled')
                    self.disconnect_btn.configure(state='normal')
                    self.enable_operation_buttons()
                    self.update_status("Connected - Ready")
                    if data:
                        self.display_instrument_info(data)

                elif status_type == "error":
                    self.log_message(data, "ERROR")
                    self.update_status("Error occurred")

                elif status_type == "screenshot_success":
                    self.log_message(f"Screenshot saved: {Path(data).name}", "SUCCESS")
                    self.update_status("Screenshot saved")

                elif status_type == "data_acquired":
                    self.last_acquired_data = data
                    # Handle both single channel and multi-channel data
                    if isinstance(data, dict) and 'channel' not in data:
                        # Multi-channel data
                        total_points = sum(ch_data['points_count'] for ch_data in data.values())
                        self.log_message(f"Data acquired: {len(data)} channels, {total_points} total points", "SUCCESS")
                    else:
                        # Single channel data
                        self.log_message(f"Data acquired: {data['points_count']} points Ch{data['channel']}", "SUCCESS")
                    self.update_status("Data acquired")

                elif status_type == "csv_exported":
                    # Handle both single file and multiple files
                    if isinstance(data, list):
                        for filepath in data:
                            self.log_message(f"CSV exported: {Path(filepath).name}", "SUCCESS")
                        self.log_message(f"Total: {len(data)} CSV file(s) exported", "SUCCESS")
                    else:
                        self.log_message(f"CSV exported: {Path(data).name}", "SUCCESS")
                    self.update_status("CSV exported")

                elif status_type == "plot_generated":
                    # Handle both single plot and multiple plots
                    if isinstance(data, list):
                        for filepath in data:
                            self.log_message(f"Plot generated: {Path(filepath).name}", "SUCCESS")
                        self.log_message(f"Total: {len(data)} plot(s) generated", "SUCCESS")
                    else:
                        self.log_message(f"Plot generated: {Path(data).name}", "SUCCESS")
                    self.update_status("Plot generated")

                elif status_type == "channel_configured":
                    self.log_message(data, "SUCCESS")
                    self.update_status("Channel configured")

                elif status_type == "full_automation_complete":
                    self.last_acquired_data = data['data']
                    channels = data.get('channels', [])
                    self.log_message(f"Full automation completed for {len(channels)} channel(s)!", "SUCCESS")
                    self.log_message(f"Screenshot: {Path(data['screenshot']).name}", "SUCCESS")
                    
                    # Log CSV files
                    if isinstance(data['csv'], list):
                        for csv_file in data['csv']:
                            self.log_message(f"CSV: {Path(csv_file).name}", "SUCCESS")
                    else:
                        self.log_message(f"CSV: {Path(data['csv']).name}", "SUCCESS")
                    
                    # Log plot files
                    if isinstance(data['plot'], list):
                        for plot_file in data['plot']:
                            self.log_message(f"Plot: {Path(plot_file).name}", "SUCCESS")
                    else:
                        self.log_message(f"Plot: {Path(data['plot']).name}", "SUCCESS")
                    
                    self.update_status("Automation complete")

                    # Create summary message for dialog
                    csv_count = len(data['csv']) if isinstance(data['csv'], list) else 1
                    plot_count = len(data['plot']) if isinstance(data['plot'], list) else 1
                    
                    messagebox.showinfo("Complete!", 
                        f"Full automation complete!\n\n"
                        f"Channels: {', '.join(map(str, channels))}\n"
                        f"Screenshot: 1 file\n"
                        f"CSV files: {csv_count}\n"
                        f"Plots: {plot_count}")

        except queue.Empty:
            pass
        except Exception as e:
            print(f"Status update error: {e}")
        finally:
            self.root.after(100, self.check_status_updates)

    def run(self):
        """Start the application"""
        try:
            self.log_message("Enhanced Oscilloscope Automation - Professional Version", "SUCCESS")
            self.log_message("Features: Browse folders + Responsive layout + Auto-scaling", "SUCCESS")
            self.log_message("Ready to connect to oscilloscope")
            self.root.mainloop()
        except KeyboardInterrupt:
            self.log_message("Application interrupted")
        except Exception as e:
            self.log_message(f"Application error: {e}", "ERROR")
        finally:
            if hasattr(self, 'oscilloscope') and self.oscilloscope:
                try:
                    self.oscilloscope.disconnect()
                except:
                    pass

def main():
    """Main entry point"""
    print("Enhanced Keysight Oscilloscope Automation - Professional Version")
    print("Features: Browse folders + Responsive layout + True auto-scaling")
    print("Capabilities: Custom save locations + Custom titles + Responsive GUI")
    print("=" * 80)

    try:
        app = TrulyResponsiveAutomationGUI()
        app.run()
    except Exception as e:
        print(f"Error: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
