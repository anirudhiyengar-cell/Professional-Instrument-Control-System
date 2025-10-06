3#!/usr/bin/env python3
"""
Professional Instrument Control Automation Application - Enhanced UI

This application provides comprehensive automation for precision power supply control
and high-accuracy multimeter measurements with a professional terminal interface.

Application: instrument_automation_system_enhanced
Author: Professional Instrument Control Team
Version: 1.1.0
License: MIT
Dependencies: pyvisa, numpy, colorama, instrument_control modules

Features:
    - Professional terminal UI with colors and formatting
    - Multi-instrument discovery and coordination
    - Interactive configuration with enhanced validation
    - Precision measurement with statistical analysis
    - Real-time status displays and progress indicators
    - Comprehensive error handling and logging

Usage:
    python instrument_automation_system_enhanced.py
"""

import sys
import os
import time
import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, Dict, Any, List
from dataclasses import dataclass
from enum import Enum

# Enhanced UI imports
try:
    from colorama import init, Fore, Back, Style
    init(autoreset=True)  # Initialize colorama for Windows compatibility
    COLORAMA_AVAILABLE = True
except ImportError:
    # Fallback if colorama is not available
    COLORAMA_AVAILABLE = False
    class DummyColor:
        RED = GREEN = YELLOW = BLUE = MAGENTA = CYAN = WHITE = RESET = ""
        BRIGHT = DIM = ""
    Fore = Back = Style = DummyColor()

# Import instrument control modules
try:
    from instrument_control.keithley_power_supply import KeithleyPowerSupply, KeithleyPowerSupplyError
    from instrument_control.keithley_dmm import KeithleyDMM6500, KeithleyDMM6500Error
    from instrument_control.keysight_oscilloscope import KeysightDSOX6004A, KeysightDSOX6004AError
except ImportError as e:
    print(f"{Fore.RED}Error importing instrument control module: {e}")
    print("Please ensure all instrument control modules are in the 'instrument_control' package")
    sys.exit(1)

try:
    import pyvisa
except ImportError as e:
    print(f"{Fore.RED}PyVISA library is required. Install with: pip install pyvisa")
    sys.exit(1)


class UIColors:
    """Professional color scheme for terminal interface."""

    # Status colors
    SUCCESS = Fore.GREEN + Style.BRIGHT
    ERROR = Fore.RED + Style.BRIGHT
    WARNING = Fore.YELLOW + Style.BRIGHT
    INFO = Fore.CYAN

    # Section headers
    HEADER = Fore.BLUE + Style.BRIGHT
    SUBHEADER = Fore.MAGENTA + Style.BRIGHT

    # Data display
    VALUE = Fore.WHITE + Style.BRIGHT
    UNIT = Fore.CYAN + Style.DIM

    # Interactive elements
    PROMPT = Fore.YELLOW
    INPUT = Fore.WHITE

    # Separators and formatting
    SEPARATOR = Fore.BLUE + Style.DIM
    RESET = Style.RESET_ALL


class SystemState(Enum):
    """Enumeration of system operational states."""
    UNINITIALIZED = "uninitialized"
    DISCOVERING = "discovering"
    CONNECTING = "connecting"
    READY = "ready"
    RUNNING = "running"
    ERROR = "error"
    SHUTDOWN = "shutdown"


@dataclass
class TestConfiguration:
    """Data class for test configuration parameters."""
    channel: int
    voltage: float
    current_limit: float
    measurement_count: int = 1
    measurement_interval: float = 0.1
    enable_statistics: bool = False


@dataclass
class TestResults:
    """Data class for test execution results."""
    psu_voltage: float
    psu_current: float
    dmm_voltage: Optional[float]
    dmm_statistics: Optional[Dict[str, float]]
    measurement_accuracy: Optional[float]
    oscilloscope_screenshot_path: Optional[str]  
    timestamp: datetime


class ProgressIndicator:
    """Professional progress indicator for terminal."""

    def __init__(self, message: str, total_steps: int = 0):
        self.message = message
        self.total_steps = total_steps
        self.current_step = 0
        self._running = False
        self._thread = None
        self.spinner_chars = "|/-\\"
        self.spinner_index = 0

    def start(self):
        """Start the progress indicator."""
        self._running = True
        if self.total_steps > 0:
            self._show_progress()
        else:
            self._thread = threading.Thread(target=self._spin)
            self._thread.daemon = True
            self._thread.start()

    def update(self, step: int = None, message: str = None):
        """Update progress."""
        if step is not None:
            self.current_step = step
        if message is not None:
            self.message = message

        if self.total_steps > 0:
            self._show_progress()

    def stop(self, success: bool = True):
        """Stop the progress indicator."""
        self._running = False
        if self._thread:
            self._thread.join(timeout=0.1)

        # Clear the line and show final status
        print(f"\r{' ' * 80}\r", end='')
        status = f"{UIColors.SUCCESS}COMPLETE" if success else f"{UIColors.ERROR}FAILED"
        print(f"{status}{UIColors.RESET} {self.message}")

    def _spin(self):
        """Spinning animation for indeterminate progress."""
        while self._running:
            spinner = self.spinner_chars[self.spinner_index % len(self.spinner_chars)]
            print(f"\r{UIColors.INFO}{spinner} {self.message}...{UIColors.RESET}", end='')
            self.spinner_index += 1
            time.sleep(0.1)

    def _show_progress(self):
        """Show progress bar for determinate progress."""
        if self.total_steps == 0:
            return

        progress = min(self.current_step / self.total_steps, 1.0)
        bar_width = 40
        filled = int(bar_width * progress)
        bar = "█" * filled + "░" * (bar_width - filled)
        percentage = int(progress * 100)

        print(f"\r{UIColors.INFO}[{bar}] {percentage:3d}% {self.message}{UIColors.RESET}", end='')


class InstrumentAutomationSystemError(Exception):
    """Custom exception for automation system errors."""
    pass


class EnhancedInstrumentAutomationSystem:
    """
    Professional instrument automation system with enhanced UI.

    This class coordinates multiple instruments with a professional terminal
    interface featuring colors, progress indicators, and formatted displays.
    """

    def __init__(self, log_directory: str = "logs") -> None:
        """
        Initialize the automation system with enhanced UI.

        Args:
            log_directory: Directory path for log file storage
        """
        # Create log directory if it doesn't exist
        self._log_directory = Path(log_directory)
        self._log_directory.mkdir(exist_ok=True)

        # Initialize logging system
        self._setup_logging()
        self._logger = logging.getLogger(self.__class__.__name__)

        # Initialize system state
        self._system_state = SystemState.UNINITIALIZED

        # Initialize instrument instances
        self._power_supply: Optional[KeithleyPowerSupply] = None
        self._multimeter: Optional[KeithleyDMM6500] = None
        self._oscilloscope: Optional[KeysightDSOX6004A] = None

        # Store discovered instrument addresses
        self._instrument_addresses: Dict[str, str] = {}

        # Define safety limits
        self._max_safe_voltage = 30.0   # Maximum safe voltage (V)
        self._max_safe_current = 3.0    # Maximum safe current (A)

        # Test execution state
        self._current_test_config: Optional[TestConfiguration] = None
        self._test_results: List[TestResults] = []

        self._logger.info("Enhanced instrument automation system initialized")

    def run(self) -> None:
        """
        Execute the main automation sequence with enhanced UI.

        This method orchestrates the complete test sequence with professional
        visual feedback and status displays.
        """
        try:
            self._print_system_banner()
            self._print_system_status()

            # Phase 1: Instrument Discovery
            self._system_state = SystemState.DISCOVERING
            if not self._discover_instruments():
                self._print_error("Instrument discovery failed")
                return

            # Phase 2: Instrument Connection
            self._system_state = SystemState.CONNECTING
            if not self._connect_instruments():
                self._print_error("Instrument connection failed")
                return

            # Phase 3: System Ready - Enter Main Loop
            self._system_state = SystemState.READY
            self._print_success("System ready for operation")

            # Main automation loop
            while True:
                try:
                    self._clear_screen_section()
                    self._print_main_menu()

                    # Get test configuration from user
                    test_config = self._get_test_configuration()
                    if test_config is None:
                        self._logger.info("User cancelled test configuration")
                        break

                    # Execute test sequence
                    self._system_state = SystemState.RUNNING
                    success = self._execute_test_sequence(test_config)

                    if success:
                        self._logger.info("Test sequence completed successfully")
                        self._display_test_results()
                    else:
                        self._logger.warning("Test sequence completed with errors")

                    # Return to ready state
                    self._system_state = SystemState.READY

                    # Ask user if they want to continue
                    if not self._prompt_continue():
                        break

                except KeyboardInterrupt:
                    self._print_warning("Operation interrupted by user")
                    break

                except Exception as e:
                    self._logger.error(f"Unexpected error in main loop: {e}")
                    self._print_error(f"Unexpected error: {e}")
                    self._system_state = SystemState.ERROR
                    break

        except Exception as e:
            self._logger.error(f"Fatal error in automation system: {e}")
            self._print_error(f"Fatal system error: {e}")
            self._system_state = SystemState.ERROR

        finally:
            # Always perform safe shutdown
            self._system_state = SystemState.SHUTDOWN
            self._perform_safe_shutdown()

    def _setup_logging(self) -> None:
        """Configure comprehensive logging system."""
        # Create timestamp for log filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = self._log_directory / f"automation_enhanced_{timestamp}.log"

        # Configure logging format
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'

        # Setup logging with file output only (console has enhanced UI)
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            handlers=[
                logging.FileHandler(log_filename, mode='w', encoding='utf-8')
            ]
        )

        # Log system information
        self._print_info(f"Logging initialized: {log_filename}")

    def _print_system_banner(self) -> None:
        """Display professional system banner with enhanced formatting."""
        width = 88

        print(f"\n{UIColors.SEPARATOR}{'═' * width}{UIColors.RESET}")
        print(f"{UIColors.HEADER}{'PROFESSIONAL INSTRUMENT CONTROL AUTOMATION SYSTEM':^{width}}{UIColors.RESET}")
        print(f"{UIColors.SUBHEADER}{'Precision Power Supply Control & High-Accuracy Measurements':^{width}}{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'═' * width}{UIColors.RESET}")

        # Feature list with professional formatting
        features = [
            "Multi-instrument coordination and synchronization",
            "High-precision measurements with statistical analysis",
            "Real-time monitoring and safety interlocks",
            "Professional logging and data management"
        ]

        print(f"\n{UIColors.SUBHEADER}Key Features:{UIColors.RESET}")
        for feature in features:
            print(f"  {UIColors.SUCCESS}▶{UIColors.RESET} {feature}")

        print(f"\n{UIColors.INFO}Session Started:{UIColors.RESET} {UIColors.VALUE}{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * width}{UIColors.RESET}")

    def _print_system_status(self) -> None:
        """Display current system status with visual indicators."""
        status_map = {
            SystemState.UNINITIALIZED: ("INITIALIZING", UIColors.WARNING),
            SystemState.DISCOVERING: ("DISCOVERING", UIColors.INFO),
            SystemState.CONNECTING: ("CONNECTING", UIColors.INFO),
            SystemState.READY: ("READY", UIColors.SUCCESS),
            SystemState.RUNNING: ("RUNNING", UIColors.SUCCESS),
            SystemState.ERROR: ("ERROR", UIColors.ERROR),
            SystemState.SHUTDOWN: ("SHUTDOWN", UIColors.WARNING)
        }

        status_text, color = status_map.get(self._system_state, ("UNKNOWN", UIColors.ERROR))
        print(f"\n{UIColors.SUBHEADER}System Status:{UIColors.RESET} {color}● {status_text}{UIColors.RESET}")

    def _print_main_menu(self) -> None:
        """Display the main menu with enhanced formatting."""
        print(f"\n{UIColors.HEADER}MAIN CONTROL PANEL{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

        # Show connected instruments
        instruments = [
            ("Power Supply", self._power_supply, self._instrument_addresses.get('power_supply')),
            ("Multimeter", self._multimeter, self._instrument_addresses.get('multimeter')),
            ("Oscilloscope", self._oscilloscope, self._instrument_addresses.get('oscilloscope'))
        ]

        for name, instance, address in instruments:
            if instance and hasattr(instance, 'is_connected') and instance.is_connected:
                status = f"{UIColors.SUCCESS}● CONNECTED{UIColors.RESET}"
            elif address:
                status = f"{UIColors.WARNING}● CONFIGURED{UIColors.RESET}"
            else:
                status = f"{UIColors.ERROR}● NOT FOUND{UIColors.RESET}"

            print(f"  {name:<15} {status}")

    def _discover_instruments(self) -> bool:
        """
        Discover and identify connected instruments with enhanced UI.

        Returns:
            True if required instruments discovered, False otherwise
        """
        print(f"\n{UIColors.HEADER}PHASE 1: INSTRUMENT DISCOVERY{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 60}{UIColors.RESET}")

        progress = ProgressIndicator("Scanning VISA resources")
        progress.start()

        try:
            # Create VISA resource manager for discovery
            resource_manager = pyvisa.ResourceManager()
            available_resources = list(resource_manager.list_resources())

            progress.stop(success=True)

            if not available_resources:
                self._print_error("No VISA instruments detected")
                print(f"\n{UIColors.WARNING}Troubleshooting Suggestions:{UIColors.RESET}")
                suggestions = [
                    "Verify instrument power and USB connections",
                    "Check NI-VISA installation and drivers", 
                    "Try different USB ports or cables",
                    "Ensure instruments are not in use by other software"
                ]
                for i, suggestion in enumerate(suggestions, 1):
                    print(f"  {UIColors.INFO}{i}.{UIColors.RESET} {suggestion}")
                return False

            self._print_success(f"Discovered {len(available_resources)} VISA resources")

            # Display discovered resources in a formatted table
            print(f"\n{UIColors.SUBHEADER}Discovered Resources:{UIColors.RESET}")
            print(f"{UIColors.SEPARATOR}{'─' * 80}{UIColors.RESET}")

            for i, resource in enumerate(available_resources, 1):
                print(f"  {UIColors.VALUE}{i:2d}.{UIColors.RESET} {resource}")

            # Classify instruments by querying identification
            discovered_instruments = {'power_supply': None, 'multimeter': None, 'oscilloscope': None}

            print(f"\n{UIColors.SUBHEADER}Identifying Instruments:{UIColors.RESET}")
            print(f"{UIColors.SEPARATOR}{'─' * 80}{UIColors.RESET}")

            for resource in available_resources:
                try:
                    # Show progress for each instrument
                    id_progress = ProgressIndicator(f"Identifying {resource}")
                    id_progress.start()

                    # Open resource with extended timeout for identification
                    instrument = resource_manager.open_resource(resource, timeout=10000)

                    # Query instrument identification
                    identification = instrument.query("*IDN?").strip().upper()
                    instrument.close()

                    id_progress.stop(success=True)

                    # Display identification info
                    print(f"  {UIColors.INFO}●{UIColors.RESET} {resource}")
                    print(f"    {UIColors.VALUE}{identification[:70]}{UIColors.RESET}")

                    # Classify instrument based on identification
                    if 'KEITHLEY' in identification:
                        if any(model in identification for model in ['2230', '2231', '2280', '2260', '2268']):
                            discovered_instruments['power_supply'] = resource
                            print(f"    {UIColors.SUCCESS}→ Keithley Power Supply{UIColors.RESET}")
                        elif any(model in identification for model in ['DMM6500', 'DMM7510', '6500', '7510']):
                            discovered_instruments['multimeter'] = resource
                            print(f"    {UIColors.SUCCESS}→ Keithley Multimeter{UIColors.RESET}")
                    elif 'KEYSIGHT' in identification or 'AGILENT' in identification:
                        # Check for common Keysight/Agilent oscilloscope models
                        if any(model in identification.replace('-', '') for model in ['DSOX', 'MSOX']):
                            discovered_instruments['oscilloscope'] = resource
                            print(f"    {UIColors.SUCCESS}→ Keysight/Agilent Oscilloscope{UIColors.RESET}")
                    else:
                        print(f"    {UIColors.WARNING}→ Unknown instrument type{UIColors.RESET}")

                except Exception as e:
                    id_progress.stop(success=False)
                    print(f"    {UIColors.ERROR}→ Identification failed: {str(e)[:50]}{UIColors.RESET}")

            # Store discovered addresses
            self._instrument_addresses = discovered_instruments

            # Check for required instruments
            required_instruments = ['power_supply', 'multimeter']
            missing_instruments = [instr for instr in required_instruments 
                                 if discovered_instruments[instr] is None]

            print(f"\n{UIColors.SUBHEADER}Discovery Summary:{UIColors.RESET}")
            print(f"{UIColors.SEPARATOR}{'─' * 40}{UIColors.RESET}")

            for instr_type in ['power_supply', 'multimeter', 'oscilloscope']:
                required = instr_type in required_instruments
                found = discovered_instruments[instr_type] is not None

                if found:
                    status = f"{UIColors.SUCCESS}● FOUND{UIColors.RESET}"
                elif required:
                    status = f"{UIColors.ERROR}● MISSING (REQUIRED){UIColors.RESET}"
                else:
                    status = f"{UIColors.WARNING}● NOT FOUND (OPTIONAL){UIColors.RESET}"

                name = instr_type.replace('_', ' ').title()
                print(f"  {name:<15} {status}")

            if missing_instruments:
                self._print_warning(f"Missing required instruments: {', '.join(missing_instruments)}")

                # Offer manual configuration option
                response = self._get_user_input(
                    f"\n{UIColors.PROMPT}Enter instrument addresses manually? (y/N):{UIColors.RESET} ",
                    default="n"
                ).lower()

                if response in ['y', 'yes']:
                    return self._manual_instrument_configuration()
                else:
                    return False

            self._print_success("All required instruments discovered")
            return True

        except Exception as e:
            progress.stop(success=False)
            self._logger.error(f"Instrument discovery failed: {e}")
            self._print_error(f"Discovery failed: {e}")
            return False
        finally:
            try:
                resource_manager.close()
            except:
                pass

    def _manual_instrument_configuration(self) -> bool:
        """Allow manual entry of instrument VISA addresses with enhanced UI."""
        print(f"\n{UIColors.HEADER}MANUAL INSTRUMENT CONFIGURATION{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

        try:
            # Get power supply address
            psu_address = self._get_user_input(
                f"{UIColors.PROMPT}Power Supply VISA address:{UIColors.RESET} ",
                required=True
            )
            if not psu_address:
                self._print_error("Power supply address required")
                return False

            # Get multimeter address
            dmm_address = self._get_user_input(
                f"{UIColors.PROMPT}Multimeter VISA address:{UIColors.RESET} ",
                required=True
            )
            if not dmm_address:
                self._print_error("Multimeter address required")
                return False

            # Get optional oscilloscope address
            scope_address = self._get_user_input(
                f"{UIColors.PROMPT}Oscilloscope VISA address (optional):{UIColors.RESET} "
            )

            # Update instrument addresses
            self._instrument_addresses['power_supply'] = psu_address
            self._instrument_addresses['multimeter'] = dmm_address
            if scope_address:
                self._instrument_addresses['oscilloscope'] = scope_address

            self._print_success("Manual configuration completed")
            return True

        except KeyboardInterrupt:
            self._print_warning("Manual configuration cancelled")
            return False

    def _connect_instruments(self) -> bool:
        """
        Connect to discovered instruments with enhanced UI.

        Returns:
            True if all required instruments connected successfully
        """
        print(f"\n{UIColors.HEADER}PHASE 2: INSTRUMENT CONNECTION{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 60}{UIColors.RESET}")

        connection_success = True

        # Connect to power supply
        if self._instrument_addresses.get('power_supply'):
            progress = ProgressIndicator("Connecting to power supply")
            progress.start()

            try:
                self._power_supply = KeithleyPowerSupply(
                    self._instrument_addresses['power_supply'],
                    timeout_ms=15000
                )

                if self._power_supply.connect():
                    progress.stop(success=True)

                    info = self._power_supply.get_instrument_info()
                    if info:
                        print(f"  {UIColors.SUCCESS}Model:{UIColors.RESET} {info['manufacturer']} {info['model']}")
                        print(f"  {UIColors.INFO}Channels:{UIColors.RESET} {info['max_channels']}")
                        print(f"  {UIColors.INFO}Ratings:{UIColors.RESET} {info['max_voltage']}V / {info['max_current']}A")
                else:
                    progress.stop(success=False)
                    connection_success = False

            except Exception as e:
                progress.stop(success=False)
                self._print_error(f"Power supply connection failed: {e}")
                connection_success = False
        else:
            self._print_error("No power supply address available")
            connection_success = False

        # Connect to multimeter
        if self._instrument_addresses.get('multimeter'):
            progress = ProgressIndicator("Connecting to multimeter")
            progress.start()

            try:
                self._multimeter = KeithleyDMM6500(
                    self._instrument_addresses['multimeter'],
                    timeout_ms=30000  # Extended timeout for precision measurements
                )

                if self._multimeter.connect():
                    progress.stop(success=True)

                    info = self._multimeter.get_instrument_info()
                    if info:
                        print(f"  {UIColors.SUCCESS}Model:{UIColors.RESET} {info['manufacturer']} {info['model']}")
                        print(f"  {UIColors.INFO}Timeout:{UIColors.RESET} {info['timeout_ms']}ms")
                        print(f"  {UIColors.INFO}Max Range:{UIColors.RESET} {info['max_voltage_range']}V")

                        # Verify no initial errors
                        if info['current_errors'] != 'None':
                            self._print_warning(f"Initial errors: {info['current_errors']}")

                    # Test basic communication
                    test_progress = ProgressIndicator("Testing communication")
                    test_progress.start()

                    test_voltage = self._multimeter.measure_dc_voltage_fast()
                    if test_voltage is not None:
                        test_progress.stop(success=True)
                        print(f"  {UIColors.SUCCESS}Communication Test:{UIColors.RESET} {test_voltage:.6f}V")
                    else:
                        test_progress.stop(success=False)
                        self._print_warning("Communication test failed")

                else:
                    progress.stop(success=False)
                    connection_success = False

            except Exception as e:
                progress.stop(success=False)
                self._print_error(f"Multimeter connection failed: {e}")
                connection_success = False
        else:
            self._print_error("No multimeter address available")
            connection_success = False

        # Connect to oscilloscope (optional)
        if self._instrument_addresses.get('oscilloscope'):
            progress = ProgressIndicator("Connecting to oscilloscope")
            progress.start()

            try:
                self._oscilloscope = KeysightDSOX6004A(
                    self._instrument_addresses['oscilloscope'],
                    timeout_ms=15000
                )

                if self._oscilloscope.connect():
                    progress.stop(success=True)

                    info = self._oscilloscope.get_instrument_info()
                    if info:
                        print(f"  {UIColors.SUCCESS}Model:{UIColors.RESET} {info['manufacturer']} {info['model']}")
                        print(f"  {UIColors.INFO}Bandwidth:{UIColors.RESET} {info['bandwidth_hz']/1e9:.1f} GHz")
                        print(f"  {UIColors.INFO}Channels:{UIColors.RESET} {info['max_channels']}")
                else:
                    progress.stop(success=False)
                    # Oscilloscope is optional, don't fail overall connection

            except Exception as e:
                progress.stop(success=False)
                self._print_warning(f"Oscilloscope connection failed: {e}")
                # Oscilloscope is optional, continue without it

        # Display final connection status
        print(f"\n{UIColors.SUBHEADER}Connection Summary:{UIColors.RESET}")
        if connection_success:
            self._print_success("All required instruments connected successfully")
        else:
            self._print_error("Failed to connect to required instruments")

        return connection_success

    def _get_test_configuration(self) -> Optional[TestConfiguration]:
        """
        Get test configuration from user with enhanced UI validation.

        Returns:
            TestConfiguration object or None if cancelled
        """
        print(f"\n{UIColors.HEADER}PHASE 3: TEST CONFIGURATION{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

        if not self._power_supply:
            self._print_error("Power supply not available")
            return None

        try:
            # Display power supply capabilities in a professional format
            info = self._power_supply.get_instrument_info()
            if info:
                print(f"\n{UIColors.SUBHEADER}Power Supply Specifications:{UIColors.RESET}")
                print(f"  {UIColors.INFO}Model:{UIColors.RESET} {info['manufacturer']} {info['model']}")
                print(f"  {UIColors.INFO}Channels:{UIColors.RESET} 1 to {info['max_channels']}")
                print(f"  {UIColors.INFO}Max Voltage:{UIColors.RESET} {info['max_voltage']}V")
                print(f"  {UIColors.INFO}Max Current:{UIColors.RESET} {info['max_current']}A")
                print(f"  {UIColors.WARNING}Safety Limits:{UIColors.RESET} {self._max_safe_voltage}V / {self._max_safe_current}A")

            # Get channel selection with enhanced validation
            while True:
                try:
                    channel_input = self._get_user_input(
                        f"\n{UIColors.PROMPT}Select PSU channel (1-3):{UIColors.RESET} ",
                        required=True
                    )
                    if not channel_input:
                        return None

                    channel = int(channel_input)
                    max_channels = info['max_channels'] if info else 3

                    if 1 <= channel <= max_channels:
                        self._print_success(f"Channel {channel} selected")
                        break
                    else:
                        self._print_error(f"Channel must be 1-{max_channels}")

                except ValueError:
                    self._print_error("Please enter a valid channel number")
                except KeyboardInterrupt:
                    return None

            # Get voltage setting with visual feedback
            while True:
                try:
                    voltage_input = self._get_user_input(
                        f"{UIColors.PROMPT}Enter voltage (0.0-{self._max_safe_voltage:.1f}V):{UIColors.RESET} ",
                        required=True
                    )
                    if not voltage_input:
                        return None

                    voltage = float(voltage_input)

                    if 0.0 <= voltage <= self._max_safe_voltage:
                        self._print_success(f"Voltage set to {voltage:.3f}V")
                        break
                    else:
                        self._print_error(f"Voltage must be 0.0-{self._max_safe_voltage:.1f}V")

                except ValueError:
                    self._print_error("Please enter a valid voltage")
                except KeyboardInterrupt:
                    return None

            # Get current limit with validation
            while True:
                try:
                    current_input = self._get_user_input(
                        f"{UIColors.PROMPT}Enter current limit (0.01-{self._max_safe_current:.1f}A):{UIColors.RESET} ",
                        required=True
                    )
                    if not current_input:
                        return None

                    current_limit = float(current_input)

                    if 0.01 <= current_limit <= self._max_safe_current:
                        self._print_success(f"Current limit set to {current_limit:.3f}A")
                        break
                    else:
                        self._print_error(f"Current limit must be 0.01-{self._max_safe_current:.1f}A")

                except ValueError:
                    self._print_error("Please enter a valid current limit")
                except KeyboardInterrupt:
                    return None

            # Get measurement options with enhanced interface
            enable_statistics = False
            measurement_count = 1

            if self._multimeter:
                stats_response = 'y'
                #     #f"\n{UIColors.PROMPT}Enable measurement statistics? (y/N):{UIColors.RESET} ",
                #     default="y"
                # ).lower()
                #enable_statistics = stats_response in ['y', 'yes','n','no']

                if enable_statistics:
                    while True:
                        try:
                            count_input = 10
                                # f"{UIColors.PROMPT}Number of measurements (2-20) [10]:{UIColors.RESET} ",
                                # default="10"
                            # )

                            measurement_count = int(count_input)
                            if 2 <= measurement_count <= 20:
                                self._print_success(f"Will perform {measurement_count} measurements")
                                break
                            else:
                                self._print_error("Measurement count must be 2-20")
                        except ValueError:
                            self._print_error("Please enter a valid number")

            # Display configuration summary in a professional table format
            print(f"\n{UIColors.SUBHEADER}Configuration Summary:{UIColors.RESET}")
            print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")
            print(f"  {UIColors.INFO}Channel:{UIColors.RESET}       {UIColors.VALUE}{channel}{UIColors.RESET}")
            print(f"  {UIColors.INFO}Voltage:{UIColors.RESET}       {UIColors.VALUE}{voltage:.3f}{UIColors.UNIT}V{UIColors.RESET}")
            print(f"  {UIColors.INFO}Current Limit:{UIColors.RESET} {UIColors.VALUE}{current_limit:.3f}{UIColors.UNIT}A{UIColors.RESET}")
            if enable_statistics:
                print(f"  {UIColors.INFO}Measurements:{UIColors.RESET}  {UIColors.VALUE}{measurement_count}{UIColors.RESET} {UIColors.UNIT}(with statistics){UIColors.RESET}")
            print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

            # Confirm configuration
            confirm = self._get_user_input(
                f"\n{UIColors.PROMPT}Proceed with this configuration? (Y/n):{UIColors.RESET} ",
                default="y"
            ).lower()

            if confirm in ['', 'y', 'yes','n','no']:
                self._print_success("Configuration accepted")
                return TestConfiguration(
                    channel=channel,
                    voltage=voltage,
                    current_limit=current_limit,
                    measurement_count=measurement_count,
                    enable_statistics=enable_statistics
                )
            else:
                self._print_warning("Configuration cancelled")
                return None

        except KeyboardInterrupt:
            self._print_warning("Configuration cancelled by user")
            return None
        except Exception as e:
            self._logger.error(f"Error during test configuration: {e}")
            self._print_error(f"Configuration error: {e}")
            return None

    def _execute_test_sequence(self, config: TestConfiguration) -> bool:
        """
        Execute the complete test sequence with enhanced visual feedback.

        Args:
            config: Test configuration parameters

        Returns:
            True if test sequence completed successfully
        """
        print(f"\n{UIColors.HEADER}PHASE 4: TEST EXECUTION{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

        self._current_test_config = config

        try:
            # Step 1: Configure Power Supply
            step_progress = ProgressIndicator("Configuring power supply")
            step_progress.start()

            if not self._configure_power_supply(config):
                step_progress.stop(success=False)
                return False

            step_progress.stop(success=True)
            print(f"  {UIColors.SUCCESS}Configuration:{UIColors.RESET} CH{config.channel} = {config.voltage:.3f}V, {config.current_limit:.3f}A limit")

            # Step 2: Measure Resistance before enabling output
            if self._multimeter:
                step_progress = ProgressIndicator("Measuring resistance")
                step_progress.start()

                resistance = self._multimeter.measure_resistance_2w()
                if resistance is not None:
                    step_progress.stop(success=True)
                    print(f"  {UIColors.SUCCESS}Resistance Measurement:{UIColors.RESET} {UIColors.VALUE}{resistance:.3f}{UIColors.UNIT} Ω{UIColors.RESET}")
                else:
                    step_progress.stop(success=False)
                    self._print_warning("Could not measure resistance")
                    return False
            # # Step 2: Measure Capacitance before enabling output
            # if self._multimeter:
            #     step_progress = ProgressIndicator("Measuring capacitance")
            #     step_progress.start()

            #     capacitance = self._multimeter.measure_capacitance()
            #     if capacitance is not None:
            #         step_progress.stop(success=True)
            #         print(f"  {UIColors.SUCCESS}Capacitance Measurement:{UIColors.RESET} {UIColors.VALUE}{capacitance:.3f}{UIColors.UNIT}F{UIColors.RESET}")
            #     else:
            #         step_progress.stop(success=False)
            #         self._print_warning("Could not measure Capatance")
            #         return False
            

            # Step 3: Enable Output
            step_progress = ProgressIndicator("Enabling power supply output")
            step_progress.start()

            if not self._power_supply.enable_channel_output(config.channel):
                step_progress.stop(success=False)
                self._print_error("Failed to enable power supply output")
                return False

            step_progress.stop(success=True)
            print(f"  {UIColors.SUCCESS}Output Status:{UIColors.RESET} ENABLED")

            # Step 3: Verify Power Supply Output
            step_progress = ProgressIndicator("Verifying power supply output")
            step_progress.start()

            psu_measurements = self._power_supply.measure_channel_output(config.channel)
            if psu_measurements:
                psu_voltage, psu_current = psu_measurements
                step_progress.stop(success=True)
                print(f"  {UIColors.SUCCESS}PSU Measurements:{UIColors.RESET} {UIColors.VALUE}{psu_voltage:.6f}{UIColors.UNIT}V{UIColors.RESET}, {UIColors.VALUE}{psu_current:.6f}{UIColors.UNIT}A{UIColors.RESET}")
            else:
                step_progress.stop(success=False)
                self._print_error("Failed to measure power supply output")
                return False

            # Step 4: Multimeter Measurements
            dmm_voltage = None
            dmm_statistics = None

            if self._multimeter:
                if config.enable_statistics:
                    # Perform statistical measurements with progress
                    step_progress = ProgressIndicator("Performing statistical measurements", config.measurement_count)
                    step_progress.start()

                    dmm_statistics = self._multimeter.perform_measurement_statistics(
                        measurement_count=config.measurement_count,
                        measurement_interval=0.1
                    )

                    step_progress.stop(success=dmm_statistics is not None)

                    if dmm_statistics:
                        dmm_voltage = dmm_statistics['mean']
                        print(f"\n  {UIColors.SUBHEADER}Statistical Analysis (n={dmm_statistics['count']}):{UIColors.RESET}")
                        print(f"    {UIColors.INFO}Mean:{UIColors.RESET}     {UIColors.VALUE}{dmm_statistics['mean']:.9f}{UIColors.UNIT}V{UIColors.RESET}")
                        print(f"    {UIColors.INFO}Std Dev:{UIColors.RESET}  {UIColors.VALUE}{dmm_statistics['standard_deviation']:.9f}{UIColors.UNIT}V{UIColors.RESET}")
                        print(f"    {UIColors.INFO}Range:{UIColors.RESET}    {UIColors.VALUE}{dmm_statistics['range']:.9f}{UIColors.UNIT}V{UIColors.RESET}")
                        print(f"    {UIColors.INFO}CV:{UIColors.RESET}       {UIColors.VALUE}{dmm_statistics['coefficient_of_variation_percent']:.3f}{UIColors.UNIT}%{UIColors.RESET}")
                    else:
                        self._print_error("Statistical measurements failed")
                else:
                    # Perform single high-precision measurement
                    step_progress = ProgressIndicator("Performing precision measurement")
                    step_progress.start()

                    dmm_voltage = self._multimeter.measure_dc_voltage(
                        measurement_range=None,  # Auto-range
                        resolution=1e-6,         # 1µV resolution
                        nplc=1.0,               # 1 power line cycle
                        auto_zero=None          # Do not change auto-zero to avoid -113 on some models
                    )

                    step_progress.stop(success=dmm_voltage is not None)

                    if dmm_voltage is not None:
                        print(f"  {UIColors.SUCCESS}DMM Measurement:{UIColors.RESET} {UIColors.VALUE}{dmm_voltage:.9f}{UIColors.UNIT}V{UIColors.RESET}")
                    else:
                        self._print_error("DMM measurement failed")

            # Step 5: Analysis and Comparison
            if psu_measurements and dmm_voltage is not None:
                print(f"\n{UIColors.SUBHEADER}Measurement Analysis:{UIColors.RESET}")

                voltage_difference = abs(dmm_voltage - psu_voltage)
                measurement_accuracy = (voltage_difference / psu_voltage * 100) if psu_voltage > 0 else 0

                print(f"  {UIColors.INFO}PSU Reading:{UIColors.RESET}  {UIColors.VALUE}{psu_voltage:.6f}{UIColors.UNIT}V{UIColors.RESET}")
                print(f"  {UIColors.INFO}DMM Reading:{UIColors.RESET}  {UIColors.VALUE}{dmm_voltage:.9f}{UIColors.UNIT}V{UIColors.RESET}")
                print(f"  {UIColors.INFO}Difference:{UIColors.RESET}   {UIColors.VALUE}{voltage_difference*1000:.3f}{UIColors.UNIT}mV{UIColors.RESET} ({UIColors.VALUE}{measurement_accuracy:.3f}{UIColors.UNIT}%{UIColors.RESET})")

                # Assess measurement quality with color coding
                if voltage_difference < 0.001:  # 1mV
                    print(f"  {UIColors.SUCCESS}Assessment: EXCELLENT AGREEMENT{UIColors.RESET}")
                elif voltage_difference < 0.005:  # 5mV
                    print(f"  {UIColors.SUCCESS}Assessment: GOOD AGREEMENT{UIColors.RESET}")
                else:
                    print(f"  {UIColors.WARNING}Assessment: SIGNIFICANT DIFFERENCE{UIColors.RESET}")

            # Step 6: Capture Oscilloscope Screenshot
            screenshot_path = None
            if self._oscilloscope and self._oscilloscope.is_connected:
                # Get oscilloscope info for channel validation
                scope_info = self._oscilloscope.get_instrument_info()
                max_channels = scope_info['max_channels'] if scope_info else 4
                
                # Get channel selection with enhanced validation
                selected_channel = None
                while True:
                    try:
                        channel_input = self._get_user_input(
                            f"\n{UIColors.PROMPT}Select oscilloscope channel to configure (1-{max_channels}):{UIColors.RESET} ",
                            required=True
                        )
                        if not channel_input:
                            self._print_warning("Channel selection cancelled")
                            break

                        channel = int(channel_input)

                        if 1 <= channel <= max_channels:
                            selected_channel = channel
                            self._print_success(f"Channel {channel} selected for configuration")
                            break
                        else:
                            self._print_error(f"Channel must be 1-{max_channels}")

                    except ValueError:
                        self._print_error("Please enter a valid channel number")
                    except KeyboardInterrupt:
                        self._print_warning("Channel selection cancelled")
                        break

                if selected_channel:
                    step_progress = ProgressIndicator("Capturing oscilloscope screenshot")
                    step_progress.start()
                    
                    try:
                        # Configure the selected oscilloscope channel for better visibility
                        self._oscilloscope.configure_channel(selected_channel, 2.0, 0.0, "DC", 1.0)  # 2V/div for selected channel
                        time.sleep(0.5)  # Allow settings to stabilize
                        
                        # Capture screenshot with timestamp
                        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                        screenshot_filename = f"test_measurement_ch{selected_channel}_{timestamp}.png"
                        screenshot_path = self._oscilloscope.capture_screenshot(screenshot_filename, "PNG", True)
                        
                        if screenshot_path:
                            step_progress.stop(success=True)
                            print(f"  {UIColors.SUCCESS}Screenshot saved:{UIColors.RESET} {screenshot_path}")
                        else:
                            step_progress.stop(success=False)
                            self._print_warning("Failed to capture oscilloscope screenshot")
                            
                    except Exception as e:
                        step_progress.stop(success=False)
                        self._logger.error(f"Oscilloscope screenshot failed: {e}")
                        self._print_warning(f"Screenshot capture error: {e}")
                else:
                    self._print_info("No channel selected - skipping oscilloscope screenshot")
            else:
                self._print_info("Oscilloscope not available - skipping screenshot capture")

            # Step 7: Check for Errors
            step_progress = ProgressIndicator("Checking instrument errors")
            step_progress.start()

            error_count = 0
            if self._multimeter:
                dmm_errors = self._multimeter.check_instrument_errors()
                if dmm_errors:
                    error_count += len(dmm_errors)
                    step_progress.stop(success=False)
                    print(f"\n  {UIColors.WARNING}DMM Errors Detected ({len(dmm_errors)}):{UIColors.RESET}")
                    for error in dmm_errors:
                        print(f"    {UIColors.ERROR}●{UIColors.RESET} {error}")
                else:
                    step_progress.stop(success=True)

            if error_count == 0:
                print(f"  {UIColors.SUCCESS}No instrument errors detected{UIColors.RESET}")

            # Store test results
            test_result = TestResults(
                psu_voltage=psu_voltage,
                psu_current=psu_current,
                dmm_voltage=dmm_voltage,
                dmm_statistics=dmm_statistics,
                measurement_accuracy=measurement_accuracy,
                oscilloscope_screenshot_path=screenshot_path,
                timestamp=datetime.now()
            )
            self._test_results.append(test_result)

            self._print_success("Test sequence completed successfully")
            return True

        except Exception as e:
            self._logger.error(f"Test sequence failed: {e}")
            self._print_error(f"Test sequence failed: {e}")
            return False

    def _configure_power_supply(self, config: TestConfiguration) -> bool:
        """Configure power supply with specified parameters."""
        if not self._power_supply:
            return False

        try:
            return self._power_supply.configure_channel(
                channel=config.channel,
                voltage=config.voltage,
                current_limit=config.current_limit,
                ovp_level=config.voltage + 1.0,  # Set OVP 1V above target
                enable_output=False  # Will enable separately
            )
        except Exception as e:
            self._logger.error(f"Power supply configuration failed: {e}")
            return False

    def _display_test_results(self) -> None:
        """Display comprehensive test results summary with enhanced formatting."""
        if not self._test_results:
            return

        print(f"\n{UIColors.HEADER}TEST RESULTS SUMMARY{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'═' * 60}{UIColors.RESET}")

        latest_result = self._test_results[-1]

        # Create a professional results table
        print(f"\n{UIColors.SUBHEADER}Measurement Results:{UIColors.RESET}")
        print(f"  {UIColors.INFO}Timestamp:{UIColors.RESET}      {UIColors.VALUE}{latest_result.timestamp.strftime('%Y-%m-%d %H:%M:%S')}{UIColors.RESET}")
        print(f"  {UIColors.INFO}Power Supply:{UIColors.RESET}   {UIColors.VALUE}{latest_result.psu_voltage:.6f}{UIColors.UNIT}V{UIColors.RESET}, {UIColors.VALUE}{latest_result.psu_current:.6f}{UIColors.UNIT}A{UIColors.RESET}")

        if latest_result.dmm_voltage is not None:
            print(f"  {UIColors.INFO}Multimeter:{UIColors.RESET}     {UIColors.VALUE}{latest_result.dmm_voltage:.9f}{UIColors.UNIT}V{UIColors.RESET}")

            if latest_result.measurement_accuracy is not None:
                if latest_result.measurement_accuracy < 0.1:
                    accuracy_color = UIColors.SUCCESS
                elif latest_result.measurement_accuracy < 0.5:
                    accuracy_color = UIColors.WARNING
                else:
                    accuracy_color = UIColors.ERROR

                print(f"  {UIColors.INFO}Accuracy:{UIColors.RESET}       {accuracy_color}{latest_result.measurement_accuracy:.3f}% difference{UIColors.RESET}")

        if latest_result.dmm_statistics:
            stats = latest_result.dmm_statistics
            print(f"  {UIColors.INFO}Statistics:{UIColors.RESET}     σ={UIColors.VALUE}{stats['standard_deviation']:.9f}{UIColors.UNIT}V{UIColors.RESET}, CV={UIColors.VALUE}{stats['coefficient_of_variation_percent']:.3f}{UIColors.UNIT}%{UIColors.RESET}")

        print(f"{UIColors.SEPARATOR}{'─' * 60}{UIColors.RESET}")

    def _prompt_continue(self) -> bool:
        """Prompt user whether to continue with another test."""
        try:
            response = self._get_user_input(
                f"\n{UIColors.PROMPT}Perform another test? (Y/n):{UIColors.RESET} ",
                default="y"
            ).lower()
            return response in ['', 'y', 'yes']
        except KeyboardInterrupt:
            return False

    def _capture_oscilloscope_screenshot(self) -> Optional[str]:
        """Capture a screenshot from the oscilloscope with UI feedback."""
        if not self._oscilloscope or not self._oscilloscope.is_connected:
            self._print_warning("Oscilloscope not connected, skipping screenshot.")
            return None

        progress = ProgressIndicator("Capturing oscilloscope screenshot")
        progress.start()

        try:
            # Generate a unique filename for the screenshot
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"test_screenshot_{timestamp}.png"

            screenshot_path = self._oscilloscope.capture_screenshot(filename=filename)

            if screenshot_path:
                progress.stop(success=True)
                self._print_success(f"Screenshot saved to: {screenshot_path}")
                return screenshot_path
            else:
                progress.stop(success=False)
                self._print_error("Failed to capture screenshot.")
                return None

        except Exception as e:
            progress.stop(success=False)
            self._logger.error(f"Oscilloscope screenshot failed: {e}")
            self._print_error(f"Screenshot failed: {e}")
            return None

    def _perform_safe_shutdown(self) -> None:
        """Perform safe shutdown of all instruments with enhanced feedback."""
        print(f"\n{UIColors.HEADER}SAFE SHUTDOWN SEQUENCE{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'═' * 50}{UIColors.RESET}")

        try:
            # Disable all power supply outputs
            if self._power_supply and self._power_supply.is_connected:
                shutdown_progress = ProgressIndicator("Disabling all power supply outputs")
                shutdown_progress.start()

                if self._power_supply.disable_all_outputs():
                    shutdown_progress.stop(success=True)
                else:
                    shutdown_progress.stop(success=False)
                    self._print_warning("Some outputs may still be enabled")

                disconnect_progress = ProgressIndicator("Disconnecting power supply")
                disconnect_progress.start()
                self._power_supply.disconnect()
                disconnect_progress.stop(success=True)

            # Disconnect multimeter
            if self._multimeter and self._multimeter.is_connected:
                disconnect_progress = ProgressIndicator("Disconnecting multimeter")
                disconnect_progress.start()
                self._multimeter.disconnect()
                disconnect_progress.stop(success=True)

            # Disconnect oscilloscope
            if self._oscilloscope and self._oscilloscope.is_connected:
                disconnect_progress = ProgressIndicator("Disconnecting oscilloscope")
                disconnect_progress.start()
                self._oscilloscope.disconnect()
                disconnect_progress.stop(success=True)

            self._print_success("Safe shutdown completed")

        except Exception as e:
            self._logger.error(f"Error during shutdown: {e}")
            self._print_error(f"Shutdown error: {e}")

    # UI Helper Methods
    def _get_user_input(self, prompt: str, default: str = "", required: bool = False) -> str:
        """Get user input with default value handling."""
        try:
            user_input = input(prompt).strip()
            if not user_input and default:
                return default
            if required and not user_input:
                return ""
            return user_input
        except (EOFError, KeyboardInterrupt):
            raise KeyboardInterrupt()

    def _print_success(self, message: str) -> None:
        """Print success message with formatting."""
        print(f"{UIColors.SUCCESS}SUCCESS:{UIColors.RESET} {message}")

    def _print_error(self, message: str) -> None:
        """Print error message with formatting."""
        print(f"{UIColors.ERROR}ERROR:{UIColors.RESET} {message}")

    def _print_warning(self, message: str) -> None:
        """Print warning message with formatting."""
        print(f"{UIColors.WARNING}WARNING:{UIColors.RESET} {message}")

    def _print_info(self, message: str) -> None:
        """Print info message with formatting."""
        print(f"{UIColors.INFO}INFO:{UIColors.RESET} {message}")

    def _clear_screen_section(self) -> None:
        """Clear a section of the screen for better organization."""
        print("\n" * 2)

    @property
    def system_state(self) -> SystemState:
        """Get current system state."""
        return self._system_state

    @property
    def test_results(self) -> List[TestResults]:
        """Get all test results."""
        return self._test_results.copy()


def main() -> None:
    """Main application entry point with enhanced error handling."""
    try:
        # Check for colorama availability
        if not COLORAMA_AVAILABLE:
            print("Warning: colorama not available. Install with: pip install colorama")
            print("Running with basic formatting.\n")

        # Create and run enhanced automation system
        automation_system = EnhancedInstrumentAutomationSystem()
        automation_system.run()

    except KeyboardInterrupt:
        print(f"\n\n{UIColors.WARNING}Application interrupted by user{UIColors.RESET}")
    except Exception as e:
        print(f"\n{UIColors.ERROR}Fatal application error: {e}{UIColors.RESET}")
        logging.error(f"Fatal application error: {e}")
    finally:
        print(f"\n{UIColors.INFO}Application terminated{UIColors.RESET}")
        print(f"{UIColors.INFO}Session ended: {UIColors.VALUE}{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}{UIColors.RESET}")


if __name__ == "__main__":
    main()
