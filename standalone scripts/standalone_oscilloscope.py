#!/usr/bin/env python3
"""
Standalone Keysight Oscilloscope Control Script

This script provides independent control of Keysight DSOX6004A oscilloscopes with
a professional terminal interface for waveform capture and analysis.

Author: Professional Instrument Control Team
Version: 1.0.0
"""

import sys
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List

# Enhanced UI imports
try:
    from colorama import init, Fore, Back, Style
    init(autoreset=True)
    COLORAMA_AVAILABLE = True
except ImportError:
    COLORAMA_AVAILABLE = False
    class DummyColor:
        RED = GREEN = YELLOW = BLUE = MAGENTA = CYAN = WHITE = RESET = ""
        BRIGHT = DIM = ""
    Fore = Back = Style = DummyColor()

# Import instrument control module
try:
    from instrument_control.keysight_oscilloscope import KeysightDSOX6004A, KeysightDSOX6004AError
except ImportError as e:
    print(f"{Fore.RED}Error importing oscilloscope module: {e}")
    sys.exit(1)

try:
    import pyvisa
except ImportError as e:
    print(f"{Fore.RED}PyVISA library is required. Install with: pip install pyvisa")
    sys.exit(1)


class UIColors:
    """Professional color scheme for terminal interface."""
    SUCCESS = Fore.GREEN + Style.BRIGHT
    ERROR = Fore.RED + Style.BRIGHT
    WARNING = Fore.YELLOW + Style.BRIGHT
    INFO = Fore.CYAN
    HEADER = Fore.BLUE + Style.BRIGHT
    SUBHEADER = Fore.MAGENTA + Style.BRIGHT
    VALUE = Fore.WHITE + Style.BRIGHT
    UNIT = Fore.CYAN + Style.DIM
    PROMPT = Fore.YELLOW
    INPUT = Fore.WHITE
    SEPARATOR = Fore.BLUE + Style.DIM
    RESET = Style.RESET_ALL


class StandaloneOscilloscopeController:
    """Standalone Keysight Oscilloscope Controller."""

    def __init__(self, log_directory: str = "logs"):
        self._log_directory = Path(log_directory)
        self._log_directory.mkdir(exist_ok=True)
        self._setup_logging()
        self._logger = logging.getLogger(self.__class__.__name__)
        self._oscilloscope: Optional[KeysightDSOX6004A] = None

    def _setup_logging(self):
        """Configure logging system."""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
        log_filename = self._log_directory / f"oscilloscope_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[logging.FileHandler(log_filename, mode='w', encoding='utf-8')]
        )

    def _print_banner(self):
        """Display system banner."""
        width = 70
        print(f"\n{UIColors.SEPARATOR}{'═' * width}{UIColors.RESET}")
        print(f"{UIColors.HEADER}{'KEYSIGHT OSCILLOSCOPE CONTROLLER':^{width}}{UIColors.RESET}")
        print(f"{UIColors.SUBHEADER}{'Standalone Waveform Capture & Analysis':^{width}}{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'═' * width}{UIColors.RESET}")

    def _get_user_input(self, prompt: str, required: bool = False, default: str = "") -> str:
        """Get user input with validation."""
        while True:
            try:
                user_input = input(prompt).strip()
                if not user_input and default:
                    return default
                if required and not user_input:
                    print(f"{UIColors.ERROR}Input required. Please try again.{UIColors.RESET}")
                    continue
                return user_input
            except KeyboardInterrupt:
                return ""

    def discover_and_connect(self) -> bool:
        """Discover and connect to oscilloscope."""
        print(f"\n{UIColors.HEADER}DISCOVERING OSCILLOSCOPE{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

        try:
            resource_manager = pyvisa.ResourceManager()
            resources = list(resource_manager.list_resources())
            
            if not resources:
                print(f"{UIColors.ERROR}No VISA instruments found{UIColors.RESET}")
                return False

            print(f"{UIColors.SUCCESS}Found {len(resources)} VISA resources{UIColors.RESET}")
            
            # Look for Keysight/Agilent oscilloscopes
            scope_address = None
            for resource in resources:
                try:
                    instrument = resource_manager.open_resource(resource, timeout=5000)
                    idn = instrument.query("*IDN?").strip().upper()
                    instrument.close()
                    
                    if ('KEYSIGHT' in idn or 'AGILENT' in idn) and any(model in idn.replace('-', '') for model in ['DSOX', 'MSOX']):
                        scope_address = resource
                        print(f"{UIColors.SUCCESS}Found Keysight/Agilent Oscilloscope: {resource}{UIColors.RESET}")
                        print(f"  {UIColors.INFO}{idn}{UIColors.RESET}")
                        break
                except:
                    continue

            if not scope_address:
                print(f"{UIColors.WARNING}No Keysight/Agilent oscilloscope auto-detected{UIColors.RESET}")
                scope_address = self._get_user_input(
                    f"{UIColors.PROMPT}Enter oscilloscope VISA address manually: {UIColors.RESET}",
                    required=True
                )
                if not scope_address:
                    return False

            # Connect to oscilloscope
            print(f"\n{UIColors.INFO}Connecting to oscilloscope...{UIColors.RESET}")
            self._oscilloscope = KeysightDSOX6004A(scope_address, timeout_ms=15000)
            
            if self._oscilloscope.connect():
                info = self._oscilloscope.get_instrument_info()
                if info:
                    print(f"{UIColors.SUCCESS}Connected successfully!{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Model: {info['manufacturer']} {info['model']}{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Bandwidth: {info['bandwidth_hz']/1e9:.1f} GHz{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Channels: {info['max_channels']}{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Max Sample Rate: {info['max_sample_rate']/1e9:.1f} GSa/s{UIColors.RESET}")
                return True
            else:
                print(f"{UIColors.ERROR}Failed to connect to oscilloscope{UIColors.RESET}")
                return False

        except Exception as e:
            print(f"{UIColors.ERROR}Discovery failed: {e}{UIColors.RESET}")
            return False
        finally:
            try:
                resource_manager.close()
            except:
                pass

    def run_interactive_control(self):
        """Run interactive oscilloscope control."""
        if not self._oscilloscope:
            print(f"{UIColors.ERROR}Oscilloscope not connected{UIColors.RESET}")
            return

        while True:
            try:
                print(f"\n{UIColors.HEADER}OSCILLOSCOPE CONTROL MENU{UIColors.RESET}")
                print(f"{UIColors.SEPARATOR}{'─' * 40}{UIColors.RESET}")
                print(f"  {UIColors.INFO}1.{UIColors.RESET} Configure Channel")
                print(f"  {UIColors.INFO}2.{UIColors.RESET} Capture Screenshot")
                print(f"  {UIColors.INFO}3.{UIColors.RESET} Setup Output Directories")
                print(f"  {UIColors.INFO}4.{UIColors.RESET} Check Status")
                print(f"  {UIColors.INFO}5.{UIColors.RESET} Exit")

                choice = self._get_user_input(f"\n{UIColors.PROMPT}Select option (1-5): {UIColors.RESET}")

                if choice == '1':
                    self._configure_channel()
                elif choice == '2':
                    self._capture_screenshot()
                elif choice == '3':
                    self._setup_directories()
                elif choice == '4':
                    self._check_status()
                elif choice == '5':
                    break
                else:
                    print(f"{UIColors.WARNING}Invalid choice. Please select 1-5.{UIColors.RESET}")

            except KeyboardInterrupt:
                print(f"\n{UIColors.WARNING}Operation interrupted{UIColors.RESET}")
                break

    def _configure_channel(self):
        """Configure oscilloscope channel settings."""
        try:
            info = self._oscilloscope.get_instrument_info()
            max_channels = info['max_channels'] if info else 4

            # Get channel number
            channel_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter channel number (1-{max_channels}): {UIColors.RESET}",
                required=True
            )
            channel = int(channel_input)
            
            if not (1 <= channel <= max_channels):
                print(f"{UIColors.ERROR}Channel must be 1-{max_channels}{UIColors.RESET}")
                return

            # Get vertical scale
            print(f"\n{UIColors.SUBHEADER}Available vertical scales:{UIColors.RESET}")
            scales = [1e-3, 2e-3, 5e-3, 10e-3, 20e-3, 50e-3, 100e-3, 200e-3, 500e-3, 1.0, 2.0, 5.0, 10.0]
            for i, scale in enumerate(scales):
                if scale < 1.0:
                    print(f"  {UIColors.INFO}{i+1:2d}.{UIColors.RESET} {scale*1000:.0f}mV/div")
                else:
                    print(f"  {UIColors.INFO}{i+1:2d}.{UIColors.RESET} {scale:.0f}V/div")

            scale_input = self._get_user_input(
                f"\n{UIColors.PROMPT}Select scale (1-{len(scales)}) or enter custom value: {UIColors.RESET}",
                required=True
            )

            try:
                scale_index = int(scale_input) - 1
                if 0 <= scale_index < len(scales):
                    vertical_scale = scales[scale_index]
                else:
                    print(f"{UIColors.ERROR}Invalid scale selection{UIColors.RESET}")
                    return
            except ValueError:
                # Try to parse as custom value
                vertical_scale = float(scale_input)

            # Get vertical offset
            offset_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter vertical offset (V, default 0.0): {UIColors.RESET}",
                default="0.0"
            )
            vertical_offset = float(offset_input)

            # Get coupling
            coupling_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter coupling (DC/AC/GND, default DC): {UIColors.RESET}",
                default="DC"
            ).upper()

            if coupling_input not in ['DC', 'AC', 'GND']:
                print(f"{UIColors.ERROR}Invalid coupling. Must be DC, AC, or GND{UIColors.RESET}")
                return

            # Get probe attenuation
            probe_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter probe attenuation (1x, 10x, etc., default 1.0): {UIColors.RESET}",
                default="1.0"
            )
            probe_attenuation = float(probe_input)

            # Apply configuration
            print(f"{UIColors.INFO}Configuring channel {channel}...{UIColors.RESET}")
            success = self._oscilloscope.configure_channel(
                channel, vertical_scale, vertical_offset, coupling_input, probe_attenuation
            )

            if success:
                print(f"{UIColors.SUCCESS}Channel {channel} configured successfully{UIColors.RESET}")
                if vertical_scale < 1.0:
                    scale_str = f"{vertical_scale*1000:.0f}mV/div"
                else:
                    scale_str = f"{vertical_scale:.1f}V/div"
                print(f"  {UIColors.INFO}Scale: {scale_str}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Offset: {vertical_offset}V{UIColors.RESET}")
                print(f"  {UIColors.INFO}Coupling: {coupling_input}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Probe: {probe_attenuation}x{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Failed to configure channel{UIColors.RESET}")

        except ValueError:
            print(f"{UIColors.ERROR}Invalid input. Please enter numeric values.{UIColors.RESET}")
        except Exception as e:
            print(f"{UIColors.ERROR}Error configuring channel: {e}{UIColors.RESET}")

    def _capture_screenshot(self):
        """Capture oscilloscope screenshot."""
        try:
            # Get filename
            default_filename = f"scope_screenshot_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}.png"
            filename_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter filename (default: {default_filename}): {UIColors.RESET}",
                default=default_filename
            )

            # Get image format
            format_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter image format (PNG/BMP, default PNG): {UIColors.RESET}",
                default="PNG"
            ).upper()

            if format_input not in ['PNG', 'BMP']:
                print(f"{UIColors.ERROR}Invalid format. Must be PNG or BMP{UIColors.RESET}")
                return

            # Include timestamp option
            timestamp_input = self._get_user_input(
                f"{UIColors.PROMPT}Include timestamp in filename? (y/N): {UIColors.RESET}",
                default="n"
            ).lower()
            include_timestamp = timestamp_input in ['y', 'yes']

            print(f"{UIColors.INFO}Capturing screenshot...{UIColors.RESET}")
            screenshot_path = self._oscilloscope.capture_screenshot(
                filename_input, format_input, include_timestamp
            )

            if screenshot_path:
                print(f"{UIColors.SUCCESS}Screenshot saved successfully!{UIColors.RESET}")
                print(f"  {UIColors.INFO}File: {screenshot_path}{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Failed to capture screenshot{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error capturing screenshot: {e}{UIColors.RESET}")

    def _setup_directories(self):
        """Setup output directories for screenshots and data."""
        try:
            print(f"{UIColors.INFO}Setting up output directories...{UIColors.RESET}")
            self._oscilloscope.setup_output_directories()
            
            print(f"{UIColors.SUCCESS}Output directories created:{UIColors.RESET}")
            print(f"  {UIColors.INFO}Screenshots: oscilloscope_screenshots/{UIColors.RESET}")
            print(f"  {UIColors.INFO}Data: oscilloscope_data/{UIColors.RESET}")
            print(f"  {UIColors.INFO}Graphs: oscilloscope_graphs/{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error setting up directories: {e}{UIColors.RESET}")

    def _check_status(self):
        """Check oscilloscope status."""
        try:
            print(f"{UIColors.INFO}Checking status...{UIColors.RESET}")
            
            info = self._oscilloscope.get_instrument_info()
            if info:
                print(f"\n{UIColors.SUBHEADER}Oscilloscope Status:{UIColors.RESET}")
                print(f"  {UIColors.INFO}Model: {info['manufacturer']} {info['model']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Serial: {info['serial_number']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Firmware: {info['firmware_version']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Connected: {UIColors.SUCCESS}Yes{UIColors.RESET}")
                print(f"  {UIColors.INFO}Bandwidth: {info['bandwidth_hz']/1e9:.1f} GHz{UIColors.RESET}")
                print(f"  {UIColors.INFO}Channels: {info['max_channels']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Max Sample Rate: {info['max_sample_rate']/1e9:.1f} GSa/s{UIColors.RESET}")
                print(f"  {UIColors.INFO}Max Memory Depth: {info['max_memory_depth']/1e6:.1f} MSa{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Could not retrieve status{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error checking status: {e}{UIColors.RESET}")

    def disconnect(self):
        """Safely disconnect from oscilloscope."""
        if self._oscilloscope:
            print(f"{UIColors.INFO}Disconnecting oscilloscope...{UIColors.RESET}")
            self._oscilloscope.disconnect()
            print(f"{UIColors.SUCCESS}Disconnected successfully{UIColors.RESET}")

    def run(self):
        """Main execution method."""
        try:
            self._print_banner()
            
            if not self.discover_and_connect():
                print(f"{UIColors.ERROR}Failed to connect to oscilloscope{UIColors.RESET}")
                return
            
            # Setup directories on startup
            try:
                self._oscilloscope.setup_output_directories()
                print(f"{UIColors.SUCCESS}Output directories initialized{UIColors.RESET}")
            except Exception as e:
                print(f"{UIColors.WARNING}Could not setup directories: {e}{UIColors.RESET}")
            
            self.run_interactive_control()
            
        except Exception as e:
            print(f"{UIColors.ERROR}Fatal error: {e}{UIColors.RESET}")
        finally:
            self.disconnect()


if __name__ == "__main__":
    controller = StandaloneOscilloscopeController()
    controller.run()
