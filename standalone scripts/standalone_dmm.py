#!/usr/bin/env python3
"""
Standalone Keithley DMM Control Script

This script provides independent control of Keithley DMM6500 multimeters with
a professional terminal interface for precision measurements.

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
    from instrument_control.keithley_dmm import KeithleyDMM6500, KeithleyDMM6500Error
except ImportError as e:
    print(f"{Fore.RED}Error importing DMM module: {e}")
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


class StandaloneDMMController:
    """Standalone Keithley DMM Controller."""

    def __init__(self, log_directory: str = "logs"):
        self._log_directory = Path(log_directory)
        self._log_directory.mkdir(exist_ok=True)
        self._setup_logging()
        self._logger = logging.getLogger(self.__class__.__name__)
        self._dmm: Optional[KeithleyDMM6500] = None

    def _setup_logging(self):
        """Configure logging system."""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
        log_filename = self._log_directory / f"dmm_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[logging.FileHandler(log_filename, mode='w', encoding='utf-8')]
        )

    def _print_banner(self):
        """Display system banner."""
        width = 70
        print(f"\n{UIColors.SEPARATOR}{'═' * width}{UIColors.RESET}")
        print(f"{UIColors.HEADER}{'KEITHLEY DMM CONTROLLER':^{width}}{UIColors.RESET}")
        print(f"{UIColors.SUBHEADER}{'Standalone Precision Measurement System':^{width}}{UIColors.RESET}")
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
        """Discover and connect to DMM."""
        print(f"\n{UIColors.HEADER}DISCOVERING MULTIMETER{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

        try:
            resource_manager = pyvisa.ResourceManager()
            resources = list(resource_manager.list_resources())
            
            if not resources:
                print(f"{UIColors.ERROR}No VISA instruments found{UIColors.RESET}")
                return False

            print(f"{UIColors.SUCCESS}Found {len(resources)} VISA resources{UIColors.RESET}")
            
            # Look for Keithley DMMs
            dmm_address = None
            for resource in resources:
                try:
                    instrument = resource_manager.open_resource(resource, timeout=5000)
                    idn = instrument.query("*IDN?").strip().upper()
                    instrument.close()
                    
                    if 'KEITHLEY' in idn and any(model in idn for model in ['DMM6500', 'DMM7510', '6500', '7510']):
                        dmm_address = resource
                        print(f"{UIColors.SUCCESS}Found Keithley DMM: {resource}{UIColors.RESET}")
                        print(f"  {UIColors.INFO}{idn}{UIColors.RESET}")
                        break
                except:
                    continue

            if not dmm_address:
                print(f"{UIColors.WARNING}No Keithley DMM auto-detected{UIColors.RESET}")
                dmm_address = self._get_user_input(
                    f"{UIColors.PROMPT}Enter DMM VISA address manually: {UIColors.RESET}",
                    required=True
                )
                if not dmm_address:
                    return False

            # Connect to DMM
            print(f"\n{UIColors.INFO}Connecting to multimeter...{UIColors.RESET}")
            self._dmm = KeithleyDMM6500(dmm_address, timeout_ms=30000)
            
            if self._dmm.connect():
                info = self._dmm.get_instrument_info()
                if info:
                    print(f"{UIColors.SUCCESS}Connected successfully!{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Model: {info['manufacturer']} {info['model']}{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Max Voltage Range: {info['max_voltage_range']}V{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Timeout: {info['timeout_ms']}ms{UIColors.RESET}")
                return True
            else:
                print(f"{UIColors.ERROR}Failed to connect to multimeter{UIColors.RESET}")
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
        """Run interactive DMM control."""
        if not self._dmm:
            print(f"{UIColors.ERROR}Multimeter not connected{UIColors.RESET}")
            return

        while True:
            try:
                print(f"\n{UIColors.HEADER}MULTIMETER CONTROL MENU{UIColors.RESET}")
                print(f"{UIColors.SEPARATOR}{'─' * 40}{UIColors.RESET}")
                print(f"  {UIColors.INFO}1.{UIColors.RESET} Measure DC Voltage")
                print(f"  {UIColors.INFO}2.{UIColors.RESET} Measure AC Voltage")
                print(f"  {UIColors.INFO}3.{UIColors.RESET} Measure Resistance")
                print(f"  {UIColors.INFO}4.{UIColors.RESET} Measure Capacitance")
                print(f"  {UIColors.INFO}5.{UIColors.RESET} Multiple Measurements with Statistics")
                print(f"  {UIColors.INFO}6.{UIColors.RESET} Check Status")
                print(f"  {UIColors.INFO}7.{UIColors.RESET} Exit")

                choice = self._get_user_input(f"\n{UIColors.PROMPT}Select option (1-7): {UIColors.RESET}")

                if choice == '1':
                    self._measure_dc_voltage()
                elif choice == '2':
                    self._measure_ac_voltage()
                elif choice == '3':
                    self._measure_resistance()
                elif choice == '4':
                    self._measure_capacitance()
                elif choice == '5':
                    self._multiple_measurements()
                elif choice == '6':
                    self._check_status()
                elif choice == '7':
                    break
                else:
                    print(f"{UIColors.WARNING}Invalid choice. Please select 1-7.{UIColors.RESET}")

            except KeyboardInterrupt:
                print(f"\n{UIColors.WARNING}Operation interrupted{UIColors.RESET}")
                break

    def _measure_dc_voltage(self):
        """Measure DC voltage."""
        try:
            print(f"{UIColors.INFO}Measuring DC voltage...{UIColors.RESET}")
            
            # Choose measurement method
            method = self._get_user_input(
                f"{UIColors.PROMPT}Use fast measurement? (y/N): {UIColors.RESET}",
                default="n"
            ).lower()
            
            if method in ['y', 'yes']:
                voltage = self._dmm.measure_dc_voltage_fast()
                method_name = "Fast"
            else:
                voltage = self._dmm.measure_dc_voltage_precise()
                method_name = "Precise"
            
            if voltage is not None:
                print(f"{UIColors.SUCCESS}{method_name} DC Voltage: {UIColors.VALUE}{voltage:.6f}{UIColors.UNIT}V{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}DC voltage measurement failed{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error measuring DC voltage: {e}{UIColors.RESET}")

    def _measure_ac_voltage(self):
        """Measure AC voltage."""
        try:
            print(f"{UIColors.INFO}Measuring AC voltage...{UIColors.RESET}")
            voltage = self._dmm.measure_ac_voltage()
            
            if voltage is not None:
                print(f"{UIColors.SUCCESS}AC Voltage: {UIColors.VALUE}{voltage:.6f}{UIColors.UNIT}V{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}AC voltage measurement failed{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error measuring AC voltage: {e}{UIColors.RESET}")

    def _measure_resistance(self):
        """Measure resistance."""
        try:
            print(f"{UIColors.INFO}Measuring resistance...{UIColors.RESET}")
            resistance = self._dmm.measure_resistance()
            
            if resistance is not None:
                # Format resistance with appropriate units
                if resistance >= 1e6:
                    print(f"{UIColors.SUCCESS}Resistance: {UIColors.VALUE}{resistance/1e6:.3f}{UIColors.UNIT}MΩ{UIColors.RESET}")
                elif resistance >= 1e3:
                    print(f"{UIColors.SUCCESS}Resistance: {UIColors.VALUE}{resistance/1e3:.3f}{UIColors.UNIT}kΩ{UIColors.RESET}")
                else:
                    print(f"{UIColors.SUCCESS}Resistance: {UIColors.VALUE}{resistance:.3f}{UIColors.UNIT}Ω{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Resistance measurement failed{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error measuring resistance: {e}{UIColors.RESET}")

    def _measure_capacitance(self):
        """Measure capacitance."""
        try:
            print(f"{UIColors.INFO}Measuring capacitance...{UIColors.RESET}")
            capacitance = self._dmm.measure_capacitance()
            
            if capacitance is not None:
                # Format capacitance with appropriate units
                if capacitance >= 1e-6:
                    print(f"{UIColors.SUCCESS}Capacitance: {UIColors.VALUE}{capacitance*1e6:.3f}{UIColors.UNIT}μF{UIColors.RESET}")
                elif capacitance >= 1e-9:
                    print(f"{UIColors.SUCCESS}Capacitance: {UIColors.VALUE}{capacitance*1e9:.3f}{UIColors.UNIT}nF{UIColors.RESET}")
                else:
                    print(f"{UIColors.SUCCESS}Capacitance: {UIColors.VALUE}{capacitance*1e12:.3f}{UIColors.UNIT}pF{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Capacitance measurement failed{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error measuring capacitance: {e}{UIColors.RESET}")

    def _multiple_measurements(self):
        """Perform multiple measurements with statistics."""
        try:
            # Get measurement parameters
            count_input = self._get_user_input(
                f"{UIColors.PROMPT}Number of measurements (1-100): {UIColors.RESET}",
                default="10"
            )
            count = int(count_input)
            if not (1 <= count <= 100):
                print(f"{UIColors.ERROR}Count must be 1-100{UIColors.RESET}")
                return

            interval_input = self._get_user_input(
                f"{UIColors.PROMPT}Interval between measurements (0.1-10.0s): {UIColors.RESET}",
                default="0.5"
            )
            interval = float(interval_input)
            if not (0.1 <= interval <= 10.0):
                print(f"{UIColors.ERROR}Interval must be 0.1-10.0 seconds{UIColors.RESET}")
                return

            print(f"\n{UIColors.INFO}Performing {count} DC voltage measurements...{UIColors.RESET}")
            
            # Perform measurements
            measurements = []
            for i in range(count):
                print(f"{UIColors.INFO}Measurement {i+1}/{count}...{UIColors.RESET}", end='\r')
                voltage = self._dmm.measure_dc_voltage_fast()
                if voltage is not None:
                    measurements.append(voltage)
                time.sleep(interval)

            if not measurements:
                print(f"{UIColors.ERROR}No successful measurements{UIColors.RESET}")
                return

            # Calculate statistics
            import statistics
            mean_val = statistics.mean(measurements)
            std_dev = statistics.stdev(measurements) if len(measurements) > 1 else 0.0
            min_val = min(measurements)
            max_val = max(measurements)

            print(f"\n{UIColors.SUBHEADER}Measurement Statistics ({len(measurements)} samples):{UIColors.RESET}")
            print(f"  {UIColors.SUCCESS}Mean: {UIColors.VALUE}{mean_val:.6f}{UIColors.UNIT}V{UIColors.RESET}")
            print(f"  {UIColors.INFO}Std Dev: {UIColors.VALUE}{std_dev:.6f}{UIColors.UNIT}V{UIColors.RESET}")
            print(f"  {UIColors.INFO}Min: {UIColors.VALUE}{min_val:.6f}{UIColors.UNIT}V{UIColors.RESET}")
            print(f"  {UIColors.INFO}Max: {UIColors.VALUE}{max_val:.6f}{UIColors.UNIT}V{UIColors.RESET}")
            print(f"  {UIColors.INFO}Range: {UIColors.VALUE}{max_val-min_val:.6f}{UIColors.UNIT}V{UIColors.RESET}")

        except ValueError:
            print(f"{UIColors.ERROR}Invalid input. Please enter numeric values.{UIColors.RESET}")
        except Exception as e:
            print(f"{UIColors.ERROR}Error in multiple measurements: {e}{UIColors.RESET}")

    def _check_status(self):
        """Check DMM status."""
        try:
            print(f"{UIColors.INFO}Checking status...{UIColors.RESET}")
            
            info = self._dmm.get_instrument_info()
            if info:
                print(f"\n{UIColors.SUBHEADER}Multimeter Status:{UIColors.RESET}")
                print(f"  {UIColors.INFO}Model: {info['manufacturer']} {info['model']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Serial: {info['serial_number']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Firmware: {info['firmware_version']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Connected: {UIColors.SUCCESS}Yes{UIColors.RESET}")
                print(f"  {UIColors.INFO}Max Voltage Range: {info['max_voltage_range']}V{UIColors.RESET}")
                
                # Check for errors
                if info['current_errors'] != 'None':
                    print(f"  {UIColors.WARNING}Errors: {info['current_errors']}{UIColors.RESET}")
                else:
                    print(f"  {UIColors.SUCCESS}Errors: None{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Could not retrieve status{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error checking status: {e}{UIColors.RESET}")

    def disconnect(self):
        """Safely disconnect from DMM."""
        if self._dmm:
            print(f"{UIColors.INFO}Disconnecting multimeter...{UIColors.RESET}")
            self._dmm.disconnect()
            print(f"{UIColors.SUCCESS}Disconnected successfully{UIColors.RESET}")

    def run(self):
        """Main execution method."""
        try:
            self._print_banner()
            
            if not self.discover_and_connect():
                print(f"{UIColors.ERROR}Failed to connect to multimeter{UIColors.RESET}")
                return
            
            self.run_interactive_control()
            
        except Exception as e:
            print(f"{UIColors.ERROR}Fatal error: {e}{UIColors.RESET}")
        finally:
            self.disconnect()


if __name__ == "__main__":
    controller = StandaloneDMMController()
    controller.run()
