#!/usr/bin/env python3
"""
Standalone Keithley Power Supply Control Script

This script provides independent control of Keithley power supplies with
a professional terminal interface for voltage and current control.

Author: Professional Instrument Control Team
Version: 1.0.0
"""

import sys
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

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
    from instrument_control.keithley_power_supply import KeithleyPowerSupply, KeithleyPowerSupplyError
except ImportError as e:
    print(f"{Fore.RED}Error importing power supply module: {e}")
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


class StandalonePowerSupplyController:
    """Standalone Keithley Power Supply Controller."""

    def __init__(self, log_directory: str = "logs"):
        self._log_directory = Path(log_directory)
        self._log_directory.mkdir(exist_ok=True)
        self._setup_logging()
        self._logger = logging.getLogger(self.__class__.__name__)
        self._power_supply: Optional[KeithleyPowerSupply] = None
        self._max_safe_voltage = 30.0
        self._max_safe_current = 3.0

    def _setup_logging(self):
        """Configure logging system."""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
        log_filename = self._log_directory / f"power_supply_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[logging.FileHandler(log_filename, mode='w', encoding='utf-8')]
        )

    def _print_banner(self):
        """Display system banner."""
        width = 70
        print(f"\n{UIColors.SEPARATOR}{'═' * width}{UIColors.RESET}")
        print(f"{UIColors.HEADER}{'KEITHLEY POWER SUPPLY CONTROLLER':^{width}}{UIColors.RESET}")
        print(f"{UIColors.SUBHEADER}{'Standalone Precision Power Control':^{width}}{UIColors.RESET}")
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
        """Discover and connect to power supply."""
        print(f"\n{UIColors.HEADER}DISCOVERING POWER SUPPLY{UIColors.RESET}")
        print(f"{UIColors.SEPARATOR}{'─' * 50}{UIColors.RESET}")

        try:
            resource_manager = pyvisa.ResourceManager()
            resources = list(resource_manager.list_resources())
            
            if not resources:
                print(f"{UIColors.ERROR}No VISA instruments found{UIColors.RESET}")
                return False

            print(f"{UIColors.SUCCESS}Found {len(resources)} VISA resources{UIColors.RESET}")
            
            # Look for Keithley power supplies
            psu_address = None
            for resource in resources:
                try:
                    instrument = resource_manager.open_resource(resource, timeout=5000)
                    idn = instrument.query("*IDN?").strip().upper()
                    instrument.close()
                    
                    if 'KEITHLEY' in idn and any(model in idn for model in ['2230', '2231', '2280', '2260', '2268']):
                        psu_address = resource
                        print(f"{UIColors.SUCCESS}Found Keithley PSU: {resource}{UIColors.RESET}")
                        print(f"  {UIColors.INFO}{idn}{UIColors.RESET}")
                        break
                except:
                    continue

            if not psu_address:
                print(f"{UIColors.WARNING}No Keithley power supply auto-detected{UIColors.RESET}")
                psu_address = self._get_user_input(
                    f"{UIColors.PROMPT}Enter PSU VISA address manually: {UIColors.RESET}",
                    required=True
                )
                if not psu_address:
                    return False

            # Connect to power supply
            print(f"\n{UIColors.INFO}Connecting to power supply...{UIColors.RESET}")
            self._power_supply = KeithleyPowerSupply(psu_address, timeout_ms=15000)
            
            if self._power_supply.connect():
                info = self._power_supply.get_instrument_info()
                if info:
                    print(f"{UIColors.SUCCESS}Connected successfully!{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Model: {info['manufacturer']} {info['model']}{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Channels: {info['max_channels']}{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Max Voltage: {info['max_voltage']}V{UIColors.RESET}")
                    print(f"  {UIColors.INFO}Max Current: {info['max_current']}A{UIColors.RESET}")
                return True
            else:
                print(f"{UIColors.ERROR}Failed to connect to power supply{UIColors.RESET}")
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
        """Run interactive power supply control."""
        if not self._power_supply:
            print(f"{UIColors.ERROR}Power supply not connected{UIColors.RESET}")
            return

        while True:
            try:
                print(f"\n{UIColors.HEADER}POWER SUPPLY CONTROL MENU{UIColors.RESET}")
                print(f"{UIColors.SEPARATOR}{'─' * 40}{UIColors.RESET}")
                print(f"  {UIColors.INFO}1.{UIColors.RESET} Set Voltage and Current")
                print(f"  {UIColors.INFO}2.{UIColors.RESET} Enable/Disable Output")
                print(f"  {UIColors.INFO}3.{UIColors.RESET} Read Measurements")
                print(f"  {UIColors.INFO}4.{UIColors.RESET} Check Status")
                print(f"  {UIColors.INFO}5.{UIColors.RESET} Exit")

                choice = self._get_user_input(f"\n{UIColors.PROMPT}Select option (1-5): {UIColors.RESET}")

                if choice == '1':
                    self._set_voltage_current()
                elif choice == '2':
                    self._toggle_output()
                elif choice == '3':
                    self._read_measurements()
                elif choice == '4':
                    self._check_status()
                elif choice == '5':
                    break
                else:
                    print(f"{UIColors.WARNING}Invalid choice. Please select 1-5.{UIColors.RESET}")

            except KeyboardInterrupt:
                print(f"\n{UIColors.WARNING}Operation interrupted{UIColors.RESET}")
                break

    def _set_voltage_current(self):
        """Set voltage and current for a channel."""
        try:
            # Get channel
            channel_input = self._get_user_input(f"{UIColors.PROMPT}Enter channel (1-3): {UIColors.RESET}", required=True)
            channel = int(channel_input)
            
            if not (1 <= channel <= 3):
                print(f"{UIColors.ERROR}Channel must be 1-3{UIColors.RESET}")
                return

            # Get voltage
            voltage_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter voltage (0.0-{self._max_safe_voltage}V): {UIColors.RESET}",
                required=True
            )
            voltage = float(voltage_input)
            
            if not (0.0 <= voltage <= self._max_safe_voltage):
                print(f"{UIColors.ERROR}Voltage must be 0.0-{self._max_safe_voltage}V{UIColors.RESET}")
                return

            # Get current limit
            current_input = self._get_user_input(
                f"{UIColors.PROMPT}Enter current limit (0.01-{self._max_safe_current}A): {UIColors.RESET}",
                required=True
            )
            current_limit = float(current_input)
            
            if not (0.01 <= current_limit <= self._max_safe_current):
                print(f"{UIColors.ERROR}Current must be 0.01-{self._max_safe_current}A{UIColors.RESET}")
                return

            # Apply settings
            print(f"{UIColors.INFO}Applying settings...{UIColors.RESET}")
            if self._power_supply.set_voltage(channel, voltage):
                print(f"{UIColors.SUCCESS}Voltage set to {voltage}V{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Failed to set voltage{UIColors.RESET}")
                return

            if self._power_supply.set_current_limit(channel, current_limit):
                print(f"{UIColors.SUCCESS}Current limit set to {current_limit}A{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Failed to set current limit{UIColors.RESET}")

        except ValueError:
            print(f"{UIColors.ERROR}Invalid input. Please enter numeric values.{UIColors.RESET}")
        except Exception as e:
            print(f"{UIColors.ERROR}Error setting parameters: {e}{UIColors.RESET}")

    def _toggle_output(self):
        """Enable or disable output."""
        try:
            channel_input = self._get_user_input(f"{UIColors.PROMPT}Enter channel (1-3): {UIColors.RESET}", required=True)
            channel = int(channel_input)
            
            if not (1 <= channel <= 3):
                print(f"{UIColors.ERROR}Channel must be 1-3{UIColors.RESET}")
                return

            action = self._get_user_input(f"{UIColors.PROMPT}Enable output? (y/N): {UIColors.RESET}", default="n").lower()
            
            if action in ['y', 'yes']:
                if self._power_supply.enable_output(channel):
                    print(f"{UIColors.SUCCESS}Output enabled for channel {channel}{UIColors.RESET}")
                else:
                    print(f"{UIColors.ERROR}Failed to enable output{UIColors.RESET}")
            else:
                if self._power_supply.disable_output(channel):
                    print(f"{UIColors.SUCCESS}Output disabled for channel {channel}{UIColors.RESET}")
                else:
                    print(f"{UIColors.ERROR}Failed to disable output{UIColors.RESET}")

        except ValueError:
            print(f"{UIColors.ERROR}Invalid channel number{UIColors.RESET}")
        except Exception as e:
            print(f"{UIColors.ERROR}Error toggling output: {e}{UIColors.RESET}")

    def _read_measurements(self):
        """Read voltage and current measurements."""
        try:
            channel_input = self._get_user_input(f"{UIColors.PROMPT}Enter channel (1-3): {UIColors.RESET}", required=True)
            channel = int(channel_input)
            
            if not (1 <= channel <= 3):
                print(f"{UIColors.ERROR}Channel must be 1-3{UIColors.RESET}")
                return

            print(f"{UIColors.INFO}Reading measurements...{UIColors.RESET}")
            
            voltage = self._power_supply.measure_voltage(channel)
            current = self._power_supply.measure_current(channel)
            
            print(f"\n{UIColors.SUBHEADER}Channel {channel} Measurements:{UIColors.RESET}")
            if voltage is not None:
                print(f"  {UIColors.SUCCESS}Voltage: {UIColors.VALUE}{voltage:.6f}{UIColors.UNIT}V{UIColors.RESET}")
            else:
                print(f"  {UIColors.ERROR}Voltage: Measurement failed{UIColors.RESET}")
                
            if current is not None:
                print(f"  {UIColors.SUCCESS}Current: {UIColors.VALUE}{current:.6f}{UIColors.UNIT}A{UIColors.RESET}")
            else:
                print(f"  {UIColors.ERROR}Current: Measurement failed{UIColors.RESET}")

        except ValueError:
            print(f"{UIColors.ERROR}Invalid channel number{UIColors.RESET}")
        except Exception as e:
            print(f"{UIColors.ERROR}Error reading measurements: {e}{UIColors.RESET}")

    def _check_status(self):
        """Check power supply status."""
        try:
            print(f"{UIColors.INFO}Checking status...{UIColors.RESET}")
            
            info = self._power_supply.get_instrument_info()
            if info:
                print(f"\n{UIColors.SUBHEADER}Power Supply Status:{UIColors.RESET}")
                print(f"  {UIColors.INFO}Model: {info['manufacturer']} {info['model']}{UIColors.RESET}")
                print(f"  {UIColors.INFO}Connected: {UIColors.SUCCESS}Yes{UIColors.RESET}")
                
                # Check each channel status
                for ch in range(1, 4):
                    try:
                        voltage = self._power_supply.measure_voltage(ch)
                        current = self._power_supply.measure_current(ch)
                        print(f"  {UIColors.INFO}Channel {ch}: {voltage:.3f}V, {current:.3f}A{UIColors.RESET}")
                    except:
                        print(f"  {UIColors.WARNING}Channel {ch}: Status unavailable{UIColors.RESET}")
            else:
                print(f"{UIColors.ERROR}Could not retrieve status{UIColors.RESET}")

        except Exception as e:
            print(f"{UIColors.ERROR}Error checking status: {e}{UIColors.RESET}")

    def disconnect(self):
        """Safely disconnect from power supply."""
        if self._power_supply:
            print(f"{UIColors.INFO}Disconnecting power supply...{UIColors.RESET}")
            self._power_supply.disconnect()
            print(f"{UIColors.SUCCESS}Disconnected successfully{UIColors.RESET}")

    def run(self):
        """Main execution method."""
        try:
            self._print_banner()
            
            if not self.discover_and_connect():
                print(f"{UIColors.ERROR}Failed to connect to power supply{UIColors.RESET}")
                return
            
            self.run_interactive_control()
            
        except Exception as e:
            print(f"{UIColors.ERROR}Fatal error: {e}{UIColors.RESET}")
        finally:
            self.disconnect()


if __name__ == "__main__":
    controller = StandalonePowerSupplyController()
    controller.run()
