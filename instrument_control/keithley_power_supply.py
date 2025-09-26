#!/usr/bin/env python3
"""
Keithley Power Supply Control Library

This module provides a comprehensive interface for controlling Keithley multi-channel
power supplies via SCPI commands over VISA communication protocol.

Module: instrument_control.keithley_power_supply
Author: Professional Instrument Control Team
Version: 1.0.0
License: MIT
Dependencies: pyvisa

Supported Models:
    - 2230-30-3 (Triple output, 30V/3A per channel)
    - 2231A-30-3 (Triple output with USB data logging)
    - 2280S Series (Single output, high current)
    - 2260B/2268 Series (Programmable DC loads)

Features:
    - Multi-channel voltage and current control
    - Real-time output monitoring and measurement
    - Over-voltage and over-current protection
    - Output sequencing and ramping capabilities
    - Comprehensive safety interlocks

Usage:
    from instrument_control.keithley_power_supply import KeithleyPowerSupply

    power_supply = KeithleyPowerSupply('USB0::0x05E6::0x2230::9103456::INSTR')
    power_supply.connect()
    power_supply.configure_channel(channel=1, voltage=5.0, current_limit=1.0)
    power_supply.enable_channel_output(channel=1)
    voltage, current = power_supply.measure_channel_output(channel=1)
    power_supply.disconnect()
"""

import logging
import time
from typing import Optional, Dict, Any, List, Tuple, Union
from dataclasses import dataclass
from enum import Enum

try:
    import pyvisa
    from pyvisa.errors import VisaIOError
except ImportError as e:
    raise ImportError(
        "PyVISA library is required. Install with: pip install pyvisa"
    ) from e


class OutputState(Enum):
    """Enumeration of channel output states."""
    OFF = "OFF"
    ON = "ON"


class ProtectionState(Enum):
    """Enumeration of protection states."""
    CLEAR = "CLEAR"
    TRIPPED = "TRIPPED"


@dataclass
class ChannelConfiguration:
    """Data class for channel configuration parameters."""
    voltage: float
    current_limit: float
    ovp_level: Optional[float] = None  # Over-voltage protection level
    ocp_level: Optional[float] = None  # Over-current protection level
    enabled: bool = False


@dataclass
class ChannelMeasurement:
    """Data class for channel measurement results."""
    voltage: float
    current: float
    power: float
    timestamp: float


class KeithleyPowerSupplyError(Exception):
    """Custom exception for Keithley power supply errors."""
    pass


class KeithleyPowerSupply:
    """
    Control interface for Keithley multi-channel power supplies.

    This class provides methods for precision power supply control with
    comprehensive safety features and real-time monitoring capabilities.
    All methods follow IEEE 488.2 and SCPI standards.

    Attributes:
        visa_address (str): VISA resource identifier string
        timeout_ms (int): Communication timeout in milliseconds
        max_channels (int): Number of output channels available
        max_voltage (float): Maximum output voltage per channel
        max_current (float): Maximum output current per channel
    """

    def __init__(self, visa_address: str, timeout_ms: int = 10000) -> None:
        """
        Initialize power supply control instance.

        Args:
            visa_address: VISA resource string (e.g., 'USB0::0x05E6::0x2230::9103456::INSTR')
            timeout_ms: Communication timeout in milliseconds

        Raises:
            ValueError: If visa_address is empty or invalid format
        """
        if not visa_address or not isinstance(visa_address, str):
            raise ValueError("visa_address must be a non-empty string")

        # Store configuration parameters
        self._visa_address = visa_address
        self._timeout_ms = timeout_ms

        # Initialize VISA communication objects
        self._resource_manager: Optional[pyvisa.ResourceManager] = None
        self._instrument: Optional[pyvisa.Resource] = None
        self._is_connected = False

        # Initialize logging for this instance
        self._logger = logging.getLogger(f'{self.__class__.__name__}.{id(self)}')

        # Initialize instrument specifications (will be determined during connection)
        self.max_channels = 3    # Default for 2230 series
        self.max_voltage = 30.0  # Maximum voltage per channel
        self.max_current = 3.0   # Maximum current per channel

        # Store channel configurations
        self._channel_configs: Dict[int, ChannelConfiguration] = {}

        # Define voltage and current step limits for safety
        self._max_voltage_step = 5.0   # Maximum voltage change per step (V)
        self._max_current_step = 1.0   # Maximum current change per step (A)

        # Define settling times for different operations
        self._voltage_settling_time = 0.1   # Time to wait after voltage change (s)
        self._current_settling_time = 0.1   # Time to wait after current change (s)
        self._output_enable_time = 0.2      # Time to wait after output enable (s)

    def connect(self) -> bool:
        """
        Establish communication with the power supply.

        This method creates the VISA resource manager, opens the instrument
        connection, performs model detection, and initializes all channels
        to safe default states.

        Returns:
            True if connection successful, False otherwise

        Raises:
            KeithleyPowerSupplyError: If critical connection error occurs
        """
        try:
            # Create VISA resource manager instance
            self._resource_manager = pyvisa.ResourceManager()
            self._logger.info("VISA resource manager created successfully")

            # Open connection to specified instrument
            self._instrument = self._resource_manager.open_resource(self._visa_address)
            self._logger.info(f"Opened connection to {self._visa_address}")

            # Configure communication parameters
            self._instrument.timeout = self._timeout_ms
            self._instrument.read_termination = '\n'  # Line feed termination
            self._instrument.write_termination = '\n'  # Line feed termination

            # Verify instrument communication with identification query
            identification = self._instrument.query("*IDN?")
            self._logger.info(f"Instrument identification: {identification.strip()}")

            # Parse identification to determine model capabilities
            idn_parts = identification.strip().split(',')
            manufacturer = idn_parts[0] if len(idn_parts) > 0 else ""
            model = idn_parts[1] if len(idn_parts) > 1 else ""

            # Validate manufacturer
            if "KEITHLEY" not in manufacturer.upper():
                self._logger.warning(f"Unexpected manufacturer: {manufacturer}")

            # Determine model-specific parameters
            self._configure_model_parameters(model.strip())

            # Perform instrument initialization sequence
            self._instrument.write("*CLS")  # Clear status registers
            self._instrument.write("*RST")  # Reset to known state
            time.sleep(2.0)  # Allow reset to complete

            # Initialize all channels to safe states
            self._initialize_channels_safe_state()

            # Verify instrument is responsive after initialization
            self._instrument.query("*OPC?")  # Operation complete query

            # Mark connection as established
            self._is_connected = True
            self._logger.info(f"Successfully connected to Keithley {model}")

            return True

        except VisaIOError as e:
            self._logger.error(f"VISA communication error during connection: {e}")
            self._cleanup_connection()
            return False

        except Exception as e:
            self._logger.error(f"Unexpected error during connection: {e}")
            self._cleanup_connection()
            raise KeithleyPowerSupplyError(f"Connection failed: {e}") from e

    def disconnect(self) -> None:
        """
        Safely disconnect from power supply and release resources.

        This method ensures all outputs are disabled, puts the instrument
        in a safe state, and properly cleans up resources.
        """
        try:
            if self._instrument is not None:
                # Disable all outputs for safety before disconnection
                for channel in range(1, self.max_channels + 1):
                    try:
                        self._disable_channel_output_internal(channel)
                    except Exception as e:
                        self._logger.warning(f"Failed to disable channel {channel}: {e}")

                # Put instrument in safe state
                self._instrument.write("*CLS")  # Clear status registers

                # Close instrument connection
                self._instrument.close()
                self._logger.info("Instrument connection closed")

            if self._resource_manager is not None:
                # Close resource manager
                self._resource_manager.close()
                self._logger.info("VISA resource manager closed")

        except Exception as e:
            self._logger.error(f"Error during disconnection: {e}")

        finally:
            # Reset connection state and object references
            self._cleanup_connection()
            self._logger.info("Disconnection completed")

    def configure_channel(self, 
                         channel: int, 
                         voltage: float, 
                         current_limit: float,
                         ovp_level: Optional[float] = None,
                         enable_output: bool = False) -> bool:
        """
        Configure channel output parameters with comprehensive validation.

        Args:
            channel: Channel number (1 to max_channels)
            voltage: Output voltage in volts
            current_limit: Current limit in amperes
            ovp_level: Over-voltage protection level in volts (None for default)
            enable_output: Whether to enable output after configuration

        Returns:
            True if configuration successful, False otherwise

        Raises:
            KeithleyPowerSupplyError: If channel or parameter values are invalid
        """
        if not self._is_connected:
            raise KeithleyPowerSupplyError("Power supply not connected")

        # Validate channel number
        if not (1 <= channel <= self.max_channels):
            raise ValueError(f"Channel must be 1-{self.max_channels}, got {channel}")

        # Validate voltage range
        if not (0.0 <= voltage <= self.max_voltage):
            raise ValueError(f"Voltage must be 0-{self.max_voltage}V, got {voltage}V")

        # Validate current limit range
        if not (0.001 <= current_limit <= self.max_current):
            raise ValueError(f"Current limit must be 0.001-{self.max_current}A, got {current_limit}A")

        try:
            # Ensure output is disabled during configuration for safety
            self._disable_channel_output_internal(channel)

            # Select the specified channel
            self._instrument.write(f":INSTrument:SELect CH{channel}")

            # Configure output voltage with gradual ramping for safety
            current_voltage = self._get_channel_voltage_setpoint(channel)
            if current_voltage is not None:
                self._ramp_voltage_safely(channel, current_voltage, voltage)
            else:
                # Direct setting if current value cannot be read
                self._instrument.write(f":VOLTage {voltage:.6f}")
                time.sleep(self._voltage_settling_time)

            # Configure current limit
            self._instrument.write(f":CURRent {current_limit:.6f}")
            time.sleep(self._current_settling_time)

            # Configure over-voltage protection if specified
            if ovp_level is not None:
                if ovp_level > voltage:
                    self._instrument.write(f":VOLTage:PROTection:LEVel {ovp_level:.6f}")
                    self._instrument.write(":VOLTage:PROTection:STATe ON")
                else:
                    self._logger.warning(f"OVP level {ovp_level}V must be greater than voltage {voltage}V")

            # Verify configuration was applied correctly
            actual_voltage = float(self._instrument.query(":VOLTage?"))
            actual_current = float(self._instrument.query(":CURRent?"))

            # Check for significant discrepancies
            if abs(actual_voltage - voltage) > 0.001:  # 1mV tolerance
                self._logger.warning(f"Voltage setpoint mismatch: requested {voltage:.3f}V, "
                                   f"actual {actual_voltage:.3f}V")

            if abs(actual_current - current_limit) > 0.001:  # 1mA tolerance
                self._logger.warning(f"Current limit mismatch: requested {current_limit:.3f}A, "
                                   f"actual {actual_current:.3f}A")

            # Store configuration
            self._channel_configs[channel] = ChannelConfiguration(
                voltage=actual_voltage,
                current_limit=actual_current,
                ovp_level=ovp_level,
                enabled=False  # Will be set True if enable_output is requested
            )

            # Enable output if requested
            if enable_output:
                success = self.enable_channel_output(channel)
                if success:
                    self._channel_configs[channel].enabled = True

            self._logger.info(f"Channel {channel} configured: {actual_voltage:.6f}V, "
                            f"{actual_current:.6f}A limit, Output: "
                            f"{'Enabled' if enable_output else 'Disabled'}")

            return True

        except (VisaIOError, ValueError) as e:
            self._logger.error(f"Failed to configure channel {channel}: {e}")
            return False

    def enable_channel_output(self, channel: int) -> bool:
        """
        Enable output for specified channel with safety checks.

        Args:
            channel: Channel number (1 to max_channels)

        Returns:
            True if output enabled successfully, False otherwise
        """
        if not self._is_connected:
            raise KeithleyPowerSupplyError("Power supply not connected")

        if not (1 <= channel <= self.max_channels):
            raise ValueError(f"Channel must be 1-{self.max_channels}, got {channel}")

        try:
            self._logger.info(f"Enabling output on channel {channel}")

            # Select the specified channel
            self._instrument.write(f":INSTrument:SELect CH{channel}")

            # Enable the output
            self._instrument.write(":OUTPut ON")

            # Wait for output to stabilize
            time.sleep(self._output_enable_time)

            # Verify output is enabled
            output_state = self._instrument.query(":OUTPut?").strip()
            is_enabled = output_state in ["1", "ON"]

            if is_enabled:
                # Update configuration if it exists
                if channel in self._channel_configs:
                    self._channel_configs[channel].enabled = True

                self._logger.info(f"Channel {channel} output enabled successfully")
                return True
            else:
                self._logger.error(f"Failed to enable channel {channel} output")
                return False

        except (VisaIOError, ValueError) as e:
            self._logger.error(f"Failed to enable output on channel {channel}: {e}")
            return False

    def disable_channel_output(self, channel: int) -> bool:
        """
        Disable output for specified channel.

        Args:
            channel: Channel number (1 to max_channels)

        Returns:
            True if output disabled successfully, False otherwise
        """
        if not self._is_connected:
            raise KeithleyPowerSupplyError("Power supply not connected")

        if not (1 <= channel <= self.max_channels):
            raise ValueError(f"Channel must be 1-{self.max_channels}, got {channel}")

        return self._disable_channel_output_internal(channel)

    def measure_channel_output(self, channel: int) -> Optional[Tuple[float, float]]:
        """
        Measure actual output voltage and current for specified channel.

        Args:
            channel: Channel number (1 to max_channels)

        Returns:
            Tuple of (voltage, current) measurements, or None if measurement failed
        """
        if not self._is_connected:
            self._logger.error("Cannot measure output: power supply not connected")
            return None

        if not (1 <= channel <= self.max_channels):
            self._logger.error(f"Invalid channel {channel}")
            return None

        try:
            # Select the specified channel
            self._instrument.write(f":INSTrument:SELect CH{channel}")

            # Measure output voltage
            voltage_str = self._instrument.query(":MEASure:VOLTage?")
            voltage = float(voltage_str.strip())

            # Measure output current
            current_str = self._instrument.query(":MEASure:CURRent?")
            current = float(current_str.strip())

            self._logger.debug(f"Channel {channel} measurements: {voltage:.6f}V, {current:.6f}A")

            return voltage, current

        except (VisaIOError, ValueError) as e:
            self._logger.error(f"Failed to measure output on channel {channel}: {e}")
            return None

    def disable_all_outputs(self) -> bool:
        """
        Disable all channel outputs for emergency shutdown.

        Returns:
            True if all outputs disabled successfully, False otherwise
        """
        if not self._is_connected:
            self._logger.error("Cannot disable outputs: power supply not connected")
            return False

        success = True
        self._logger.info("Disabling all channel outputs")

        for channel in range(1, self.max_channels + 1):
            if not self._disable_channel_output_internal(channel):
                success = False

        if success:
            self._logger.info("All outputs disabled successfully")
        else:
            self._logger.warning("Some outputs may still be enabled")

        return success

    def get_instrument_info(self) -> Optional[Dict[str, Any]]:
        """
        Retrieve comprehensive instrument information and status.

        Returns:
            Dictionary containing instrument details, or None if query failed
        """
        if not self._is_connected:
            return None

        try:
            # Query instrument identification
            idn_response = self._instrument.query("*IDN?").strip()
            idn_parts = idn_response.split(',')

            # Extract identification components
            manufacturer = idn_parts[0] if len(idn_parts) > 0 else "Unknown"
            model = idn_parts[1] if len(idn_parts) > 1 else "Unknown"
            serial_number = idn_parts[2] if len(idn_parts) > 2 else "Unknown"
            firmware_version = idn_parts[3] if len(idn_parts) > 3 else "Unknown"

            # Get channel status information
            channel_status = {}
            for channel in range(1, self.max_channels + 1):
                try:
                    self._instrument.write(f":INSTrument:SELect CH{channel}")

                    voltage_setpoint = float(self._instrument.query(":VOLTage?"))
                    current_limit = float(self._instrument.query(":CURRent?"))
                    output_enabled = self._instrument.query(":OUTPut?").strip() in ["1", "ON"]

                    measurements = self.measure_channel_output(channel)
                    if measurements:
                        actual_voltage, actual_current = measurements
                    else:
                        actual_voltage = actual_current = 0.0

                    channel_status[f'channel_{channel}'] = {
                        'voltage_setpoint': voltage_setpoint,
                        'current_limit': current_limit,
                        'output_enabled': output_enabled,
                        'measured_voltage': actual_voltage,
                        'measured_current': actual_current
                    }
                except Exception as e:
                    self._logger.warning(f"Failed to get status for channel {channel}: {e}")

            # Compile comprehensive instrument information
            info = {
                'manufacturer': manufacturer,
                'model': model,
                'serial_number': serial_number,
                'firmware_version': firmware_version,
                'visa_address': self._visa_address,
                'connection_status': 'Connected' if self._is_connected else 'Disconnected',
                'max_channels': self.max_channels,
                'max_voltage': self.max_voltage,
                'max_current': self.max_current,
                'channel_status': channel_status
            }

            return info

        except Exception as e:
            self._logger.error(f"Failed to retrieve instrument information: {e}")
            return None

    def _configure_model_parameters(self, model: str) -> None:
        """Configure instrument parameters based on detected model."""
        model_upper = model.upper()

        if "2230" in model_upper:
            self.max_channels = 3
            self.max_voltage = 30.0
            self.max_current = 3.0
        elif "2231" in model_upper:
            self.max_channels = 3
            self.max_voltage = 30.0
            self.max_current = 3.0
        elif "2280" in model_upper:
            self.max_channels = 1
            self.max_voltage = 72.0
            self.max_current = 6.0
        elif "2260" in model_upper or "2268" in model_upper:
            self.max_channels = 1
            self.max_voltage = 600.0
            self.max_current = 10.0
        else:
            self._logger.warning(f"Unknown model {model}, using default parameters")

        self._logger.info(f"Configured for model {model}: {self.max_channels} channels, "
                         f"{self.max_voltage}V/{self.max_current}A max")

    def _initialize_channels_safe_state(self) -> None:
        """Initialize all channels to safe default states."""
        for channel in range(1, self.max_channels + 1):
            try:
                self._instrument.write(f":INSTrument:SELect CH{channel}")
                self._instrument.write(":VOLTage 0.0")      # Set voltage to 0V
                self._instrument.write(":CURRent 0.1")      # Set current limit to 100mA
                self._instrument.write(":OUTPut OFF")       # Ensure output is disabled
            except Exception as e:
                self._logger.warning(f"Failed to initialize channel {channel}: {e}")

    def _disable_channel_output_internal(self, channel: int) -> bool:
        """Internal method to disable channel output."""
        try:
            self._logger.debug(f"Disabling output on channel {channel}")

            # Select the specified channel
            self._instrument.write(f":INSTrument:SELect CH{channel}")

            # Disable the output
            self._instrument.write(":OUTPut OFF")

            # Brief wait for command to take effect
            time.sleep(0.1)

            # Verify output is disabled
            output_state = self._instrument.query(":OUTPut?").strip()
            is_disabled = output_state in ["0", "OFF"]

            if is_disabled:
                # Update configuration if it exists
                if channel in self._channel_configs:
                    self._channel_configs[channel].enabled = False

                self._logger.debug(f"Channel {channel} output disabled successfully")
                return True
            else:
                self._logger.error(f"Failed to disable channel {channel} output")
                return False

        except Exception as e:
            self._logger.error(f"Error disabling channel {channel} output: {e}")
            return False

    def _get_channel_voltage_setpoint(self, channel: int) -> Optional[float]:
        """Get current voltage setpoint for specified channel."""
        try:
            self._instrument.write(f":INSTrument:SELect CH{channel}")
            voltage_str = self._instrument.query(":VOLTage?")
            return float(voltage_str.strip())
        except Exception as e:
            self._logger.debug(f"Failed to read voltage setpoint for channel {channel}: {e}")
            return None

    def _ramp_voltage_safely(self, channel: int, start_voltage: float, target_voltage: float) -> None:
        """Safely ramp voltage from start to target with step limiting."""
        voltage_difference = target_voltage - start_voltage

        # If change is small, apply directly
        if abs(voltage_difference) <= self._max_voltage_step:
            self._instrument.write(f":VOLTage {target_voltage:.6f}")
            time.sleep(self._voltage_settling_time)
            return

        # Calculate number of steps needed
        num_steps = int(abs(voltage_difference) / self._max_voltage_step) + 1
        voltage_step = voltage_difference / num_steps

        # Perform stepped voltage change
        for step in range(1, num_steps + 1):
            intermediate_voltage = start_voltage + (voltage_step * step)
            self._instrument.write(f":VOLTage {intermediate_voltage:.6f}")
            time.sleep(self._voltage_settling_time)

    def _cleanup_connection(self) -> None:
        """Clean up connection state and references."""
        self._is_connected = False
        self._instrument = None
        self._resource_manager = None
        self._channel_configs.clear()

    @property
    def is_connected(self) -> bool:
        """Check if power supply is currently connected."""
        return self._is_connected

    @property
    def visa_address(self) -> str:
        """Get the VISA address for this instrument."""
        return self._visa_address

    @property
    def channel_configurations(self) -> Dict[int, ChannelConfiguration]:
        """Get current channel configurations."""
        return self._channel_configs.copy()


def main() -> None:
    """Example usage demonstration."""
    # Configuration parameters
    power_supply_address = "USB0::0x05E6::0x2230::9103456::INSTR"

    # Create power supply instance
    psu = KeithleyPowerSupply(power_supply_address)

    try:
        # Connect to instrument
        if not psu.connect():
            print("Failed to connect to power supply")
            return

        print("Connected to power supply successfully")

        # Configure channel 1
        success = psu.configure_channel(
            channel=1,
            voltage=5.0,        # 5V output
            current_limit=0.5,  # 500mA current limit
            ovp_level=6.0,      # 6V over-voltage protection
            enable_output=True  # Enable output immediately
        )

        if success:
            print("Channel 1 configured and enabled successfully")

            # Take measurement
            measurements = psu.measure_channel_output(channel=1)
            if measurements:
                voltage, current = measurements
                power = voltage * current
                print(f"Channel 1 output: {voltage:.3f}V, {current:.3f}A, {power:.3f}W")
            else:
                print("Measurement failed")

        # Display instrument information
        info = psu.get_instrument_info()
        if info:
            print(f"Instrument: {info['manufacturer']} {info['model']}")
            print(f"Serial: {info['serial_number']}")
            print(f"Channels: {info['max_channels']}")
            print(f"Max ratings: {info['max_voltage']}V / {info['max_current']}A")

    except KeithleyPowerSupplyError as e:
        print(f"Power supply error: {e}")

    except Exception as e:
        print(f"Unexpected error: {e}")

    finally:
        # Always disable outputs and disconnect
        if psu.is_connected:
            psu.disable_all_outputs()
        psu.disconnect()
        print("Disconnected from power supply")


if __name__ == "__main__":
    main()
