import logging
import time
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any, List, Tuple, Union
import numpy as np
from .scpi_wrapper import SCPIWrapper

class KeysightDSOX6004AError(Exception):
    """Custom exception for Keysight DSOX6004A oscilloscope errors."""
    pass

class KeysightDSOX6004A:
    def __init__(self, visa_address: str, timeout_ms: int = 10000) -> None:
        self._scpi_wrapper = SCPIWrapper(visa_address, timeout_ms)
        self._logger = logging.getLogger(f'{self.__class__.__name__}.{id(self)}')
        self.max_channels = 4
        self.max_sample_rate = 20e9
        self.max_memory_depth = 16e6
        self.bandwidth_hz = 1e9
        self._valid_vertical_scales = [
            1e-3, 2e-3, 5e-3, 10e-3, 20e-3, 50e-3,
            100e-3, 200e-3, 500e-3, 1.0, 2.0, 5.0, 10.0
        ]
        self._valid_timebase_scales = [
            1e-12, 2e-12, 5e-12, 10e-12, 20e-12, 50e-12,
            100e-12, 200e-12, 500e-12, 1e-9, 2e-9, 5e-9,
            10e-9, 20e-9, 50e-9, 100e-9, 200e-9, 500e-9,
            1e-6, 2e-6, 5e-6, 10e-6, 20e-6, 50e-6,
            100e-6, 200e-6, 500e-6, 1e-3, 2e-3, 5e-3,
            10e-3, 20e-3, 50e-3, 100e-3, 200e-3, 500e-3,
            1.0, 2.0, 5.0, 10.0, 20.0, 50.0
        ]

    def connect(self) -> bool:
        if self._scpi_wrapper.connect():
            try:
                identification = self._scpi_wrapper.query("*IDN?")
                self._logger.info(f"Instrument identification: {identification.strip()}")
                if "KEYSIGHT" not in identification.upper() and "AGILENT" not in identification.upper():
                    self._logger.warning(f"Unexpected manufacturer in IDN response: {identification}")
                if "DSOX6004A" not in identification.upper():
                    self._logger.warning(f"Unexpected model in IDN response: {identification}")
                self._scpi_wrapper.write("*CLS")
                time.sleep(0.5)
                self._scpi_wrapper.query("*OPC?")
                self._logger.info("Successfully connected to Keysight DSOX6004A")
                return True
            except Exception as e:
                self._logger.error(f"Error during instrument identification: {e}")
                self._scpi_wrapper.disconnect()
                return False
        return False

    def disconnect(self) -> None:
        self._scpi_wrapper.disconnect()
        self._logger.info("Disconnection completed")

    @property
    def is_connected(self) -> bool:
        return self._scpi_wrapper.is_connected

    def get_instrument_info(self) -> Optional[Dict[str, Any]]:
        """Get comprehensive instrument information."""
        if not self.is_connected:
            return None
        
        try:
            idn = self._scpi_wrapper.query("*IDN?").strip()
            parts = idn.split(',')
            
            return {
                'manufacturer': parts[0] if len(parts) > 0 else 'Unknown',
                'model': parts[1] if len(parts) > 1 else 'Unknown',
                'serial_number': parts[2] if len(parts) > 2 else 'Unknown',
                'firmware_version': parts[3] if len(parts) > 3 else 'Unknown',
                'max_channels': self.max_channels,
                'bandwidth_hz': self.bandwidth_hz,
                'max_sample_rate': self.max_sample_rate,
                'max_memory_depth': self.max_memory_depth,
                'identification': idn
            }
        except Exception as e:
            self._logger.error(f"Failed to get instrument info: {e}")
            return None

    def configure_channel(self, channel: int, vertical_scale: float, vertical_offset: float = 0.0, coupling: str = "DC", probe_attenuation: float = 1.0) -> bool:
        if not self.is_connected:
            raise KeysightDSOX6004AError("Oscilloscope not connected")
        if not (1 <= channel <= self.max_channels):
            raise ValueError(f"Channel must be 1-{self.max_channels}, got {channel}")
        if vertical_scale not in self._valid_vertical_scales:
            closest_scale = min(self._valid_vertical_scales, key=lambda x: abs(x - vertical_scale))
            self._logger.warning(f"Invalid vertical scale {vertical_scale}, using {closest_scale}")
            vertical_scale = closest_scale
        valid_coupling = ["AC", "DC", "GND"]
        if coupling.upper() not in valid_coupling:
            raise ValueError(f"Coupling must be one of {valid_coupling}, got {coupling}")
        try:
            self._scpi_wrapper.write(f":CHANnel{channel}:DISPlay ON")
            self._scpi_wrapper.write(f":CHANnel{channel}:SCALe {vertical_scale}")
            self._scpi_wrapper.write(f":CHANnel{channel}:OFFSet {vertical_offset}")
            self._scpi_wrapper.write(f":CHANnel{channel}:COUPling {coupling}")
            self._scpi_wrapper.write(f":CHANnel{channel}:PROBe {probe_attenuation}")
            actual_scale = float(self._scpi_wrapper.query(f":CHANnel{channel}:SCALe?"))
            actual_offset = float(self._scpi_wrapper.query(f":CHANnel{channel}:OFFSet?"))
            self._logger.info(f"Channel {channel} configured: Scale={actual_scale}V/div, Offset={actual_offset}V, Coupling={coupling}, Probe={probe_attenuation}x")
            return True
        except Exception as e:
            self._logger.error(f"Failed to configure channel {channel}: {e}")
            return False

    def capture_screenshot(self, filename: Optional[str] = None, image_format: str = "PNG", include_timestamp: bool = True) -> Optional[str]:
        if not self.is_connected:
            self._logger.error("Cannot capture screenshot: not connected")
            return None
        try:
            self.setup_output_directories()
            if filename is None:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
                if include_timestamp:
                    filename = f"scope_screenshot_{timestamp}.{image_format.lower()}"
                else:
                    filename = f"scope_screenshot.{image_format.lower()}"
            if not filename.lower().endswith(f".{image_format.lower()}"):
                filename += f".{image_format.lower()}"
            screenshot_path = self.screenshot_dir / filename
            self._logger.info(f"Capturing screenshot in {image_format} format...")
            self._scpi_wrapper.write(":HARDcopy:DESTination FILE")
            time.sleep(0.2)
            self._scpi_wrapper.write(f":HARDcopy:FORMat {image_format}")
            time.sleep(0.2)
            image_data = self._scpi_wrapper.query_binary_values(f":DISPlay:DATA? {image_format}", datatype='B')
            if image_data:
                with open(screenshot_path, 'wb') as f:
                    f.write(bytes(image_data))
                self._logger.info(f"Screenshot saved: {screenshot_path}")
                return str(screenshot_path)
            return None
        except Exception as e:
            self._logger.error(f"Screenshot capture failed: {e}")
            return None

    def setup_output_directories(self) -> None:
        base_path = Path.cwd()
        self.screenshot_dir = base_path / "oscilloscope_screenshots"
        self.data_dir = base_path / "oscilloscope_data"
        self.graph_dir = base_path / "oscilloscope_graphs"
        for directory in [self.screenshot_dir, self.data_dir, self.graph_dir]:
            directory.mkdir(exist_ok=True)
