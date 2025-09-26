# Professional Instrument Control Library

A comprehensive, enterprise-grade Python library for controlling laboratory and test equipment with precision and reliability. Designed for production environments where accuracy, safety, and maintainability are paramount.

## Features

### Multi-Instrument Support
- **Keithley Power Supplies**: 2230 Series, 2231A Series, 2280S Series
- **Keithley Multimeters**: DMM6500, DMM7510 with high-precision measurements
- **Keysight Oscilloscopes**: DSOX6000 Series with advanced triggering

### Professional Quality
- **Industry Standards**: Full IEEE 488.2 and SCPI compliance
- **Enterprise Architecture**: Modular design with comprehensive error handling
- **Production Ready**: Extensive logging, validation, and safety interlocks
- **Maintainable**: Clean code structure with full documentation

### Advanced Capabilities
- **Statistical Analysis**: Automated measurement statistics and data validation
- **Safety Systems**: Over-voltage/current protection and emergency shutdown
- **Real-time Monitoring**: Continuous measurement and status reporting
- **Configuration Management**: Persistent settings and automated discovery

## Quick Start

### Installation

```bash
# Install from PyPI
pip install professional-instrument-control

# Or install from source
git clone https://github.com/example/professional-instrument-control.git
cd professional-instrument-control
pip install -r requirements.txt
python setup.py install
```

### Basic Usage

```python
from instrument_control.keithley_power_supply import KeithleyPowerSupply
from instrument_control.keithley_dmm import KeithleyDMM6500

# Initialize instruments
power_supply = KeithleyPowerSupply('USB0::0x05E6::0x2230::9103456::INSTR')
multimeter = KeithleyDMM6500('USB0::0x05E6::0x6500::04561287::INSTR')

# Connect to instruments
power_supply.connect()
multimeter.connect()

# Configure power supply
power_supply.configure_channel(channel=1, voltage=5.0, current_limit=1.0)
power_supply.enable_channel_output(channel=1)

# Perform high-precision measurement
voltage = multimeter.measure_dc_voltage(
    measurement_range=10.0,
    resolution=1e-6,  # 1µV resolution
    nplc=10.0        # High accuracy mode
)

print(f"Measured voltage: {voltage:.9f}V")

# Safe shutdown
power_supply.disable_all_outputs()
power_supply.disconnect()
multimeter.disconnect()
```

### Complete Automation System

```bash
# Run the interactive automation application
python instrument_automation_system.py
```

## Architecture

### Library Structure
```
instrument_control/
├── __init__.py
├── keithley_power_supply.py    # Multi-channel PSU control
├── keithley_dmm.py            # High-precision measurements
├── keysight_oscilloscope.py   # Waveform capture and analysis
└── common/
    ├── exceptions.py           # Custom exception classes
    ├── validators.py           # Parameter validation utilities
    └── logging_config.py       # Professional logging setup
```

### Key Design Principles

1. **Safety First**: All operations include safety checks and interlocks
2. **Error Resilience**: Comprehensive exception handling with recovery procedures
3. **SCPI Compliance**: Full adherence to industry communication standards
4. **Modular Design**: Independent modules with clean interfaces
5. **Professional Logging**: Detailed operation tracking for debugging and audit

## Supported Instruments

### Keithley Power Supplies
- **2230-30-3**: Triple output, 30V/3A per channel
- **2231A-30-3**: Enhanced triple output with data logging
- **2280S Series**: Single output, high current applications
- **2260B/2268**: Programmable DC loads

### Keithley Multimeters  
- **DMM6500**: 6.5-digit precision with 1µV resolution
- **DMM7510**: 7.5-digit graphical sampling multimeter
- Statistical analysis with up to 10 NPLC integration
- Auto-zero correction for maximum accuracy

### Keysight Oscilloscopes
- **DSOX6004A**: 4-channel, 1 GHz bandwidth, 20 GSa/s
- Advanced triggering and measurement capabilities
- Waveform capture and analysis functions

## Advanced Features

### Statistical Measurements
```python
# Perform automated statistical analysis
stats = multimeter.perform_measurement_statistics(
    measurement_count=20,
    measurement_interval=0.1
)

print(f"Mean: {stats['mean']:.9f}V")
print(f"Std Dev: {stats['standard_deviation']:.9f}V") 
print(f"Coefficient of Variation: {stats['coefficient_of_variation_percent']:.3f}%")
```

### Safety Interlocks
```python
# Configure over-voltage protection
power_supply.configure_channel(
    channel=1,
    voltage=5.0,
    current_limit=1.0,
    ovp_level=6.0  # Trip if voltage exceeds 6V
)

# Emergency shutdown
power_supply.disable_all_outputs()  # Immediate shutdown of all channels
```

### Professional Logging
```python
import logging

# Configure comprehensive logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('instrument_control.log'),
        logging.StreamHandler()
    ]
)

# All instrument operations are automatically logged
```

## Error Handling

The library implements comprehensive error handling with specific exception types:

```python
from instrument_control.keithley_dmm import KeithleyDMM6500Error

try:
    voltage = multimeter.measure_dc_voltage(measurement_range=10.0)
except KeithleyDMM6500Error as e:
    print(f"Multimeter error: {e}")
    # Handle specific instrument error
except Exception as e:
    print(f"Unexpected error: {e}")
    # Handle general error
```

## Configuration Management

### Instrument Discovery
```python
from instrument_control.common.discovery import discover_instruments

# Automatic instrument detection
instruments = discover_instruments()
for instrument_type, address in instruments.items():
    print(f"Found {instrument_type}: {address}")
```

### Configuration Persistence
```python
# Save instrument configuration
config = {
    'power_supply': {
        'address': 'USB0::0x05E6::0x2230::9103456::INSTR',
        'channels': {
            1: {'voltage': 5.0, 'current_limit': 1.0},
            2: {'voltage': 3.3, 'current_limit': 0.5}
        }
    }
}

# Configuration automatically loaded on next session
```

## Testing and Validation

### Unit Tests
```bash
# Run comprehensive test suite
pytest tests/ -v --cov=instrument_control

# Run specific instrument tests
pytest tests/test_keithley_dmm.py -v
```

### Integration Tests
```bash
# Test with actual hardware (requires connected instruments)
pytest tests/integration/ -v --hardware
```

### Performance Benchmarks
```bash
# Measure measurement performance
python benchmarks/measurement_speed.py
```

## Development Guidelines

### Code Quality Standards
- **PEP 8**: Full compliance with Python style guidelines
- **Type Hints**: Complete type annotations for all public APIs
- **Documentation**: Comprehensive docstrings in NumPy format
- **Testing**: 95%+ code coverage with unit and integration tests

### Contributing
1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Run tests: `pytest tests/ -v`
4. Ensure code quality: `black . && pylint instrument_control/`
5. Submit a pull request

## Performance Considerations

### Measurement Speed
- **Fast measurements**: ~10ms per measurement
- **High precision**: ~100ms per measurement (10 NPLC)
- **Statistical analysis**: Configurable measurement intervals

### Memory Usage
- **Minimal footprint**: <10MB typical usage
- **Efficient buffering**: Optimized VISA communication
- **Resource cleanup**: Automatic connection management

### Communication Reliability
- **Timeout handling**: Configurable timeouts per instrument type
- **Error recovery**: Automatic retry with exponential backoff
- **Connection monitoring**: Health checks and reconnection

## Troubleshooting

### Common Issues

**Instrument not detected**
```bash
# Check VISA installation
python -c "import pyvisa; rm = pyvisa.ResourceManager(); print(rm.list_resources())"

# Verify USB drivers
# Update NI-VISA or install pyvisa-py backend
```

**Communication timeouts**
```python
# Increase timeout for precision measurements
multimeter = KeithleyDMM6500(address, timeout_ms=60000)  # 60 second timeout
```

**Permission errors**
```bash
# On Linux, add user to dialout group
sudo usermod -a -G dialout $USER
# Logout and login again
```

### Debug Logging
```python
# Enable detailed debug logging
logging.getLogger('instrument_control').setLevel(logging.DEBUG)
```

## License

MIT License - see LICENSE file for details.

## Support

- **Documentation**: https://professional-instrument-control.readthedocs.io/
- **Issues**: https://github.com/example/professional-instrument-control/issues
- **Discussions**: https://github.com/example/professional-instrument-control/discussions

## Changelog

### Version 1.0.0
- Initial release with full instrument support
- Professional-grade architecture and error handling
- Comprehensive test suite and documentation
- Statistical analysis capabilities
- Safety interlocks and emergency shutdown procedures
