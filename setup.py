#!/usr/bin/env python3
"""
Professional Instrument Control Library Setup Script

This setup script configures the professional instrument control library
for installation and distribution.

Author: Professional Instrument Control Team
Version: 1.0.0
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file for long description
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding='utf-8') if readme_file.exists() else ""

# Read requirements from requirements.txt
requirements_file = Path(__file__).parent / "requirements.txt"
if requirements_file.exists():
    requirements = []
    for line in requirements_file.read_text().strip().split('\n'):
        line = line.strip()
        if line and not line.startswith('#') and not line.startswith('-'):
            # Remove version comments and extract package name
            package = line.split('#')[0].strip()
            if package:
                requirements.append(package)
else:
    requirements = [
        'pyvisa>=1.13.0',
        'pyvisa-py>=0.5.3',
        'numpy>=1.21.0'
    ]

setup(
    name="professional-instrument-control",
    version="1.0.0",
    author="Professional Instrument Control Team",
    author_email="support@example.com",
    description="Professional-grade instrument control library for laboratory automation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/example/professional-instrument-control",

    packages=find_packages(include=['instrument_control', 'instrument_control.*']),

    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Science/Research",
        "Intended Audience :: Manufacturing",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Scientific/Engineering",
        "Topic :: Scientific/Engineering :: Interface Engine/Protocol Translator",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: System :: Hardware :: Hardware Drivers",
    ],

    python_requires=">=3.8",
    install_requires=requirements,

    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "black>=22.0.0",
            "pylint>=2.12.0",
            "mypy>=0.910",
            "pytest-cov>=4.0.0",
        ],
        "analysis": [
            "scipy>=1.7.0",
            "pandas>=1.3.0",
            "matplotlib>=3.5.0",
        ],
    },

    entry_points={
        "console_scripts": [
            "instrument-automation=instrument_automation_system:main",
        ],
    },

    include_package_data=True,
    zip_safe=False,

    project_urls={
        "Bug Reports": "https://github.com/example/professional-instrument-control/issues",
        "Source": "https://github.com/example/professional-instrument-control",
        "Documentation": "https://professional-instrument-control.readthedocs.io/",
    },
)
