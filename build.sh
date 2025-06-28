#!/bin/bash
# Install build dependencies
pip install cython numpy

# Clone pandas with specific fixes
git clone https://github.com/pandas-dev/pandas.git
cd pandas
git checkout tags/v2.1.3  # Use a known stable version

# Apply any necessary patches for Python 3.13
# (Check pandas GitHub issues for known fixes)

# Build and install
python setup.py install