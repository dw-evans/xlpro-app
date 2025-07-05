from setuptools import setup, find_packages

import sys
from pathlib import Path

setup(
    name="xlpro_cli",
    version="0.0.0",
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'xlpro-cli=xlpro_cli.commands:main',
        ],
    },
    author='Daniel Evans',
    description='xlpro command line interface package',
    url='N/A',
    install_requires=[],
)

