from setuptools import setup, find_packages

import sys
from pathlib import Path

# import xlpro

pass
setup(
    name="xlpro_cli",
    # version=xlpro.__version__,
    version="0.0.4",
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'xlpro-cli=xlpro_cli.commands:main',
        ],
    },
    author='Daniel Evans',
    description='xlpro Command Line Interface',
    url='N/A',
    install_requires=[],
)

