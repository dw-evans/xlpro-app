from setuptools import setup, find_packages

setup(
    name="xlpro_cli",
    version='0.1.0',
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'xlpcli=xlpro_cli.commands:start_server',
        ],
    },
    author='Daniel Evans',
    description='xlpro command line interface',
    url='tbc',
    install_requires=[],
)
