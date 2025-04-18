# xlpro Application

## xlpro_module

- Python module which manages python server creation.
- Wheel file is created for internal distribution to the installer

## xlpro_cli

- Command line tool which manages workbook and virtual environment configuration
- Generates data within the installation directory.


## xlpro_installer

- Installer for distribution
- Includes the following:
  - uv.exe (indirectly fetched)
  - xlpro.xlam (xlpro-add in file, copy moved to xlstart)
  - xlpro*.whl (wheel file for building source for each virtual environment)
  - xlpro-cli.exe (cli tool added to path to work with xlpro.xlam)
  - envs/ (folder for storing independent virtual environments)
  - *configuration files and other*