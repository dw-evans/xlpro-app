<p align="center">
  <img src="./assets/banner.png" alt="Project Banner" />
</p>

# xlpro

xlpro brings local Python to Excel. Develop workbooks leveraging Python without limits - from macros, to custom functions and dynamic charts. Write Python code which responds to changes to your Excel Workbooks

Find the installation guide and project documentation at [xlpro.pages.dev](http://xlpro.pages.dev).

This project is currently under development. 

The project is MIT-licensed, however the code has not yet been made public. Most of the core engine is source-available if you inspect the bundled `xlpro` Python package, and `xlpro.xlam` Add-In file within the application.

Latest version: v0.0.9.


## Release Notes

Pending. Release notes will be published on the first public release.


---

## Developer Features

### How does this work?

By creating a lightweight Python calculation server which can receive commands from Excel, with a bit of back and forth, we can register Python functions in Excel as user-defined-functions (UDFs) or macros. Through type conversion when communicating between Excel and Python, we can get Excel's array data into a Python-friendly format, and Python data into an Excel format.

This allows the user to create a lightweight set of scripting files alongside their workbook. This set of files includes the infromation needed to rebuild the environment for sharing between collaborators. Using `uv`, the environments can be built extremely quickly 

- A VBA-enabled Excel Add-in file is generated from the `xlpro_addin folder` from VBA source code.
- This add-in includes code which communicates with `xlpro-cli` via the ribbon. This interaction ultimately runs an independent Python calculation server for the active Workbook.
- `xlpro-cli` is the main desktop application. It is responsible for managing venv creation and VSCode settings management. This is called from the Excel Add-in to do things such as:
  - Initialize a new virtual environment.
  - Start the Python calculation server.
  - Open VSCode and configure it for debugging.


### Build Process

1. Use `XLPRO_BUILD_ALL` VSCode command. Python 3.10, Inno Setup and Go are required dependencies.
2. Tag release
3. Create new release, and tag
4. Upload installer binary on GitHub.
5. Push new version link and change notes to to primary site [xlpro.pages.dev](xlpro.pages.dev).


### TODOs

- Split xlpro-app and xlpro Python module repositories
- Remove xlpro-releases module
- Push xlpro to Pypi
- Handle xlpro and xlpro-app version compatibilit
- Figures do not update if equivalent inputs are generated
- Add win32typelibs support.
- Automated Git integrations for version control safety

### Bugs

- Review function name security risk.