
"""
# commands for syncing the uv environment to
python --version
pip freeze
uv resolve --lockfile C:\path\to\your\project\uv.lock
uv pip sync --lockfile C:\path\to\your\project\uv.lock

# installing uv globally
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
curl -LsSf https://astral.sh/uv/install.sh | sh
pip install uv

# check if not uv.exe exists before installing it?


xlpro sync /path/to/directory
  checks python version compatibility
  warns if version mismatch > do you want to install this version
  checks for dependencies in the current python environment
  syncs the python environment if the user confirms.
  presumably within this it will check for incompatibility

  pipe the uv stdout to the console window
  
  all of this should be spawned from a CLI coming from an xlpro add-in button

  
pip install xlpro
  installs xlpro globally, access to the cli
  xlpro init /path/to/book1.xlsx -> initialises the .xlpro directory next to a workbook. folder could be book1.xlpro.
  xlpro sync /path/to/book1.xlpro/ -> creates a new virtual environment using the current python version, syncs with the pyproject.toml
  sync
    when we link a venv to an excel file, we should store the link in a dict
  atm we support multiple workbooks from the same virtual environment
  

file structure

c:/users/daniel.evans/

  .xlpro/

    venv-mappings.json -> maps the filenames to the virtual environment. Now if there is a live duplicate, xlpro can handle it
      "c/users/daniel.evans/documents/book1.xlsx": "../envs/2025-03-10_a2h3g5j123hj12h3_book1xlsx/.venv"
      "c/users/daniel.evans/documents/book2.xlsx": "../envs/2025-03-13_h35kj123j5k521k4_book2xlsx/.venv"

    envs/ -> stores all xlpro registered environments
      2025-03-10_a2h3g5j123hj12h3_book1xlsx/
        .venv/
          lib
          scripts
          ...
        .xlpro/ -> stores xlpro metadata
          xlpro.lock -> tells us if this venv is already running xlpro
          xlpro.log -> logging information for the session
          xlpro.meta -> stores metadata, version, sync time sync user
          pyproject.toml or requirements.txt
          .python-version
        
      2025-03-13-h35kj123j5k521k4_book2xlsx/
        .venv/ -> junction symlink to the live venv can work
        .xlpro/

    xlpro.exe -> runs xlpro.run_server.py within a venv using the configured virtual environment and registers the lockfile

  
other functions needed

  xlpro init
    configures the workbook for xlpro usage
    creates or selects a python environment
    uses global xlpro to initialize the environment
    uses global uv installation to create the environment via xlpro. Thinking about it now, uv should just be a dependency within xlpro
      maybe hide it behind > xlpro uv ... similar to 
    installs the specific xlpro version within the virtual environment
    

  xlpro sync
    syncs the python virtual environment for the specified fname.xlpro/pyproject.toml and .python-version


      
  
  
  
  
  

"""