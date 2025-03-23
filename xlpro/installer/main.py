


"""
Installer steps

install location

default %userprofile%/.xlpro

creates the following structure
bin/
    xlpro-cli.exe
        (responsible for creating virtual environments and maintaining the mappings) 
        -help command
lib/
    (stores app dependencies xlpro and sdks etc)
    uv.exe 
        XLPRO_UV_PATH=lib/uv.exe
        (necessary to download new virtual environments via cli interface)
    net6.0-windows/...
        XLPRO_NET6_0_SDK_PATH=lib/...
        (requires get net6.0-windows function)
    ExcelDna/...
        XLPRO_EXCELDNA_PATH=lib/...
        (necessary for the xll compiler)
        (requires get ExcelDna function)
        (XXX - todo - does this require getting intellisense also?)
        
install.exe
    installer probably just configures the xlpro-cli bin path and adds it to
    installs dependencies
        uv.exe
        net6.0-windows
        exceldna
    configures path variables (xlpro should look for these but default to the local locations)

uninstall.exe
    uninstaller (undoes whatever is done in install)

config.toml
    (configuration for xlpro-cli (none pending atm))
venv_mappings.json
    (stores the virtual environment mappings in a json)
envs/
    (stores the virtual environments)





xlpro python package
    responsible for core xlpro functionality

pip install xlpro



"""