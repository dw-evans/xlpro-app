
import xlpro
from pathlib import Path

wd = Path(__file__).parent

# Short script to pre-build the commandline header for the binary

def write_header_txt():
    s = f"""

        ██╗  ██╗██╗     ██████╗ ██████╗  ██████╗ 
        ╚██╗██╔╝██║     ██╔══██╗██╔══██╗██╔═══██╗
         ╚███╔╝ ██║     ██████╔╝██████╔╝██║   ██║
         ██╔██╗ ██║     ██╔═══╝ ██╔══██╗██║   ██║
        ██╔╝ ██╗███████╗██║     ██║  ██║╚██████╔╝
        ╚═╝  ╚═╝╚══════╝╚═╝     ╚═╝  ╚═╝ ╚═════╝ 

          xlpro v{xlpro.__version__}
          Copyright (c) 2025 Daniel Evans
          License: MIT. Free for commercial use.

    """
    with open(wd / "cli-header.txt", "w", encoding="utf-8") as f:
        f.write(s)

def sync_version_file():
    wd = Path(__file__).parent

    # Get version from your_module/version.py
    # version_file = (wd / "xlpro/__init__.py").read_text()
    # version = re.search(r"__version__\s*=\s*['\"](.+)['\"]", version_file).group(1)

    import xlpro
    version = xlpro.__version__

    # Replace version in pyproject.toml
    pyproject = wd / "pyproject.toml"
    py_lines = pyproject.read_text().splitlines()
    new_lines = []
    for line in py_lines:
        if line.strip().startswith("version ="):
            line = f'version = "{version}" # Placeholder will be auto-synced'
        new_lines.append(line)
    pyproject.write_text("\n".join(new_lines) + "\n")

    print(f"✅ Synced version {version} into pyproject.toml")

def main():
    write_header_txt()
    sync_version_file()

if __name__ == "__main__":
    main()