
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

def main():
    write_header_txt()

if __name__ == "__main__":
    main()