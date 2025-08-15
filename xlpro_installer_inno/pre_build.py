from pathlib import Path
import re
import xlpro
import pythoncom


wd = Path(__file__).parent

fp_inno = wd / "build_installer.iss"


with open(fp_inno, "r") as f:
    txt = f.read()

txt2 = txt

pat = r'^#define MyAppVersion\s*".*"\s*$'
x1a = re.search(pat, txt2, flags=re.MULTILINE)
if not re.search(pat, txt2, flags=re.MULTILINE):
    raise Exception
txt2 = re.sub(pat, f"#define MyAppVersion \"{xlpro.__version__}\"", txt2, flags=re.MULTILINE)
x1 = re.search(pat, txt2, flags=re.MULTILINE)

pat = "^AppId=.*$"
x2a = re.search(pat, txt2, flags=re.MULTILINE)
if not re.search(pat, txt2, flags=re.MULTILINE):
    raise Exception
txt2 = re.sub(pat, "AppId={" + str(pythoncom.CreateGuid()), txt2, flags=re.MULTILINE)
x2 = re.search(pat, txt2, flags=re.MULTILINE)


with open(fp_inno, "w") as f:
    f.write(txt2)