from pathlib import Path
import re
import xlpro
import pythoncom


wd = Path(__file__).parent

fp_inno = wd / "build_installer.iss"


with open(fp_inno, "r") as f:
    txt = f.read()

txt2 = txt

version = xlpro.__version__

import json

with open(wd / "app_id_map.json", "r") as f:
    tmp = f.read()
    app_id_map = json.loads(tmp)

if not version in app_id_map:
    with open(wd / "app_id_map.json", "w") as f:
        app_id_map = app_id_map | {version: str(pythoncom.CreateGuid())}
        f.write(json.dumps(app_id_map, indent=4))

else:
    version_guid = app_id_map[version]

pat = r'^#define MyAppVersion\s*".*"\s*$'
x1a = re.search(pat, txt2, flags=re.MULTILINE)
if not re.search(pat, txt2, flags=re.MULTILINE):
    raise Exception
txt2 = re.sub(pat, f"#define MyAppVersion \"{version}\"", txt2, flags=re.MULTILINE)
x1 = re.search(pat, txt2, flags=re.MULTILINE)

pat = "^AppId=.*$"
x2a = re.search(pat, txt2, flags=re.MULTILINE)
if not re.search(pat, txt2, flags=re.MULTILINE):
    raise Exception
txt2 = re.sub(pat, "AppId={" + version_guid, txt2, flags=re.MULTILINE)
x2 = re.search(pat, txt2, flags=re.MULTILINE)


with open(fp_inno, "w") as f:
    f.write(txt2)