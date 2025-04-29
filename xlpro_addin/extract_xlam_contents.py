import zipfile
from pathlib import Path
import shutil
import olefile
from oletools.olevba import VBA_Parser
import os
from win32com.client import Dispatch
import regex as re
root_dir = Path(__file__).parent

xlpro_xlam_fp = Path(r"C:\Users\Daniel Evans\AppData\Roaming\Microsoft\Excel\XLSTART\xlpro.xlam")
temp_dir = root_dir / "tmp"
src_dir = root_dir / "src"
dst_dir = root_dir / "dist"
xlam_dist_file_name = xlpro_xlam_fp.stem + "_TEST" + xlpro_xlam_fp.suffix
# xlam_dist_file_name = xlpro_xlam_fp.name

custom_ui_xml_name = "customui14.xml"


import typing
if typing.TYPE_CHECKING:
    from win32typelibs import excel as xl
    from win32typelibs import vbide as vb


def regen_src_dir():
    p = src_dir
    try:
        print(f"deleting {p}...")
        shutil.rmtree(p)
    except FileNotFoundError:
        print(f"not found, no need to delete {p}...")
        pass
    p.mkdir()

def regen_temp_dir():
    p = temp_dir
    try:
        print(f"deleting {p}...")
        shutil.rmtree(p)
    except FileNotFoundError:
        print(f"not found, no need to delete {p}...")
        pass
    p.mkdir()


def extract():
    regen_temp_dir()
    regen_src_dir()

    xlapp:'xl._Application' = Dispatch("excel.application")
    xlapp.Visible = True

    wb = xlapp.Workbooks.Open(str(xlpro_xlam_fp))

    # tmp_xlam_fp = temp_dir / "xlpro_tmp.xlam"

    proj = wb.VBProject

    # === EXPORT MODULES ===
    for vbcomp in proj.VBComponents:
        name = vbcomp.Name
        comp_type = vbcomp.Type

        # Determine extension based on type
        if comp_type == 1:   # Standard module
            ext = ".bas"
        elif comp_type == 2: # Class module
            ext = ".cls"
        elif comp_type == 3: # UserForm
            ext = ".frm"
        elif comp_type == 100: # Document (e.g., ThisWorkbook, Sheet1)
            ext = ".cls"
        else:
            ext = ".txt"  # fallback, unknown type

        export_path = os.path.join(src_dir, name + ext)
        vbcomp.Export(export_path)
        print(f"Exported: {name}{ext}")



    print(f"extracting to zip archive")
    with zipfile.ZipFile(xlpro_xlam_fp, "r") as ref:
        archive_fp = temp_dir / xlpro_xlam_fp.name
        ref.extractall(archive_fp)

    # print(f"fetching vba project contents {src_dir}...")
    # # get the xml file contents
    # vba_proj_fp = archive_fp / "xl/vbaproject.bin"
    # vbaparser = VBA_Parser(vba_proj_fp)
    # if vbaparser.detect_vba_macros():
    #     for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    #         with open(fp:=(src_dir / f"{vba_filename}"), "w") as f:
    #             print(f"writing vba code to {fp.name}")
    #             f.write(vba_code)
    # else:
    #     print(f"no vba macros found within {vba_proj_fp.name}")
    # vbaparser.close()

    print(f"fetching the custom ui compoenent")
    # get the custom ui contents
    custom_ui_fp = archive_fp / "customui/customui14.xml"

    shutil.copy2(custom_ui_fp, src_dir / custom_ui_fp.name)

    print("deleting temprary directory")
    shutil.rmtree(temp_dir)

    print("done")


def zip_folder(folder_path, output_zip):
    folder = Path(folder_path)
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in folder.rglob('*'):  # rglob('*') will recursively get all files
            if file.is_file():
                arcname = file.relative_to(folder)  # Preserve folder structure inside the zip
                zipf.write(file, arcname)


def build():
    regen_temp_dir()
    dst_dir.mkdir(exist_ok=True)



    xlapp:'xl._Application' = Dispatch("excel.application")
    xlapp.Visible = True

    wb = xlapp.Workbooks.Add()

    tmp_xlam_fp = temp_dir / "xlpro_tmp.xlam"

    proj = wb.VBProject

    # Insert all your modules
    print(f"loading the vba contents into the workbook")
    vba_files = list(src_dir.glob("*.bas")) + list(src_dir.glob("*.cls"))

    pass

    for file in vba_files:
        if not file.stem.lower() in ["xlpro_static", "xlproeventhandler"]:
            continue
        print(f"adding vba code {file}...")
        # if file.stem == "xlpro_static":
        #     print("adding metadata to ")
        #     with open(file, "r") as f:
        #         contents = f.read()
        #         pattern = r"### BEGIN METADATA ###.*?^### END METADATA ###\n\n"
        #         re.sub(pattern, "", contents, flags=re.DOTALL)
        #         msg_str = (
        #             f"' ### BEGIN METADATA ###\n"
        #             f"' --- {file.name} ---\n"
        #             f"' Compiled with {Path(__file__).relative_to(root_dir.parent.parent)}\n"
        #             f"' At {datetime.datetime.now().strftime("%Y-%m-%d, %H:%M:%S")}\n"
        #             f"' ### END METADATA ###\n\n"
        #         )
        #         replacement = lambda m: m.group(1) + '\n' + msg_str
        #         new_contents = re.sub(r'(Attribute VB_Name = .*?\n)', replacement, contents)
        #         pass
        #     with open(file, "w") as f:
        #         f.write(new_contents)
        #     pass
        proj.VBComponents.Import(str(file.resolve()))
        pass

    wb.SaveAs(str(tmp_xlam_fp.parent / tmp_xlam_fp.stem), 55)
    wb.Close(False)
    xlapp.Quit()
    xlapp = None

    tmp_archive_name = Path() / "xlpro_tmp"

    print(f"extracting xlam to zip archive")
    with zipfile.ZipFile(tmp_xlam_fp, "r") as ref:
        archive_fp = tmp_xlam_fp.parent / tmp_archive_name
        ref.extractall(archive_fp)


    custom_ui_fp = archive_fp / "customui" / custom_ui_xml_name
    # custom_ui_fp.parent.mkdir(parents=True, exist_ok=True)
    shutil.rmtree(custom_ui_fp.parent, custom_ui_fp)
    custom_ui_fp.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src_dir / custom_ui_xml_name, custom_ui_fp)

    os.remove(tmp_xlam_fp)
    zip_folder(archive_fp, tmp_xlam_fp)

    shutil.copy2(tmp_xlam_fp, dst_dir / xlam_dist_file_name)

    print("deleting temprary directory")
    shutil.rmtree(temp_dir)

    print("done")


import datetime
def reload_to_xlstart():
    archive_dir = xlpro_xlam_fp.parent / "archive"
    archive_dir.mkdir(exist_ok=True)
    
    if xlpro_xlam_fp.exists():
        shutil.copy2(xlpro_xlam_fp, archive_dir / f"{xlpro_xlam_fp.stem}_{datetime.datetime.now().strftime("%Y-%m-%d")}{xlpro_xlam_fp.suffix}")
    
    dst_build = dst_dir / ""



import argparse
def main():
    parser = argparse.ArgumentParser(description="Build or extract VBA project files.")
    parser.add_argument('--build', action='store_true', help='Build a VBA project into an .xlam file.')
    parser.add_argument('--extract', action='store_true', help='Extract VBA project files from an .xlam file.')

    args = parser.parse_args()

    if args.build:
        print("Running build process...")
        build()

    if args.extract:
        print("Running extract process...")
        extract()

    if not args.build and not args.extract:
        print("No action provided. Use --build or --extract.")


if __name__ == "__main__":
    extract()
    build()