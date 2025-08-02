import zipfile
from pathlib import Path
import shutil
import os
from win32com.client.dynamic import Dispatch

root_dir = Path(__file__).parent

xlpro_xlam_fp = Path(r"C:\Users\Daniel Evans\AppData\Roaming\Microsoft\Excel\XLSTART\xlpro.xlam")
temp_dir = root_dir / "tmp"
src_dir = root_dir / "src"
dst_dir = root_dir / "dist"
xlam_dist_file_name = xlpro_xlam_fp.stem + xlpro_xlam_fp.suffix
# xlam_dist_file_name = xlpro_xlam_fp.name

# custom_ui_xml_name = "customui14.xml"

custom_ui_folder_name = "customUI"
rels_folder_name = "_rels"
xlstart_archive_dirname = "archive"
contenttypes_xml_name = "[Content_Types].xml"


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

    vba_dst_dir = src_dir / "vba"
    vba_dst_dir.mkdir()

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

        export_path = vba_dst_dir / f"{name}{ext}"
        vbcomp.Export(str(export_path.resolve()))
        print(f"Exported: {name}{ext}")


    print(f"extracting to zip archive")
    with zipfile.ZipFile(xlpro_xlam_fp, "r") as ref:
        archive_fp = temp_dir / xlpro_xlam_fp.name
        ref.extractall(archive_fp)

    print(f"fetching the custom ui folder from archive")
    custom_ui_fp = archive_fp / custom_ui_folder_name
    shutil.copytree(custom_ui_fp, src_dir / custom_ui_fp.name)



    print(f"fetching the rels folder from archive")
    rels_fp = archive_fp / "_rels"
    shutil.copytree(rels_fp, src_dir / rels_fp.name)

    # cleanup
    print("deleting temporary directory")
    shutil.rmtree(temp_dir)

    print("done")


def zip_folder(folder_path, output_zip):
    folder = Path(folder_path)
    # with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
    with zipfile.ZipFile(output_zip, 'w') as zipf:
        for file in folder.rglob('*'):  # rglob('*') will recursively get all files
            if file.is_file():
                arcname = file.relative_to(folder)  # Preserve folder structure inside the zip
                zipf.write(file, arcname)

def unzip_folder(src, dst):
    with zipfile.ZipFile(src, 'r') as zipf:
        zipf.extractall(dst)


def build():
    regen_temp_dir()
    dst_dir.mkdir(exist_ok=True)

    xlapp:'xl._Application' = Dispatch("excel.application")
    xlapp.Visible = True

    wb = xlapp.Workbooks.Add()


    tmp_xlam_fp = temp_dir / "xlpro_tmp.xlam"

    proj = wb.VBProject

    # Insert all your modules
    vba_src_dir = src_dir / "vba"
    print(f"loading the vba contents into the workbook")

    for comp in wb.VBProject.VBComponents:
        if comp.Name == "ThisWorkbook":
            with open(vba_src_dir / "ThisWorkbook.cls") as f:
                comp.CodeModule.AddFromString(f.read())
                pass
    
    # import the user form
    wb.VBProject.VBComponents.Import(str((vba_src_dir / "UserForm1.frm").resolve()))
    
    pass

    vba_files = list(vba_src_dir.glob("*.bas")) + list(vba_src_dir.glob("*.cls"))

    valid_file_stems = [
        "xlpro_static",
        "xlproeventhandler",
        "undomanager",
        "modundoentrypoint",
        "cworkbookeventhandler",
        "clsundoitem",
        "modutils",
    ]

    for file in vba_files:
        if not file.stem.lower() in valid_file_stems:
            continue
        print(f"adding vba code {file}...")
        with open(file, "r") as f:
            contents = f.read()
            new_contents = contents
            # pattern = r"### BEGIN METADATA ###.*?^### END METADATA ###\n\n"
            # new_contents = re.sub(pattern, "", contents, flags=re.DOTALL)
            # msg_str = (
            #     f"' ### BEGIN METADATA ###\n"
            #     f"' --- {file.name} ---\n"
            #     f"' Compiled with {Path(__file__).relative_to(root_dir.parent.parent)}\n"
            #     f"' At {datetime.datetime.now().strftime("%Y-%m-%d, %H:%M:%S")}\n"
            #     f"' ### END METADATA ###\n\n"
            # )
            # replacement = lambda m: m.group(1) + '\n' + msg_str
            # new_contents = re.sub(r'(Attribute VB_Name = .*?\n)', replacement, new_contents)
            pass
        with open(file, "w") as f:
            f.write(new_contents)
        pass
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

    # custom_ui_fp = archive_fp / "customUI" / custom_ui_xml_name
    # # custom_ui_fp.parent.mkdir(parents=True, exist_ok=True)
    # shutil.rmtree(custom_ui_fp.parent, custom_ui_fp)
    # custom_ui_fp.parent.mkdir(parents=True, exist_ok=True)
    # shutil.copy2(src_dir / custom_ui_xml_name, custom_ui_fp)

    def replace_folder(src, dst):
        try:
            shutil.rmtree(dst)
        except FileNotFoundError:
            pass
        shutil.copytree(src, dst)

    # remove the tree in the archive
    # and copy the source customUI directory into the zip file.
    archive_custom_ui_dir = archive_fp / custom_ui_folder_name
    replace_folder(src_dir / custom_ui_folder_name, archive_custom_ui_dir)

    # copy the content types into the archive
    os.remove(archive_fp / contenttypes_xml_name)
    shutil.copy2(src_dir / contenttypes_xml_name, archive_fp / contenttypes_xml_name)

    # replace top level rels
    rels_dir = archive_fp / rels_folder_name
    replace_folder(src_dir / rels_folder_name, rels_dir)


    pass
    # remove the temporary zipped file
    os.remove(tmp_xlam_fp)
    # zip the archive and conver to xlam extension
    zip_folder(archive_fp, tmp_xlam_fp)

    # copy the xlam to the distribution folder
    shutil.copy2(tmp_xlam_fp, dst_dir / xlam_dist_file_name)

    # cleanup
    print("deleting temporary directory")
    shutil.rmtree(temp_dir)
    print("done")


import datetime
def reload_to_xlstart():
    print(f"initializing archive directory in xlstart")
    archive_dir = xlpro_xlam_fp.parent / xlstart_archive_dirname
    archive_dir.mkdir(exist_ok=True)
    
    src_build = dst_dir / xlam_dist_file_name
    dst_file = xlpro_xlam_fp.parent / src_build.name

    # create the archived version just in case...
    print(f"archiving existing version")
    if dst_file.exists():
        print(f"existing file found, archiving")
        shutil.copy2(dst_file, archive_dir / f"{dst_file.stem}_{datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")}{dst_file.suffix}")
        os.remove(dst_file)
    
    print(f"moving xlpro to xlstart")
    shutil.copy2(src_build, dst_file)



import argparse
def main():
    parser = argparse.ArgumentParser(description="Build or extract VBA project files.")
    parser.add_argument('--build', action='store_true', help='Build a VBA project into an .xlam file.')
    parser.add_argument('--push', action='store_true', help='Build a VBA project into an .xlam file.')
    # parser.add_argument('--extract', action='store_true', help='Extract VBA project files from an .xlam file.')

    args = parser.parse_args()

    # if args.extract:
    #     print("Running extract process...")
    #     extract()

    if args.build:
        print("Running build process...")
        build()
    if args.push:
        print("Pushing to xlstart...")
        reload_to_xlstart()


    if not args.build and not args.extract:
        print("No action provided. Use --build or --extract.")


import sys

if __name__ == "__main__":
    sys.argv = [__file__, "--build", "--push"]
    main()
    # build()
    # reload_to_xlstart()