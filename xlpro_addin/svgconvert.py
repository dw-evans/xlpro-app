from pathlib import Path
import shutil
import subprocess
import os

WIDTH, HEIGHT = 24, 24

# import xml.etree.ElementTree as ET

# def change_stroke(fp, stroke:str):
#     tree = ET.parse(fp)
#     root = tree.getroot()

#     for elem in root.iter():
#         if "stroke" in elem.attrib:
#             elem.set("stroke-width", s)  # or any value

#     with open("temp.svg", "w") as f:
#         tree.write(f)

#     return f

from PIL import Image

def convert_svg_to_png(w, h, padding:int, svg_path:str, output_png:str):
    # Step 3: Call Inkscape to export PNG
    subprocess.run([
        # "inkscape",
        r"C:\Program Files\Inkscape\bin\inkscape.exe",
        svg_path,
        f"--export-type=png",
        f"--export-filename={output_png}",
        f"--export-width={w}",
        f"--export-height={h}"
    ], 
    check=True,
    )
    img = Image.open(output_png)

    if padding > 0:
        # Define padding (4px on all sides)
        new_width = img.width + padding * 2
        new_height = img.height + padding * 2
        # Create new image with transparent background (RGBA)
        new_img = Image.new("RGBA", (new_width, new_height), (0, 0, 0, 0))

        # Paste original image in the center
        new_img.paste(img, (padding, padding))

        # Save the padded image
        new_img.save(output_png)

    print(f"Exported to {output_png}")

# Example usage

wd = Path(__file__).parent
d_in = wd / "assets"
d_out = wd / "assets" / f"exports"
d_out.mkdir(exist_ok=True)
if __name__ == "__main__":
    for fp in d_in.glob("*.svg"):
        if fp.stem.endswith("16"):
            convert_svg_to_png(
                16, 
                16, 
                0,
                str(fp.resolve()), 
                str((d_out / f"{fp.stem}.png").resolve()), 
            )
        elif fp.stem.endswith("24"):
            convert_svg_to_png(
                24, 
                24, 
                4,
                str(fp.resolve()), 
                str((d_out / f"{fp.stem}.png").resolve()), 
            )

    for p in d_out.glob("*"):
        t = """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/{}.png" Id="{}" />"""
        print("    " + t.format(p.stem, p.stem))

    