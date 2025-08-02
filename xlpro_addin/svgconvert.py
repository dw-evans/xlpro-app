from __future__ import annotations
from pathlib import Path
import subprocess
import struct
from pathlib import Path
from PIL import Image
import multiprocessing

# def write_ico_from_pngs(png_paths: list[str], output_path: str):
#     """
#     Write a multi-resolution ICO file from exact PNG images (no resizing/compression).

#     :param png_paths: list of PNG file paths, sorted from smallest to largest size (width==height)
#     :param output_path: output .ico file path
#     """
#     with open(output_path, 'wb') as f:
#         num_images = len(png_paths)
#         f.write(struct.pack('<3H', 0, 1, num_images))  # Reserved, type=1, count

#         dir_entries = []
#         image_data = b''
#         offset = 6 + num_images * 16  # header + entry table

#         for path in png_paths:
#             data = Path(path).read_bytes()
#             img = data

#             with open(path, 'rb') as pf:
#                 if pf.read(8) != b'\x89PNG\r\n\x1a\n':
#                     raise ValueError(f"{path} is not a valid PNG")

#             from PIL import Image
#             im = Image.open(path)
#             width, height = im.size
#             width_byte = width if width < 256 else 0
#             height_byte = height if height < 256 else 0

#             width_byte = width if width < 256 else 0
#             height_byte = height if height < 256 else 0

#             entry = struct.pack(
#                 '<BBBBHHII',
#                 # width_byte,
#                 width_byte,
#                 # height_byte,
#                 height_byte,
#                 0,      # Color count
#                 0,      # Reserved
#                 1,      # Color planes
#                 32,     # Bits per pixel
#                 len(img),
#                 offset
#             )

#             dir_entries.append(entry)
#             image_data += img
#             offset += len(img)

#         # Write all directory entries
#         for entry in dir_entries:
#             f.write(entry)

#         # Write image data
#         f.write(image_data)


# def bake_several_pngs_to_ico(sourcefiles, targetfile):
#     """Converts several PNG files into one ICO file.

#     args:
#         sourcefiles (list of str): A list of pathnames of PNG files.
#         targetfile (str): Pathname of the resulting ICO file.

#     Use this function if you want to have fine-grained control over
#     the resulting icon file, providing each possible icon resolution
#     individually.

#     Example::

#         sourcefiles = [
#             "Path/to/logo_16x16.png",
#             "Path/to/logo_32x32.png",
#             "Path/to/logo_48x48.png"
#         ]
#         targetfile = "Path/to/logo.ico"
#         bake_several_pngs_to_ico(sourcefiles, targetfile)
#     """

#     # Write the global header
#     number_of_sources = len(sourcefiles)
#     data = bytes((0, 0, 1, 0, number_of_sources, 0))
#     offset = 6 + number_of_sources * 16

#     # Write the header entries for each individual image
#     for sourcefile in sourcefiles:
#         img = Image.open(sourcefile)
#         w = 0 if img.width < 256 else 0
#         h = 0 if img.width < 256 else 0
#         data += bytes((w, h, 0, 0, 1, 0, 32, 0, ))
#         bytesize = Path(sourcefile).stat().st_size
#         data += bytesize.to_bytes(4, byteorder="little")
#         data += offset.to_bytes(4, byteorder="little")
#         offset += bytesize

#     # Write the individual image data
#     for sourcefile in sourcefiles:
#         data += Path(sourcefile).read_bytes()

#     # Save the icon file
#     Path(targetfile).write_bytes(data)


# def write_ico_from_pngs(iconfile, imagedatas, imagenames):
#     """
#     Write icon to already-opened iconfile for already-loaded imagedatas.
#     All integer items in file are written little-endian, MSB last ('<').

#     For Windows, packs any-sized images into file after a directory,
#     and Windows will choose and possibly scale an appropriate match.
#     This is substantially simpler than Mac .icns format (see ahead).

#     [2.0] Recoded to order images by most-to-least pixels.  This may
#     or may not matters in modern Windows.  Window title-bar icons may
#     be blurrier in the opposite order, and the new order shouldn't hurt.
    
#     Hint: large images may be best - scaling down is better than up.
#     Caveat: this script supports PNG images, but for simplicity not BMP.
#     """
    
#     # 1) HEADER: 3 2-byte ints

#     hdrsize = 6
#     hdrfmt  = "<hhh"
#     iconfile.write(struct.pack(hdrfmt,
#                     0,                       # reserved, always 0
#                     1,                       # 1=.ico icon, 2=.cur cursor 
#                     len(imagedatas)          # number images in file
#                     ))

#     # 2) DIRECTORY: 1..N 16-byte records, bytes/shorts/ints, for width/height/size/offset/etc.

#     dirsize = 16
#     dirfmt  = "<BBBBhhii"
#     imgoffset = hdrsize + (dirsize * len(imagedatas))

#     # bit per pixel: this table seems a hack, but no direct map in PIL itself?
#     # example: common PNG => 'RGBA' => 32bpp (8 bits (byte) * 4 bands (RGBA))
#     mode_to_bpp = {'1': 1, 'L': 8, 'P': 8,
#                    'RGB': 24, 'RGBA': 32, 'CMYK': 32, 'YCbCr': 24, 'I': 32, 'F': 32}

#     # analyze with PIL, add rank and image info
#     withRankAndPIL = []
#     for (imagedata, imagename) in zip(imagedatas, imagenames):
#         pilimg = Image.open(imagename)
#         size = pilimg.size
#         width, height = size       # a 2-tuple
#         rank = width * height
#         bitsperpixel = mode_to_bpp[pilimg.mode]
#         withRankAndPIL.append((rank, imagedata, imagename, size, bitsperpixel))

#     # order most-to-least pixels (descending)
#     withRankAndPIL.sort(key=(lambda x: x[0]), reverse=True)

#     # write directory entries
#     for (rank, imagedata, imagename, size, bitsperpixel) in withRankAndPIL:
#         width, height = size                 # a 2-tuple
#         if width  >= 256: width  = 0         # 0 means 256 (or more), per spec (and practice)
#         if height >= 256: height = 0
#         # if verbose:
#             # print('Adding:', imagename, '[%dx%d] %s' % (width, height, size))
            
#         iconfile.write(struct.pack(dirfmt,
#                     width,                   # width, in pixels,  0..255
#                     height,                  # height, in pixels, 0..255
#                     0,                       # ?color count/palette (0 if >= 8bpp)
#                     0,                       # reserved, always 0
#                     1,                       # ?color planes: 0 or 1 treated same (>1 matters)
#                     bitsperpixel,            # ?bits per pixel (0 apparently means inferred)
#                     len(imagedata),          # size of image data in bytes
#                     imgoffset                # offset to image data from start of file            
#                     ))
#         imgoffset += len(imagedata)

#     # 3) IMAGE BYTES: packed into rest of icon file sequentially

#     # in same order as directory entries
#     for imagedata in (info[1] for info in withRankAndPIL):
#         iconfile.write(imagedata)


def write_ico_from_pngs(png_fps:list[Path], ico_fp:Path):
    bheader = bytes()
    bentries = bytes()
    bpngdata = bytes()
    # header
    bheader += struct.pack("<HHH", 0, 1, len(png_fps))

    offset = len(bheader)
    offset = 6 + 16 *len(png_fps)

    for png_fp in png_fps:
        img = Image.open(png_fp)
        w, h = img.width, img.height
        w = w if w < 256 else 0
        h = h if h < 256 else 0
        print(w, h)
        with open(png_fp, "rb") as f:
            png_bytes = f.read()
        img = Image.open(png_fp)
        print(img.info.get("dpi", "dpi unavailable"))
        bentries += (v:=struct.pack("<BBBBHHII", w, h, 0, 0, 1, 32, len(png_bytes), offset))
        bpngdata += png_bytes
        offset += len(png_bytes)
        pass

    with open(ico_fp, "wb") as f:
        f.write(bheader)
        f.write(bentries)
        f.write(bpngdata)



def convert_svg_to_png(w, h, padding:int, svg_path:str, output_png:str):
    # Step 3: Call Inkscape to export PNG
    attempts = 0
    max_attempts = 20
    finished = False
    while not finished:
        try:
            result = subprocess.run([
                # "inkscape",
                r"C:\Program Files\Inkscape\bin\inkscape.exe",
                svg_path,
                f"--export-type=png",
                f"--export-filename={output_png}",
                f"--export-width={w}",
                f"--export-height={h}",
            ], 
            check=True,
            )
            finished=True
        except:
            print("error occurred")
        attempts += 1
        if attempts > max_attempts:
            raise Exception("Failed")
            break

    img = Image.open(output_png)
    
    pass

    if padding > 0:
        # create a modified image for padding
        new_width = img.width + padding * 2
        new_height = img.height + padding * 2
        # Create new image with transparent background (RGBA)
        new_img = Image.new("RGBA", (new_width, new_height), (0, 0, 0, 0))

        # Paste original image in the center
        new_img.paste(img, (padding, padding))

    else:
        new_img = Image.open(output_png)
        
    new_img.save(output_png, dpi=(96, 96))
    img_loaded = Image.open(output_png)

    dpi = img_loaded.info.get("dpi")
    print(f"dpi={dpi}")

    pass


    print(f"Exported to {output_png}")

# Example usage

wd = Path(__file__).parent
d_in = wd / "assets"
d_out = wd / "assets" / f"exports"
d_out.mkdir(exist_ok=True)

def process_icon(fp:Path):
    # sizes = [16, 24, 32, 48, 64, 128, 256]
    sizes = [16, 24, 32, 48, 64]
    # sizes.sort(reverse=True)
    for padding_frac in [0.0, 0.125]:
        png_fps = []
        is_padded = padding_frac > 0
        ico_fout = d_out /  f"{fp.stem}_{int(is_padded)}.ico"
        for size in sizes:  
            padding = round(size * padding_frac)
            convert_svg_to_png(
                size - padding * 2, 
                size - padding * 2, 
                padding,
                str(fp.resolve()), 
                str(fout:=((d_out / f"{fp.stem}_{int(is_padded)}_{size}.png").resolve())), 
            )
            png_fps.append(fout)

        # with open(ico_fout, "wb") as f:
        #     imagedatas = []
        #     imagenames = [] 
        #     for fp in fps:
        #         with open(fp, "rb") as fimg:
        #             imagedatas.append(fimg.read())
        #         imagenames.append(str(fp))
                
        #     write_ico_from_pngs(f, imagedatas, imagenames)

        # write_ico_from_pngs(png_fps, ico_fout)

        # bake_several_pngs_to_ico([str(x) for x in fps], str(ico_fout))
        # bake_several_pngs_to_ico([str(x) for x in fps], str(ico_fout))
        pass

if __name__ == "__main__":
    fp_list = list(d_in.glob("*.svg"))
    # process_icon(fp_list[0])

    # sys.exit()

    for p in d_out.glob("*.png"):
        t = """<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/{}.png" Id="{}" />"""
        print("    " + t.format(p.stem, p.stem))


    num_workers = multiprocessing.cpu_count()
    with multiprocessing.Pool(processes=num_workers) as pool:
        results = pool.map(process_icon, fp_list)

    pass



    