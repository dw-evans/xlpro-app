input_str = """
██╗  ██╗██╗     ██████╗ ██████╗  ██████╗ 
╚██╗██╔╝██║     ██╔══██╗██╔══██╗██╔═══██╗
 ╚███╔╝ ██║     ██████╔╝██████╔╝██║   ██║
 ██╔██╗ ██║     ██╔═══╝ ██╔══██╗██║   ██║
██╔╝ ██╗███████╗██║     ██║  ██║╚██████╔╝
╚═╝  ╚═╝╚══════╝╚═╝     ╚═╝  ╚═╝ ╚═════╝ 
"""[1:-1]


import numpy as np

def hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    """Convert hex color string to RGB tuple."""
    hex_str = hex_str.lstrip('#')
    if len(hex_str) != 6:
        raise ValueError("Input must be a 6-digit hex string.")
    return np.array(tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4)), dtype=np.int32)


def rgb_to_hsl(r: int, g: int, b: int) -> tuple[float, float, float]:
    """Convert RGB (0–255) to HSL (H in 0–360, S and L in 0–1)."""
    r, g, b = [x / 255.0 for x in (r, g, b)]
    max_c = max(r, g, b)
    min_c = min(r, g, b)
    delta = max_c - min_c

    # Lightness
    l = (max_c + min_c) / 2

    # Saturation
    if delta == 0:
        s = 0
        h = 0  # achromatic
    else:
        s = delta / (1 - abs(2 * l - 1))

        # Hue
        if max_c == r:
            h = 60 * (((g - b) / delta) % 6)
        elif max_c == g:
            h = 60 * (((b - r) / delta) + 2)
        else:
            h = 60 * (((r - g) / delta) + 4)

    return np.array((round(h, 2), round(s, 4), round(l, 4)), dtype=np.int32)



colors = [
    "#526cfe",
    "#182e49",
]

gcoords = np.array([
        (0, 0),
        (max([len(x) for x in input_str.split("\n")]), 0),
    ],
    dtype=np.int32
)

grid = np.meshgrid(*[np.arange(*x, step=1, dtype=np.int32) for x in gcoords.T])


def directional_gradient(width, height, x0, y0, x1, y1, color_start, color_end):
    """
    Creates a 2D gradient from (x0, y0) to (x1, y1) over a grid of size (width, height).
    color_start and color_end: (R, G, B) tuples
    """

    # Create a grid of pixel coordinates
    y, x = np.meshgrid(np.arange(height), np.arange(width), indexing='ij')
    coords = np.stack([x, y], axis=-1)  # shape: (H, W, 2)

    # Define gradient direction vector
    grad_vec = np.array([x1 - x0, y1 - y0], dtype=float)
    grad_len_sq = np.dot(grad_vec, grad_vec)

    # Avoid divide-by-zero if start == end
    if grad_len_sq == 0:
        return np.full((height, width, 3), color_start, dtype=np.uint8)

    # Project each pixel onto the gradient vector
    rel_coords = coords - np.array([x0, y0])
    t = np.clip(np.sum(rel_coords * grad_vec, axis=-1) / grad_len_sq, 0, 1)  # shape: (H, W)

    # Linear interpolation of colors
    color_start = np.array(color_start, dtype=float)
    color_end = np.array(color_end, dtype=float)
    result = (1 - t[..., None]) * color_start + t[..., None] * color_end

    return result.astype(np.uint8)

def rgb_to_ansi(r, g, b, background=False):
    """
    Returns ANSI escape code for given RGB color.
    Set background=True for background color.
    """
    code = 48 if background else 38
    return f"\033[{code};2;{r};{g};{b}m"


x0 = 0
y0 = 0

x1 = max([len(x) for x in input_str.split("\n")]) - 1
y1 = len(input_str.split("\n")) + 1

width = x1 - x0 + 1
height = y1 - y0 + 1

color1 = "#526cfe"
color2 = "#182e49"

gradient = directional_gradient(width, height, x0, y0, x1, y1, hex_to_rgb(color1), hex_to_rgb(color2))

s_list = []
input_lines = input_str.splitlines()

for i, line in enumerate(input_lines):
    new_line = ""
    for j, char in enumerate(line):
        if char == " ":
            new_line += char
            continue
        new_line += rgb_to_ansi(*gradient[i, j]) + char
        pass
    s_list.append(new_line)

new_s = "\n".join(s_list)

print(new_s)

pass


pass