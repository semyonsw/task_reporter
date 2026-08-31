"""Generate windows/TaskReporter.ico - no third-party imaging libraries.

Windows accepts PNG-compressed images inside an .ico, so the whole file is a
handful of PNGs behind a directory header.  Drawing is done by supersampling a
few primitives, which is enough for a flat icon and keeps this dependency-free.
"""

import os
import struct
import zlib

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(BASE_DIR, "TaskReporter.ico")
SIZES = (256, 128, 64, 48, 32, 16)
SS = 4  # supersampling factor, for antialiased edges

BG = (22, 27, 39)        # #161B27 - the app's card colour
EDGE = (30, 38, 64)      # #1E2640
ACCENT = (109, 143, 255) # #6D8FFF
INK = (232, 240, 254)    # #E8F0FE


def _blend(dst, src, alpha):
    return tuple(round(d + (s - d) * alpha) for d, s in zip(dst, src))


def _rounded_rect(x, y, w, h, r):
    """Coverage test for a rounded rectangle.

    The corner a point belongs to is chosen by clamping it into the inset
    rectangle, so a point is only ever measured against its own arc.
    """

    def inside(px, py):
        if not (x <= px <= x + w and y <= py <= y + h):
            return False
        cx = min(max(px, x + r), x + w - r)
        cy = min(max(py, y + r), y + h - r)
        return (px - cx) ** 2 + (py - cy) ** 2 <= r * r


    return inside


def _thick_line(x1, y1, x2, y2, width):
    dx, dy = x2 - x1, y2 - y1
    length_sq = dx * dx + dy * dy
    half = width / 2.0

    def inside(px, py):
        if length_sq == 0:
            return (px - x1) ** 2 + (py - y1) ** 2 <= half * half
        t = ((px - x1) * dx + (py - y1) * dy) / length_sq
        t = max(0.0, min(1.0, t))
        nx, ny = x1 + t * dx, y1 + t * dy
        return (px - nx) ** 2 + (py - ny) ** 2 <= half * half
    return inside


def render(size):
    """Return (rgba_rows) for one square icon at `size` pixels."""
    n = size * SS
    u = n / 256.0  # design units: the artwork is drawn on a 256 grid

    plate = _rounded_rect(6 * u, 6 * u, 244 * u, 244 * u, 52 * u)
    inner = _rounded_rect(14 * u, 14 * u, 228 * u, 228 * u, 44 * u)

    # Three "report lines" down the left, shortest last.
    lines = [
        _rounded_rect(56 * u, 74 * u, 96 * u, 15 * u, 7 * u),
        _rounded_rect(56 * u, 118 * u, 74 * u, 15 * u, 7 * u),
        _rounded_rect(56 * u, 162 * u, 52 * u, 15 * u, 7 * u),
    ]
    # A tick, drawn over the lines' right-hand side.
    tick_a = _thick_line(126 * u, 158 * u, 158 * u, 190 * u, 26 * u)
    tick_b = _thick_line(158 * u, 190 * u, 214 * u, 84 * u, 26 * u)

    # Supersampled coverage accumulation, one pass per layer.
    acc = [[(0, 0, 0, 0.0) for _ in range(size)] for _ in range(size)]

    for py in range(size):
        for px in range(size):
            r = g = b = 0.0
            cover = 0.0
            for sy in range(SS):
                for sx in range(SS):
                    fx = px * SS + sx + 0.5
                    fy = py * SS + sy + 0.5
                    if not plate(fx, fy):
                        continue
                    colour = EDGE if not inner(fx, fy) else BG
                    if tick_a(fx, fy) or tick_b(fx, fy):
                        colour = ACCENT
                    else:
                        for index, line in enumerate(lines):
                            if line(fx, fy):
                                colour = INK if index == 0 else _blend(BG, INK, 0.55)
                                break
                    r += colour[0]
                    g += colour[1]
                    b += colour[2]
                    cover += 1.0
            total = SS * SS
            if cover == 0:
                acc[py][px] = (0, 0, 0, 0.0)
            else:
                acc[py][px] = (
                    round(r / cover), round(g / cover), round(b / cover),
                    cover / total,
                )
    return acc


def to_png(pixels):
    size = len(pixels)
    raw = bytearray()
    for row in pixels:
        raw.append(0)  # filter type 0
        for r, g, b, a in row:
            raw += bytes((r, g, b, round(a * 255)))

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(
            ">I", zlib.crc32(body) & 0xFFFFFFFF
        )

    header = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
        + chunk(b"IEND", b"")
    )


def main():
    images = []
    for size in SIZES:
        images.append((size, to_png(render(size))))
        print(f"  rendered {size}x{size}")

    offset = 6 + 16 * len(images)
    directory = bytearray()
    body = bytearray()
    for size, png in images:
        directory += struct.pack(
            "<BBBBHHII",
            0 if size >= 256 else size,
            0 if size >= 256 else size,
            0, 0, 1, 32, len(png), offset,
        )
        body += png
        offset += len(png)

    with open(OUT_PATH, "wb") as handle:
        handle.write(struct.pack("<HHH", 0, 1, len(images)))
        handle.write(directory)
        handle.write(body)
    print(f"wrote {OUT_PATH} ({os.path.getsize(OUT_PATH)} bytes)")


if __name__ == "__main__":
    main()
