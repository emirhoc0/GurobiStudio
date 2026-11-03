#!/usr/bin/env python3
"""Writes the original ribbon icons as PNGs using only the stdlib."""

import struct
import sys
import zlib
from pathlib import Path

BLUE = (43, 117, 200)
GREEN = (30, 150, 90)
RED = (200, 60, 50)
GRAY = (110, 115, 125)


def seg(x, y, ax, ay, bx, by, width):
    dx, dy = bx - ax, by - ay
    t = max(0.0, min(1.0, ((x - ax) * dx + (y - ay) * dy) / (dx * dx + dy * dy)))
    return (x - ax - t * dx) ** 2 + (y - ay - t * dy) ** 2 <= (width / 2) ** 2


def table(x, y):
    w = 0.09
    on_v = any(seg(x, y, c, 0.14, c, 0.86, w) for c in (0.14, 0.5, 0.86))
    return on_v or any(seg(x, y, 0.14, r, 0.86, r, w) for r in (0.14, 0.38, 0.62, 0.86))


def code(x, y):
    w = 0.11
    return (
        seg(x, y, 0.34, 0.24, 0.12, 0.5, w)
        or seg(x, y, 0.12, 0.5, 0.34, 0.76, w)
        or seg(x, y, 0.66, 0.24, 0.88, 0.5, w)
        or seg(x, y, 0.88, 0.5, 0.66, 0.76, w)
        or seg(x, y, 0.58, 0.18, 0.42, 0.82, w)
    )


def play(x, y):
    return x >= 0.22 and abs(y - 0.5) <= 0.38 * (1 - (x - 0.22) / 0.62)


def stop(x, y):
    return 0.2 <= x <= 0.8 and 0.2 <= y <= 0.8


def menu(x, y):
    return any(seg(x, y, 0.16, r, 0.84, r, 0.11) for r in (0.26, 0.5, 0.74))


def app(x, y):
    # Rounded tile with a play mark knocked out.
    r, a, b = 0.18, abs(x - 0.5), abs(y - 0.5)
    inside = max(a, b) <= 0.46 and (
        min(a, b) <= 0.46 - r or (a - (0.46 - r)) ** 2 + (b - (0.46 - r)) ** 2 <= r**2
    )
    return inside and not play((x - 0.5) * 1.5 + 0.47, (y - 0.5) * 1.5 + 0.5)


ICONS = {
    "data": (table, BLUE),
    "model": (code, BLUE),
    "solve": (play, GREEN),
    "stop": (stop, RED),
    "menu": (menu, GRAY),
    "app": (app, BLUE),
}


def png(shape, color, size, ss=4):
    rows = bytearray()
    for py in range(size):
        rows.append(0)
        for px in range(size):
            hits = sum(
                shape((px + (i + 0.5) / ss) / size, (py + (j + 0.5) / ss) / size)
                for i in range(ss)
                for j in range(ss)
            )
            rows += bytes((*color, round(255 * hits / ss**2)))

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    header = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(bytes(rows)))
        + chunk(b"IEND", b"")
    )


out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
for name, (shape, color) in ICONS.items():
    for size in (16, 32, 64, 80):
        (out / f"{name}-{size}.png").write_bytes(png(shape, color, size))
