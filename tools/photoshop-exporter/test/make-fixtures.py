#!/usr/bin/env python3
"""Create tiny, deterministic PSD fixtures for the actual Photoshop plugin tests.

Uses the public PSD v1 raw-channel format, no third-party PSD parser or renderer.
The single layer has transparent margins, half-transparent red and opaque blue.
"""
import argparse
import pathlib
import struct
import zlib


def u16(value):
    return struct.pack('>H', value)


def u32(value):
    return struct.pack('>I', value)


def chunk(kind, data):
    return u32(len(data)) + kind + data + u32(zlib.crc32(kind + data))


def write_fixture(output, stem, empty=False):
    width, height = 64, 48
    left, top, lw, lh = 7, 9, 30, 22
    pixel = lambda x, y: bytes((255, 0, 0, 128)) if x < 15 else bytes((0, 0, 255, 255))
    rgba = bytes(lw * lh * 4) if empty else b''.join(pixel(x, y) for y in range(lh) for x in range(lw))
    planes = [rgba[c::4] for c in range(4)]
    channels = b''.join(u16(c if c < 3 else 65535) + u32(2 + len(plane))
                        for c, plane in enumerate(planes))
    name = b'Alpha fixture'
    pascal = bytes([len(name)]) + name
    pascal += b'\0' * (-len(pascal) % 4)
    extra = u32(0) + u32(0) + pascal
    record = struct.pack('>iiii', top, left, top + lh, left + lw) + u16(4) + channels
    record += b'8BIMnorm' + bytes((255, 0, 8, 0)) + u32(len(extra)) + extra
    layer_info = struct.pack('>h', -1) + record + b''.join(u16(0) + p for p in planes)
    layer_info += b'\0' * (len(layer_info) % 2)
    layer_mask = u32(len(layer_info)) + layer_info + u32(0)
    composite = bytearray(width * height * 4)
    for y in range(lh):
        offset = ((y + top) * width + left) * 4
        composite[offset:offset + lw * 4] = rgba[y * lw * 4:(y + 1) * lw * 4]
    header = b'8BPS' + u16(1) + bytes(6) + u16(4) + u32(height) + u32(width) + u16(8) + u16(3)
    psd = header + u32(0) + u32(0) + u32(len(layer_mask)) + layer_mask
    psd += u16(0) + b''.join(composite[c::4] for c in range(4))
    (output / f'{stem}-fixture.psd').write_bytes(psd)
    rows = b''.join(b'\0' + composite[y * width * 4:(y + 1) * width * 4] for y in range(height))
    png = b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0))
    png += chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b'')
    (output / f'{stem}-expected.png').write_bytes(png)
    print(output / f'{stem}-fixture.psd')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('output', type=pathlib.Path)
    options = parser.parse_args()
    options.output.mkdir(parents=True, exist_ok=True)
    write_fixture(options.output, 'alpha')
    write_fixture(options.output, 'empty', empty=True)
    (options.output / 'selected-fixture.psd').write_bytes((options.output / 'alpha-fixture.psd').read_bytes())


if __name__ == '__main__':
    main()
