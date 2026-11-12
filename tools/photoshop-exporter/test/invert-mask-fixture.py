#!/usr/bin/env python3
"""Invert one vector-mask flag in a generated 8-bit PSD without changing its path records.

Use only on host-masks.js test fixtures. The merged preview is intentionally not rewritten;
Photoshop must recomposite the layered document for the independent reference.
"""
import argparse
import struct
from pathlib import Path


def invert(source, destination, layer_id):
    data = bytearray(source.read_bytes())
    read32 = lambda offset: struct.unpack_from(">I", data, offset)[0]
    if data[:6] != b"8BPS\0\1" or struct.unpack_from(">H", data, 22)[0] != 8:
        raise ValueError("Expected a generated 8-bit PSD")
    offset = 26
    for _ in range(2):
        offset += 4 + read32(offset)
    offset += 8
    count = abs(struct.unpack_from(">h", data, offset)[0])
    offset += 2
    changed = 0
    for _ in range(count):
        channels = struct.unpack_from(">H", data, offset + 16)[0]
        offset += 18 + channels * 6 + 12
        end = offset + 4 + read32(offset)
        offset += 4
        for _ in range(2):
            offset += 4 + read32(offset)
        offset += (data[offset] + 4) // 4 * 4
        blocks = {}
        while offset + 12 <= end:
            key, length = bytes(data[offset + 4:offset + 8]), read32(offset + 8)
            if data[offset:offset + 4] != b"8BIM" or offset + 12 + length > end:
                raise ValueError("Invalid test fixture metadata")
            blocks[key] = offset + 12
            offset += 12 + (length + 1) // 2 * 2
        if b"lyid" in blocks and read32(blocks[b"lyid"]) == layer_id:
            at = blocks.get(b"vmsk", blocks.get(b"vsms"))
            if at is None or read32(at) != 3 or read32(at + 4) & 1:
                raise ValueError("Expected a non-inverted v3 vector mask")
            struct.pack_into(">I", data, at + 4, read32(at + 4) | 1)
            changed += 1
        offset = end
    if changed != 1:
        raise ValueError("Expected exactly one target mask")
    with destination.open("xb") as output:
        output.write(data)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    parser.add_argument("--layer-id", type=int, required=True)
    args = parser.parse_args()
    invert(args.source, args.destination, args.layer_id)
