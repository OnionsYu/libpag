#!/usr/bin/env python3
"""Check host-smoke exports against independent pixels and the PAGX renderer (requires Pillow)."""
import argparse
import base64
import io
import json
import pathlib
import subprocess
import xml.etree.ElementTree as ET
from PIL import Image


def assert_pixels(actual, expected):
    assert actual.size == expected.size, (actual.size, expected.size)
    a, b = actual.convert('RGBA').tobytes(), expected.convert('RGBA').tobytes()
    # RGB under zero alpha is undefined. Compare every alpha and every visible RGB component.
    differences = sum(a[i + 3] != b[i + 3] or (a[i + 3] and a[i:i + 3] != b[i:i + 3])
                      for i in range(0, len(a), 4))
    assert differences == 0, f'{differences} pixels differ'


def verify(path, expected, render):
    root = ET.parse(path).getroot()
    assert root.tag == 'pagx'
    images = root.findall('./Resources/Image')
    assert len(images) == 1, 'Composite fixture should contain exactly one embedded image'
    source = images[0].attrib['source']
    assert source.startswith('data:image/png;base64,')
    image = Image.open(io.BytesIO(base64.b64decode(source.split(',', 1)[1]))).convert('RGBA')
    assert image.size == (int(root.attrib['width']), int(root.attrib['height']))
    if expected is not None:
        assert_pixels(image, expected)
    subprocess.run(['pagx', 'verify', '--skip-render', '--skip-layout', '--json', str(path)], check=True)
    rendered = path.with_name(path.stem + '-rendered.png')
    if render:
        subprocess.run(['pagx', 'render', str(path), '-o', str(rendered)], check=True)
    assert rendered.exists(), f'Render first with pagx render -o "{rendered}", or use --render'
    assert_pixels(Image.open(rendered), image)
    print(f'PASS {path.name}: structural validation and rendered pixels match')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('folder', type=pathlib.Path)
    parser.add_argument('--render', action='store_true', help='Run pagx render (requires graphics device access)')
    parser.add_argument('--dirty-export', type=pathlib.Path, help='Optional manual 50%% opacity export')
    args = parser.parse_args()
    report = json.loads(sorted(args.folder.glob('host-report-*.json'))[-1].read_text())
    alpha = Image.open(args.folder / 'alpha-expected.png').convert('RGBA')
    empty = Image.open(args.folder / 'empty-expected.png').convert('RGBA')
    for case in report['cases']:
        assert case['passed'], case
        if 'output' not in case:
            continue
        expected = empty if case['name'] == 'empty-fixture.psd' else alpha
        if case['name'] == 'all-layer-types.psd':
            expected = None  # Complex fixture checks transport/rendering, not an independent PS baseline.
        verify(args.folder / case['output'], expected, args.render)
    if args.dirty_export:
        data = bytearray(alpha.tobytes())
        # Photoshop stores UI opacity 50% as 128/255. Half-alpha red -> 64, opaque blue -> 128.
        data[3::4] = bytes((value * 128 + 127) // 255 for value in data[3::4])
        verify(args.dirty_export, Image.frombytes('RGBA', alpha.size, bytes(data)), args.render)


if __name__ == '__main__':
    main()
