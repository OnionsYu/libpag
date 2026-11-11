#!/usr/bin/env python3
"""Compare PAGX with an independent Photoshop composite captured before export mutations."""
import argparse
import json
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image


def compare(actual, reference):
    if actual.size != reference.size:
        raise AssertionError(f"Canvas mismatch: {actual.size} != {reference.size}")
    totals = [0.0, 0.0, 0.0]
    maxima = [0.0, 0.0, 0.0]
    interior_max = 0.0
    reference_pixels = list(zip(*[iter(reference.tobytes())] * 4))
    actual_pixels = zip(*[iter(actual.convert("RGBA").tobytes())] * 4)
    for index, (p, q) in enumerate(zip(actual_pixels, reference_pixels)):
        errors = [[abs(p[3] - q[3])]]
        for background in (0, 255):
            errors.append([abs((p[c] - background) * p[3] / 255 -
                               (q[c] - background) * q[3] / 255) for c in range(3)])
        for i, values in enumerate(errors):
            totals[i] += sum(values) / len(values)
            maxima[i] = max(maxima[i], max(values))
        x, y = index % actual.width, index // actual.width
        if 0 < x < actual.width - 1 and 0 < y < actual.height - 1 and all(
                reference_pixels[index + dy * actual.width + dx] == q
                for dy in (-1, 0, 1) for dx in (-1, 0, 1)):
            interior_max = max(interior_max, *(max(values) for values in errors))
    metrics = {name: {"mae": total / (actual.width * actual.height), "max": maximum}
               for name, total, maximum in zip(("alpha", "onBlack", "onWhite"), totals, maxima)}
    return metrics, interior_max


def verify(folder, name, max_mae, max_difference, expected_images, reference_suffix="", replay=None):
    reference_name = name + reference_suffix
    meta = json.loads((folder / f"{reference_name}.reference.json").read_text())
    if replay:
        path = folder / f"{name}.replayed.pagx"
        subprocess.run([str(replay), str(folder / f"{name}.snapshot.json"), str(path)], check=True)
    else:
        export = json.loads((folder / f"{name}.export.json").read_text())
        path = folder / export["output"]
    root = ET.parse(path).getroot()
    counts = {tag: len(list(root.iter(tag))) for tag in ("Path", "Image", "Text")}
    if expected_images is not None and counts["Image"] != expected_images:
        raise AssertionError(f"{name}: expected {expected_images} images, found {counts}")
    if expected_images == 0 and not counts["Path"]:
        raise AssertionError(f"{name}: missing vector paths")
    validation = subprocess.run(["pagx", "verify", str(path), "--skip-render", "--skip-layout", "--json"],
                                capture_output=True, text=True)
    diagnostics = json.loads(validation.stdout)["diagnostics"]
    # The generic vector emitter deliberately retains source paths. These two CLI diagnostics
    # recommend primitive optimizations; they do not indicate structural or rendering errors.
    optimizations = ("Path draws an axis-aligned rectangle.", "Path draws an ellipse.")
    if validation.returncode and (not diagnostics or any(
            not item["message"].startswith(optimizations) for item in diagnostics)):
        raise AssertionError(f"{name}: PAGX validation failed: {diagnostics}")
    rendered = folder / f"{name}.render.png"
    subprocess.run(["pagx", "render", str(path), "-o", str(rendered)], check=True)
    subprocess.run(["pagx", "render", str(path), "--scale", "4", "-o", str(folder / f"{name}.4x.png")], check=True)
    reference = Image.new("RGBA", (meta["width"], meta["height"]))
    if meta["pixelWidth"] and meta["pixelHeight"]:
        pixels = Image.frombytes("RGBA" if meta["components"] == 4 else "RGB",
                                 (meta["pixelWidth"], meta["pixelHeight"]),
                                 (folder / f"{reference_name}.rgba").read_bytes()).convert("RGBA")
        reference.paste(pixels, (meta["sourceBounds"]["left"], meta["sourceBounds"]["top"]))
    reference.save(folder / f"{reference_name}.reference.png")
    if name == "hole" and reference.getpixel((100, 80))[3] != 0:
        raise AssertionError("The Photoshop compound-path fixture does not contain a transparent hole")
    metrics, interior_max = compare(Image.open(rendered), reference)
    result = {"name": name, "reference": reference_name, "output": path.name,
              "replayed": bool(replay), "nodes": counts, "metrics": metrics,
              "interiorMax": interior_max, "optimizationHints": diagnostics}
    print(json.dumps(result, ensure_ascii=False))
    if interior_max > 2:
        raise AssertionError(f"{name}: non-edge error {interior_max} exceeds two 8-bit steps")
    for channel, values in metrics.items():
        if values["mae"] > max_mae or values["max"] > max_difference:
            raise AssertionError(f"{name}/{channel}: {values} exceeds MAE {max_mae}, max {max_difference}")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", type=Path)
    parser.add_argument("--names", nargs="+", default=["shapes", "hole", "stroke", "transformed"])
    # Different vector rasterizers use different edge coverage. Interior errors remain <= 2.
    parser.add_argument("--max-mae", type=float, default=0.5)
    parser.add_argument("--max-difference", type=float, default=64)
    parser.add_argument("--images", type=int, default=0)
    parser.add_argument("--reference-suffix", default="")
    parser.add_argument("--replay", type=Path, help="Replay vector-only snapshots using the current core executable")
    parser.add_argument("--report", default="vector-verification.json", help="Report filename inside the fixture folder")
    args = parser.parse_args()
    results = [verify(args.folder, name, args.max_mae, args.max_difference, args.images,
                      args.reference_suffix, args.replay) for name in args.names]
    (args.folder / args.report).write_text(json.dumps(results, indent=2) + "\n")
