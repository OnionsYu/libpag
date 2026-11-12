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


def load_reference(folder, reference_name):
    meta = json.loads((folder / f"{reference_name}.reference.json").read_text())
    if meta.get("level", 0) != 0:
        raise AssertionError(f"{reference_name}: Photoshop returned a subsampled reference")
    reference = Image.new("RGBA", (meta["width"], meta["height"]))
    if meta["pixelWidth"] and meta["pixelHeight"]:
        pixels = Image.frombytes("RGBA" if meta["components"] == 4 else "RGB",
                                 (meta["pixelWidth"], meta["pixelHeight"]),
                                 (folder / f"{reference_name}.rgba").read_bytes()).convert("RGBA")
        reference.paste(pixels, (meta["sourceBounds"]["left"], meta["sourceBounds"]["top"]))
    reference.save(folder / f"{reference_name}.reference.png")
    return reference


def verify(folder, name, max_mae, max_difference, expected_images, reference_suffix="", replay=None,
           scales=(1, 4), cli="pagx"):
    if replay:
        path = folder / f"{name}.replayed.pagx"
        subprocess.run([str(replay), str(folder / f"{name}.snapshot.json"), str(path)], check=True)
    else:
        export = json.loads((folder / f"{name}.export.json").read_text())
        path = folder / export["output"]
    root = ET.parse(path).getroot()
    counts = {tag: len(list(root.iter(tag))) for tag in ("Path", "Rectangle", "Image", "Text")}
    if expected_images is not None and counts["Image"] != expected_images:
        raise AssertionError(f"{name}: expected {expected_images} images, found {counts}")
    if expected_images == 0 and not (counts["Path"] or counts["Rectangle"]):
        raise AssertionError(f"{name}: missing vector paths")
    validation = subprocess.run([cli, "verify", str(path), "--skip-render", "--skip-layout", "--json"],
                                capture_output=True, text=True)
    diagnostics = json.loads(validation.stdout)["diagnostics"]
    # The generic vector emitter deliberately retains source paths. These two CLI diagnostics
    # recommend primitive optimizations; they do not indicate structural or rendering errors.
    optimizations = ("Path draws an axis-aligned rectangle.", "Path draws an ellipse.")
    if validation.returncode and (not diagnostics or any(
            not item["message"].startswith(optimizations) for item in diagnostics)):
        raise AssertionError(f"{name}: PAGX validation failed: {diagnostics}")
    comparisons = []
    for scale in scales:
        suffix = "" if scale == 1 else f".{scale}x"
        reference_name = name + reference_suffix + suffix
        # Require independently rendered host pixels at every requested scale. Never upscale
        # the 1x reference or count merely producing a PNG as passing a comparison.
        reference = load_reference(folder, reference_name)
        rendered = folder / f"{name}{suffix}.render.png"
        subprocess.run([cli, "render", str(path), "--scale", str(scale), "-o", str(rendered)], check=True)
        if name == "hole" and reference.getpixel((100 * scale, 80 * scale))[3] != 0:
            raise AssertionError("The Photoshop compound-path fixture does not contain a transparent hole")
        metrics, interior_max = compare(Image.open(rendered), reference)
        comparisons.append({"scale": scale, "reference": reference_name, "metrics": metrics,
                            "interiorMax": interior_max})
        if interior_max > 2:
            raise AssertionError(f"{name}@{scale}x: non-edge error {interior_max} exceeds two 8-bit steps")
        for channel, values in metrics.items():
            if values["mae"] > max_mae or values["max"] > max_difference:
                raise AssertionError(f"{name}@{scale}x/{channel}: {values} exceeds MAE {max_mae}, max {max_difference}")
    result = {"name": name, "output": path.name, "replayed": bool(replay), "nodes": counts,
              "comparisons": comparisons, "optimizationHints": diagnostics}
    print(json.dumps(result, ensure_ascii=False))
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", type=Path)
    parser.add_argument("--names", nargs="+", default=["shapes", "hole", "stroke", "transformed"])
    # Different vector rasterizers use different edge coverage. Interior errors remain <= 2.
    parser.add_argument("--max-mae", type=float, default=0.5)
    parser.add_argument("--max-difference", type=float, default=64)
    parser.add_argument("--images", type=int, default=0)
    parser.add_argument("--scales", type=int, nargs="+", default=[1, 4], choices=range(1, 17))
    parser.add_argument("--cli", default="pagx", help="Exact PAGX executable under test")
    parser.add_argument("--reference-suffix", default="")
    parser.add_argument("--replay", type=Path, help="Replay vector-only snapshots using the current core executable")
    parser.add_argument("--report", default="vector-verification.json", help="Report filename inside the fixture folder")
    args = parser.parse_args()
    results = [verify(args.folder, name, args.max_mae, args.max_difference, args.images,
                      args.reference_suffix, args.replay, args.scales, args.cli) for name in args.names]
    (args.folder / args.report).write_text(json.dumps(results, indent=2) + "\n")
