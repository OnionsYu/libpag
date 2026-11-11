/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

namespace photoshop {

// This source model deliberately has no Photoshop SDK or PAGX dependencies. Coordinates are
// document pixels, children are in bottom-to-top paint order, and opacity values are in [0, 1].
struct Rect {
  float left = 0;
  float top = 0;
  float width = 0;
  float height = 0;
};

enum class LayerKind {
  Unknown,
  Pixel,
  Group,
  Text,
  Shape,
  Fill,
  SmartObject,
  Adjustment,
  Artboard
};

struct SolidRectangle {
  Rect bounds = {};
  float red = 0;
  float green = 0;
  float blue = 0;
  float alpha = 1;
  float roundness = 0;
};

struct Point {
  float x = 0;
  float y = 0;
};
struct PathPoint {
  Point anchor = {};
  Point incoming = {};
  Point outgoing = {};
};
struct Subpath {
  bool closed = true;
  std::vector<PathPoint> points = {};
};
enum class PathOperation { Add, Subtract, Intersect, Exclude };
struct PathComponent {
  PathOperation operation = PathOperation::Add;
  bool winding = false;
  std::vector<Subpath> subpaths = {};
};
struct Color {
  float red = 0;
  float green = 0;
  float blue = 0;
  float alpha = 1;
};
struct Stroke {
  Color color = {};
  float width = 1;
  float opacity = 1;
  float miterLimit = 4;
  std::string cap = "butt";
  std::string join = "miter";
  std::string align = "center";
  std::vector<float> dashes = {};
  float dashOffset = 0;
};
// Control points are absolute document pixels AFTER the host's shape transform. Never reapply
// the live-shape origin transform or infer geometry from the raster bounds.
struct Shape {
  std::vector<PathComponent> components = {};
  std::optional<Color> fill = {};
  std::optional<Stroke> stroke = {};
};

struct Layer {
  std::string id = {};
  int64_t sourceLayerId = -1;
  std::string name = {};
  LayerKind kind = LayerKind::Unknown;
  bool visible = true;
  float opacity = 1;
  float fillOpacity = 1;
  std::string blendMode = "unknown";
  // False means the adapter could not establish compositing semantics. Never infer "normal".
  bool compositingKnown = false;
  std::string clippingBase = {};
  // Nonempty means the current native mapping cannot prove backdrop independence.
  std::vector<std::string> features = {};
  // Known local appearance that needs a host render, but does not read the backdrop.
  std::vector<std::string> rasterFeatures = {};
  Rect bounds = {};
  std::string text = {};
  std::optional<SolidRectangle> rectangle = {};
  std::optional<Shape> shape = {};
  std::vector<Layer> children = {};
};

struct Document {
  int schemaVersion = 2;
  std::string name = {};
  std::string hostVersion = {};
  float resolution = 72;
  std::string colorMode = "RGB";
  std::string colorProfile = "sRGB IEC61966-2.1";
  int bitDepth = 8;
  // Explicit export rectangle. The importer translates its origin to PAGX (0, 0).
  Rect bounds = {};
  std::vector<std::string> features = {};
  std::vector<Layer> layers = {};
};

struct Diagnostic {
  std::string code;
  std::string layerId;
  std::string message;
};

// Captures are document-composite renders of a planned sibling range or subtree. Only its source
// visibility and necessary ancestors are retained. Ancestor opacity is neutralized; the captured
// roots' appearance, opacity and fill are baked. A captured group includes its descendants. Images
// must be full-resolution, unassociated-alpha, sRGB PNGs. Transparent captures have empty bytes
// and zero size. Source metadata remains separate from these derived rendering resources.
struct RasterAsset {
  std::string requestId;
  Rect bounds = {};
  std::vector<uint8_t> png = {};
};

// Throws std::invalid_argument for unsupported versions, malformed trees, dangling references,
// non-finite numbers, invalid ranges, or excessive depth/node count. Does not silently repair data.
void Validate(const Document& document);

// Strict JSON transport decoder. Required fields are never replaced with implicit defaults.
Document ParseSnapshot(const std::string& json);

}  // namespace photoshop
