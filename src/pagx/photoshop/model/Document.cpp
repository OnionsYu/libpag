/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include "photoshop/Document.h"
#include <cmath>
#include <stdexcept>
#include <unordered_set>

namespace photoshop {
namespace {
void Require(bool value, const std::string& message) {
  if (!value) throw std::invalid_argument("Photoshop snapshot: " + message);
}
bool Unit(float value) {
  return std::isfinite(value) && value >= 0 && value <= 1;
}
void CheckRect(const Rect& rect) {
  Require(std::isfinite(rect.left) && std::isfinite(rect.top) && std::isfinite(rect.width) &&
              std::isfinite(rect.height) && rect.width >= 0 && rect.height >= 0 &&
              std::isfinite(rect.left + rect.width) && std::isfinite(rect.top + rect.height),
          "invalid bounds");
}
void CheckColor(const Color& color) {
  Require(Unit(color.red) && Unit(color.green) && Unit(color.blue) && Unit(color.alpha),
          "invalid color");
}
void CheckPaint(const Paint& paint) {
  CheckColor(paint.color);
  if (!paint.gradient) return;
  const auto& gradient = *paint.gradient;
  Require(!gradient.type.empty() && !gradient.interpolation.empty() && Unit(gradient.smoothness),
          "invalid gradient metadata");
  for (float value : gradient.matrix) Require(std::isfinite(value), "invalid paint transform");
  const auto& m = gradient.matrix;
  const double determinant = static_cast<double>(m[0]) * m[3] - static_cast<double>(m[1]) * m[2];
  Require(determinant != 0 && std::isfinite(determinant), "singular paint transform");
  Require(!gradient.colors.empty() && gradient.colors.size() <= 1024 && !gradient.opacity.empty() &&
              gradient.opacity.size() <= 1024,
          "invalid gradient stops");
  float previous = -1;
  for (const auto& stop : gradient.colors) {
    Require(Unit(stop.position) && stop.position >= previous && Unit(stop.midpoint),
            "invalid gradient color stop");
    CheckColor(stop.color);
    previous = stop.position;
  }
  previous = -1;
  for (const auto& stop : gradient.opacity) {
    Require(Unit(stop.position) && stop.position >= previous && Unit(stop.midpoint) &&
                Unit(stop.opacity),
            "invalid gradient opacity stop");
    previous = stop.position;
  }
}
void CheckShape(const Shape& shape, size_t& pointCount) {
  Require(!shape.components.empty() && shape.components.size() <= 100000,
          "invalid path components");
  for (const auto& component : shape.components) {
    Require(component.operation == PathOperation::Add ||
                component.operation == PathOperation::Subtract ||
                component.operation == PathOperation::Intersect ||
                component.operation == PathOperation::Exclude,
            "invalid path operation");
    Require(!component.subpaths.empty() && component.subpaths.size() <= 100000, "invalid subpaths");
    for (const auto& path : component.subpaths) {
      Require(!path.points.empty(), "empty subpath");
      pointCount += path.points.size();
      Require(pointCount <= 1000000, "path point count exceeds 1000000");
      for (const auto& point : path.points) {
        for (const auto& value : {point.anchor, point.incoming, point.outgoing}) {
          Require(std::isfinite(value.x) && std::isfinite(value.y), "invalid path point");
        }
      }
    }
  }
  if (shape.fill) CheckPaint(*shape.fill);
  if (shape.stroke) {
    const auto& stroke = *shape.stroke;
    CheckPaint(stroke.paint);
    Require(std::isfinite(stroke.width) && stroke.width >= 0 && Unit(stroke.opacity) &&
                std::isfinite(stroke.miterLimit) && stroke.miterLimit >= 1 &&
                std::isfinite(stroke.dashOffset),
            "invalid stroke dimensions");
    Require(stroke.cap == "butt" || stroke.cap == "round" || stroke.cap == "square", "invalid cap");
    Require(stroke.join == "miter" || stroke.join == "round" || stroke.join == "bevel",
            "invalid join");
    Require(stroke.align == "center" || stroke.align == "inner" || stroke.align == "outer",
            "invalid alignment");
    Require(stroke.dashes.size() <= 1024, "too many stroke dashes");
    bool positive = stroke.dashes.empty();
    for (auto dash : stroke.dashes) {
      Require(std::isfinite(dash) && dash >= 0, "invalid stroke dash");
      positive |= dash > 0;
    }
    Require(positive, "zero-length dash pattern");
  }
}
void CheckLayers(const std::vector<Layer>& layers, size_t depth,
                 std::unordered_set<std::string>& ids, size_t& pointCount, int version) {
  Require(depth <= 128, "layer nesting exceeds 128");
  for (const auto& layer : layers) {
    Require(!layer.id.empty() && ids.insert(layer.id).second, "empty or duplicate layer ID");
    Require(ids.size() <= 100000, "layer count exceeds 100000");
    Require(Unit(layer.opacity) && Unit(layer.fillOpacity), "invalid opacity");
    CheckRect(layer.bounds);
    auto container = layer.kind == LayerKind::Group || layer.kind == LayerKind::Artboard;
    Require(container || layer.children.empty(), "non-container has children");
    if (layer.rectangle) {
      Require(layer.kind == LayerKind::Shape && !container, "rectangle on a non-shape layer");
      const auto& rectangle = *layer.rectangle;
      CheckRect(rectangle.bounds);
      Require(Unit(rectangle.red) && Unit(rectangle.green) && Unit(rectangle.blue) &&
                  Unit(rectangle.alpha) && std::isfinite(rectangle.roundness) &&
                  rectangle.roundness >= 0,
              "invalid rectangle appearance");
    }
    if (layer.shape) {
      Require(version >= 2 && !layer.rectangle &&
                  (layer.kind == LayerKind::Shape || layer.kind == LayerKind::Text),
              "invalid shape owner or schema");
      CheckShape(*layer.shape, pointCount);
      Require(version >= 3 || ((!layer.shape->fill || !layer.shape->fill->gradient) &&
                               (!layer.shape->stroke || !layer.shape->stroke->paint.gradient)),
              "v3 paint in legacy snapshot");
    }
    if (layer.fill) {
      Require(version >= 3 && layer.kind == LayerKind::Fill && !layer.shape && !layer.rectangle,
              "invalid fill owner or schema");
      CheckPaint(*layer.fill);
    }
    Require(version >= 2 || layer.rasterFeatures.empty(), "v2 appearance in v1 snapshot");
    CheckLayers(layer.children, depth + 1, ids, pointCount, version);
  }
}
void CheckClipping(const std::vector<Layer>& layers) {
  std::string base;
  for (const auto& layer : layers) {
    Require(layer.clippingBase.empty() || (!base.empty() && layer.clippingBase == base),
            "clipping base must be the nearest preceding unclipped sibling");
    if (layer.clippingBase.empty()) base = layer.id;
    CheckClipping(layer.children);
  }
}
}  // namespace

void Validate(const Document& document) {
  Require(document.schemaVersion >= 1 && document.schemaVersion <= 3, "unsupported schema version");
  CheckRect(document.bounds);
  Require(document.bounds.width > 0 && document.bounds.height > 0, "empty export canvas");
  Require(std::isfinite(document.resolution) && document.resolution > 0, "invalid resolution");
  Require(document.bitDepth == 8 || document.bitDepth == 16 || document.bitDepth == 32,
          "unsupported component depth");
  std::unordered_set<std::string> ids;
  size_t pointCount = 0;
  CheckLayers(document.layers, 0, ids, pointCount, document.schemaVersion);
  CheckClipping(document.layers);
}
}  // namespace photoshop
