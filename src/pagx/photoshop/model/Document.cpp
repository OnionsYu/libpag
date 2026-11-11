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
void CheckLayers(const std::vector<Layer>& layers, size_t depth,
                 std::unordered_set<std::string>& ids) {
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
    CheckLayers(layer.children, depth + 1, ids);
  }
}
void CheckClipping(const std::vector<Layer>& layers) {
  std::unordered_set<std::string> below;
  for (const auto& layer : layers) {
    Require(layer.clippingBase.empty() || below.count(layer.clippingBase) != 0,
            "clipping base must be a preceding sibling");
    below.insert(layer.id);
    CheckClipping(layer.children);
  }
}
}  // namespace

void Validate(const Document& document) {
  Require(document.schemaVersion == 1, "unsupported schema version");
  CheckRect(document.bounds);
  Require(document.bounds.width > 0 && document.bounds.height > 0, "empty export canvas");
  Require(std::isfinite(document.resolution) && document.resolution > 0, "invalid resolution");
  Require(document.bitDepth == 8 || document.bitDepth == 16 || document.bitDepth == 32,
          "unsupported component depth");
  std::unordered_set<std::string> ids;
  CheckLayers(document.layers, 0, ids);
  CheckClipping(document.layers);
}
}  // namespace photoshop
