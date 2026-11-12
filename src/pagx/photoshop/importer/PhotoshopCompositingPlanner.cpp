/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include "PhotoshopCompositingPlanner.h"
#include <algorithm>
#include "PhotoshopPaintBuilder.h"
#include "PhotoshopShapeBuilder.h"

namespace pagx {
bool IsPhotoshopGroup(const photoshop::Layer& layer) {
  return layer.kind == photoshop::LayerKind::Group;
}
bool IsNativePhotoshopLayer(const photoshop::Layer& layer) {
  return layer.rasterFeatures.empty() &&
         (layer.rectangle || (layer.shape && CanBuildShape(*layer.shape)) ||
          (layer.fill && CanBuildPhotoshopPaint(*layer.fill)));
}
bool IsEmptyPhotoshopLayer(const photoshop::Layer& layer) {
  return layer.kind == photoshop::LayerKind::Pixel && layer.bounds.width == 0 &&
         layer.bounds.height == 0 && layer.rasterFeatures.empty();
}
namespace {
using Layers = std::vector<photoshop::Layer>;
bool EscapesIsolation(const Layers& layers, PhotoshopConversionPlan& plan) {
  bool escapes = false;
  for (const auto& layer : layers) {
    if (!layer.visible) continue;
    if (!layer.compositingKnown || layer.kind == photoshop::LayerKind::Unknown ||
        layer.kind == photoshop::LayerKind::Artboard ||
        std::any_of(layer.features.begin(), layer.features.end(), [](const auto& feature) {
          return feature == "blend:knockout" || feature == "knockout" || feature == "artboard";
        })) {
      escapes = true;
      plan.diagnostics.push_back(
          {"document-dependency", layer.id,
           "Unknown composition or escaping knockout requires the document composite."});
    }
    escapes |= EscapesIsolation(layer.children, plan);
  }
  return escapes;
}
bool NeedsBackdrop(const photoshop::Layer& layer) {
  if (!layer.visible) return false;
  if (!layer.features.empty() || layer.kind == photoshop::LayerKind::Adjustment) return true;
  if (layer.blendMode == "normal") return false;
  if (IsPhotoshopGroup(layer) && layer.blendMode == "passThrough" && layer.opacity == 1 &&
      layer.fillOpacity == 1 && layer.rasterFeatures.empty()) {
    return std::any_of(layer.children.begin(), layer.children.end(), NeedsBackdrop);
  }
  return true;
}
struct Range {
  size_t first;
  size_t last;
};

void AnalyzeSiblings(const Layers& layers, const std::string& parentId,
                     PhotoshopConversionPlan& plan) {
  std::vector<Range> ranges;
  for (size_t i = 0; i < layers.size(); ++i) {
    const auto& layer = layers[i];
    if (!layer.visible) continue;
    if (NeedsBackdrop(layer)) {
      ranges.push_back({0, i});
      plan.diagnostics.push_back({"backdrop-range", layer.id,
                                  "Render this layer together with its preceding backdrop inside "
                                  "the isolation boundary."});
    } else if (!layer.clippingBase.empty()) {
      size_t first = 0;
      while (layers[first].id != layer.clippingBase) ++first;  // Validated preceding sibling.
      ranges.push_back({first, i});
      plan.diagnostics.push_back(
          {"clipping-range", layer.id,
           "Render the clipping base and clipped siblings as one replacement."});
    }
  }
  std::sort(ranges.begin(), ranges.end(),
            [](const auto& a, const auto& b) { return a.first < b.first; });
  std::vector<Range> merged;
  for (const auto& range : ranges) {
    if (!merged.empty() && range.first <= merged.back().last)
      merged.back().last = std::max(merged.back().last, range.last);
    else
      merged.push_back(range);
  }
  size_t nextRange = 0;
  for (size_t i = 0; i < layers.size(); ++i) {
    const auto& layer = layers[i];
    if (nextRange < merged.size() && merged[nextRange].first == i) {
      auto range = merged[nextRange++];
      PhotoshopRasterRequest request{"range:" + layer.id, layer.id, {}, parentId};
      for (size_t j = range.first; j <= range.last; ++j) request.layerIds.push_back(layers[j].id);
      plan.rasters.push_back(std::move(request));
      plan.statistics.fallbackScopes++;
      i = range.last;
      continue;
    }
    if (!layer.visible) continue;
    if (IsPhotoshopGroup(layer) && layer.fillOpacity == 1 && layer.rasterFeatures.empty()) {
      AnalyzeSiblings(layer.children, layer.id, plan);
    } else if (IsNativePhotoshopLayer(layer)) {
      if (layer.kind == photoshop::LayerKind::Text) {
        plan.statistics.outlinedTexts++;
        plan.diagnostics.push_back(
            {"outlined-text", layer.id,
             "Text becomes vector outlines; text editing and font hinting are not retained."});
      } else
        plan.statistics.vectorShapes++;
    } else if (!IsEmptyPhotoshopLayer(layer)) {
      plan.rasters.push_back({"layer:" + layer.id, layer.id, {layer.id}, parentId});
      if (layer.kind == photoshop::LayerKind::Pixel && layer.rasterFeatures.empty()) {
        plan.statistics.sourceImages++;
      } else {
        plan.statistics.fallbackScopes++;
        std::string reason = "Photoshop renders this isolated content.";
        if (layer.shape && !CanBuildShape(*layer.shape))
          reason +=
              " Path operations, paint or stroke semantics are not qualified for native output.";
        if (layer.fill && !CanBuildPhotoshopPaint(*layer.fill))
          reason += " The fill paint is not qualified for native output.";
        for (const auto& feature : layer.rasterFeatures) reason += " " + feature;
        plan.diagnostics.push_back({"rasterized-content", layer.id, reason});
      }
    }
  }
}
void Composite(PhotoshopConversionPlan& plan) {
  plan.composite = true;
  plan.rasters = {{"composite", "", {}, ""}};
  plan.statistics = {};
  plan.statistics.fallbackScopes = 1;
}
}  // namespace

PhotoshopConversionPlan PlanPhotoshopComposition(const photoshop::Document& source,
                                                 bool forceComposite) {
  PhotoshopConversionPlan plan;
  const bool escaping = EscapesIsolation(source.layers, plan);
  if (forceComposite || escaping || !source.features.empty() || source.colorMode != "RGB" ||
      source.colorProfile != "sRGB IEC61966-2.1" || source.bitDepth != 8) {
    Composite(plan);
    plan.diagnostics.push_back(
        {"document-composite", "", "Export the Photoshop document composite as 8-bit sRGB."});
    return plan;
  }
  AnalyzeSiblings(source.layers, "", plan);
  // A root range covering every visible sibling is exactly the original document render.
  if (plan.rasters.size() == 1 && plan.rasters[0].parentId.empty() &&
      plan.statistics.vectorShapes == 0 && plan.statistics.outlinedTexts == 0 &&
      plan.rasters[0].id.find("range:") == 0) {
    const auto& ids = plan.rasters[0].layerIds;
    if (std::all_of(source.layers.begin(), source.layers.end(), [&](const auto& layer) {
          return !layer.visible || std::find(ids.begin(), ids.end(), layer.id) != ids.end();
        }))
      Composite(plan);
  }
  return plan;
}
}  // namespace pagx
