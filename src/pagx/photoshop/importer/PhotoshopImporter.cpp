/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include "pagx/PhotoshopImporter.h"
#include <cmath>
#include <stdexcept>
#include <unordered_map>
#include "pagx/nodes/Fill.h"
#include "pagx/nodes/Image.h"
#include "pagx/nodes/ImagePattern.h"
#include "pagx/nodes/Rectangle.h"
#include "pagx/nodes/SolidColor.h"
#include "tgfx/core/Data.h"
#include "tgfx/core/ImageCodec.h"

namespace pagx {
namespace {
bool IsGroup(const photoshop::Layer& layer) {
  return layer.kind == photoshop::LayerKind::Group;
}
void AnalyzeLayers(const std::vector<photoshop::Layer>& layers, PhotoshopConversionPlan& plan) {
  for (const auto& layer : layers) {
    if (!layer.visible) continue;
    bool safeBlend = layer.blendMode == "normal" ||
                     (IsGroup(layer) && layer.blendMode == "passThrough" && layer.opacity == 1);
    bool safeKind = layer.kind != photoshop::LayerKind::Unknown &&
                    layer.kind != photoshop::LayerKind::Adjustment &&
                    layer.kind != photoshop::LayerKind::Artboard;
    if (!layer.compositingKnown || !safeBlend || !safeKind || !layer.clippingBase.empty() ||
        !layer.features.empty() || (IsGroup(layer) && layer.fillOpacity != 1)) {
      plan.composite = true;
      plan.diagnostics.push_back(
          {"composite-required", layer.id,
           "Compositing dependencies require the Photoshop document render."});
    }
    if (IsGroup(layer)) {
      AnalyzeLayers(layer.children, plan);
    } else if (!layer.rectangle) {
      plan.rasters.push_back({"layer:" + layer.id, layer.id});
      if (layer.kind != photoshop::LayerKind::Pixel) {
        plan.diagnostics.push_back(
            {"rasterized-content", layer.id,
             "Photoshop renders this layer; its content is not editable in PAGX."});
      }
    }
  }
}

using Assets = std::unordered_map<std::string, const photoshop::RasterAsset*>;

void ValidateAsset(const photoshop::RasterAsset& asset) {
  const auto& b = asset.bounds;
  if (!std::isfinite(b.left) || !std::isfinite(b.top) || !std::isfinite(b.width) ||
      !std::isfinite(b.height) || b.width < 0 || b.height < 0) {
    throw std::invalid_argument("Invalid raster bounds: " + asset.requestId);
  }
  if (asset.png.empty() && b.width == 0 && b.height == 0) return;
  if (b.width < 1 || b.height < 1 || b.width > 300000 || b.height > 300000 ||
      std::floor(b.width) != b.width || std::floor(b.height) != b.height) {
    throw std::invalid_argument("Invalid raster dimensions: " + asset.requestId);
  }
  static constexpr uint8_t Signature[] = {137, 80, 78, 71, 13, 10, 26, 10};
  if (asset.png.size() < sizeof(Signature) ||
      !std::equal(std::begin(Signature), std::end(Signature), asset.png.begin())) {
    throw std::invalid_argument("Raster must contain PNG data: " + asset.requestId);
  }
  auto codec =
      tgfx::ImageCodec::MakeFrom(tgfx::Data::MakeWithoutCopy(asset.png.data(), asset.png.size()));
  if (!codec || codec->width() != b.width || codec->height() != b.height) {
    throw std::invalid_argument("PNG dimensions do not match capture bounds: " + asset.requestId);
  }
}

class Builder {
 public:
  Builder(const photoshop::Document& source, const Assets& assets)
      : source(source), assets(assets),
        document(PAGXDocument::Make(source.bounds.width, source.bounds.height)) {
  }

  void addRaster(Layer* layer, const std::string& requestId) {
    const auto& asset = *assets.at(requestId);
    if (asset.png.empty()) return;
    Image* image = nullptr;
    // Resource equality includes the bytes, so even a future hash-based index must check collisions.
    for (const auto& entry : images) {
      if (entry.first->png == asset.png) {
        image = entry.second;
        break;
      }
    }
    if (!image) {
      image = document->makeNode<Image>("ps_image_" + std::to_string(images.size()));
      image->data = Data::MakeWithCopy(asset.png.data(), asset.png.size());
      images.push_back({&asset, image});
    }
    layer->x = asset.bounds.left - source.bounds.left;
    layer->y = asset.bounds.top - source.bounds.top;
    auto rect = document->makeNode<Rectangle>();
    rect->size = {asset.bounds.width, asset.bounds.height};
    auto pattern = document->makeNode<ImagePattern>();
    pattern->image = image;
    pattern->scaleMode = ScaleMode::None;
    auto fill = document->makeNode<Fill>();
    fill->color = pattern;
    layer->contents = {rect, fill};
  }

  void addLayers(const std::vector<photoshop::Layer>& input, std::vector<Layer*>& output) {
    for (const auto& item : input) {
      if (!item.visible) continue;
      auto layer = document->makeNode<Layer>("ps_layer_" + std::to_string(nextLayer++));
      layer->name = item.name;
      output.push_back(layer);
      if (IsGroup(item)) {
        layer->alpha = item.opacity;
        layer->passThroughBackground = false;
        addLayers(item.children, layer->children);
      } else if (item.rectangle) {
        const auto& value = *item.rectangle;
        layer->x = value.bounds.left - source.bounds.left;
        layer->y = value.bounds.top - source.bounds.top;
        layer->alpha = item.opacity;
        auto rect = document->makeNode<Rectangle>();
        rect->size = {value.bounds.width, value.bounds.height};
        rect->roundness = value.roundness;
        auto color = document->makeNode<SolidColor>();
        color->color = {value.red, value.green, value.blue, value.alpha};
        auto fill = document->makeNode<Fill>();
        fill->alpha = item.fillOpacity;
        fill->color = color;
        layer->contents = {rect, fill};
      } else {
        // Leaf opacity/fill/masks/transforms were already baked into the host capture.
        addRaster(layer, "layer:" + item.id);
      }
    }
  }

  const photoshop::Document& source;
  const Assets& assets;
  std::shared_ptr<PAGXDocument> document;
  std::vector<std::pair<const photoshop::RasterAsset*, Image*>> images = {};
  size_t nextLayer = 0;
};
}  // namespace

PhotoshopConversionPlan PhotoshopImporter::Analyze(const photoshop::Document& source,
                                                   bool forceComposite) {
  photoshop::Validate(source);
  PhotoshopConversionPlan plan;
  if (forceComposite || !source.features.empty() || source.colorMode != "RGB" ||
      source.colorProfile != "sRGB IEC61966-2.1" || source.bitDepth != 8) {
    plan.composite = true;
    plan.diagnostics.push_back(
        {"document-composite", "", "Export the Photoshop document composite."});
  }
  AnalyzeLayers(source.layers, plan);
  if (plan.composite) {
    plan.rasters = {{"composite", ""}};
  }
  return plan;
}

std::shared_ptr<PAGXDocument> PhotoshopImporter::Convert(
    const photoshop::Document& source, const std::vector<photoshop::RasterAsset>& inputAssets,
    bool forceComposite) {
  auto plan = Analyze(source, forceComposite);
  if (inputAssets.size() != plan.rasters.size()) {
    throw std::invalid_argument("Capture count does not match the conversion plan");
  }
  Assets assets;
  for (const auto& asset : inputAssets) {
    ValidateAsset(asset);
    if (!assets.emplace(asset.requestId, &asset).second) {
      throw std::invalid_argument("Duplicate capture: " + asset.requestId);
    }
  }
  for (const auto& request : plan.rasters) {
    if (assets.count(request.id) == 0)
      throw std::invalid_argument("Missing capture: " + request.id);
  }
  Builder builder(source, assets);
  if (plan.composite) {
    auto layer = builder.document->makeNode<Layer>("ps_composite");
    layer->name = source.name;
    builder.addRaster(layer, "composite");
    builder.document->layers.push_back(layer);
  } else {
    builder.addLayers(source.layers, builder.document->layers);
  }
  return builder.document;
}
}  // namespace pagx
