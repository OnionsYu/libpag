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
#include "PhotoshopCompositingPlanner.h"
#include "PhotoshopPaintBuilder.h"
#include "PhotoshopShapeBuilder.h"
#include "pagx/nodes/Fill.h"
#include "pagx/nodes/Image.h"
#include "pagx/nodes/ImagePattern.h"
#include "pagx/nodes/Rectangle.h"
#include "pagx/nodes/SolidColor.h"
#include "tgfx/core/Data.h"
#include "tgfx/core/ImageCodec.h"

namespace pagx {
namespace {
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
  Builder(const photoshop::Document& source, const Assets& assets,
          const PhotoshopConversionPlan& plan)
      : source(source), assets(assets),
        document(PAGXDocument::Make(source.bounds.width, source.bounds.height)) {
    for (const auto& request : plan.rasters) captures.emplace(request.layerId, &request);
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
    for (size_t i = 0; i < input.size(); ++i) {
      const auto& item = input[i];
      auto capture = captures.find(item.id);
      if (capture != captures.end()) {
        const auto& request = *capture->second;
        auto layer = document->makeNode<Layer>("ps_layer_" + std::to_string(nextLayer++));
        layer->name = item.name;
        addRaster(layer, request.id);
        output.push_back(layer);
        i += request.layerIds.size() - 1;
        continue;
      }
      if (!item.visible) continue;
      auto layer = document->makeNode<Layer>("ps_layer_" + std::to_string(nextLayer++));
      layer->name = item.name;
      layer->blendMode = NativePhotoshopBlendMode(item.blendMode).value_or(BlendMode::Normal);
      output.push_back(layer);
      if (IsPhotoshopGroup(item)) {
        layer->alpha = item.opacity;
        layer->passThroughBackground = item.blendMode == "passThrough";
        addLayers(item.children, layer->children);
      } else if (IsNativePhotoshopLayer(item) && item.shape) {
        layer->x = -source.bounds.left;
        layer->y = -source.bounds.top;
        layer->alpha = item.opacity;
        BuildPhotoshopShape(document.get(), layer, *item.shape, item.fillOpacity);
      } else if (IsNativePhotoshopLayer(item) && item.fill) {
        layer->alpha = item.opacity;
        layer->x = -source.bounds.left;
        layer->y = -source.bounds.top;
        auto rect = document->makeNode<Rectangle>();
        rect->position = {source.bounds.left + source.bounds.width / 2,
                          source.bounds.top + source.bounds.height / 2};
        rect->size = {source.bounds.width, source.bounds.height};
        auto fill = document->makeNode<Fill>();
        fill->alpha = item.fillOpacity;
        fill->color = BuildPhotoshopPaint(document.get(), *item.fill);
        layer->contents = {rect, fill};
      } else if (IsNativePhotoshopLayer(item) && item.rectangle) {
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
      } else if (!IsEmptyPhotoshopLayer(item)) {
        // Leaf opacity/fill/masks/transforms were already baked into the host capture.
        addRaster(layer, "layer:" + item.id);
      }
    }
  }

  const photoshop::Document& source;
  const Assets& assets;
  std::unordered_map<std::string, const PhotoshopRasterRequest*> captures = {};
  std::shared_ptr<PAGXDocument> document;
  std::vector<std::pair<const photoshop::RasterAsset*, Image*>> images = {};
  size_t nextLayer = 0;
};
}  // namespace

PhotoshopConversionPlan PhotoshopImporter::Analyze(const photoshop::Document& source,
                                                   bool forceComposite) {
  photoshop::Validate(source);
  return PlanPhotoshopComposition(source, forceComposite);
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
  Builder builder(source, assets, plan);
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
