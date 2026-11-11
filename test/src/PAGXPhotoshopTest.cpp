/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include <limits>
#include <stdexcept>
#include "pagx/PAGXExporter.h"
#include "pagx/PAGXImporter.h"
#include "pagx/PhotoshopImporter.h"
#include "pagx/nodes/Image.h"
#include "tgfx/core/ImageCodec.h"
#include "tgfx/core/Pixmap.h"
#include "utils/TestUtils.h"

namespace pag {
namespace {
photoshop::Document PSDocument() {
  photoshop::Document source;
  source.name = "测试 & <图层>";
  source.bounds = {-10, -20, 200, 100};
  return source;
}
photoshop::Layer Pixel(const std::string& id) {
  photoshop::Layer layer;
  layer.id = id;
  layer.name = id;
  layer.kind = photoshop::LayerKind::Pixel;
  layer.blendMode = "normal";
  layer.compositingKnown = true;
  return layer;
}
photoshop::RasterAsset Capture(const std::string& id) {
  // Half-transparent red and opaque blue exercise alpha-preserving PNG embedding.
  const uint8_t pixels[] = {255, 0, 0, 128, 0, 0, 255, 255};
  auto info =
      tgfx::ImageInfo::Make(2, 1, tgfx::ColorType::RGBA_8888, tgfx::AlphaType::Unpremultiplied);
  auto png = tgfx::ImageCodec::Encode(tgfx::Pixmap(info, pixels), tgfx::EncodedFormat::PNG, 100);
  photoshop::RasterAsset asset;
  asset.requestId = id;
  asset.bounds = {-5, -12, 2, 1};
  if (png) asset.png.assign(png->bytes(), png->bytes() + png->size());
  return asset;
}
}  // namespace

PAGX_TEST(PAGXPhotoshopTest, KeepsPaintOrderAndBakesLeafOpacityExactlyOnce) {
  auto source = PSDocument();
  auto group = Pixel("group");
  group.kind = photoshop::LayerKind::Group;
  group.opacity = 0.5f;
  auto bottom = Pixel("bottom");
  bottom.opacity = 0.25f;
  bottom.fillOpacity = 0.3f;
  group.children = {bottom, Pixel("top")};
  source.layers = {group};
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_FALSE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 2u);
  EXPECT_EQ(plan.rasters[0].layerId, "bottom");
  auto document =
      pagx::PhotoshopImporter::Convert(source, {Capture("layer:bottom"), Capture("layer:top")});
  ASSERT_EQ(document->layers.size(), 1u);
  auto outputGroup = document->layers[0];
  EXPECT_FLOAT_EQ(outputGroup->alpha, 0.5f);
  ASSERT_EQ(outputGroup->children.size(), 2u);
  EXPECT_EQ(outputGroup->children[0]->name, "bottom");
  EXPECT_EQ(outputGroup->children[1]->name, "top");
  EXPECT_FLOAT_EQ(outputGroup->children[0]->alpha, 1.0f);
  EXPECT_FLOAT_EQ(outputGroup->children[0]->x, 5);
  EXPECT_FLOAT_EQ(outputGroup->children[0]->y, 8);
  size_t images = 0;
  for (const auto& node : document->nodes) {
    if (node->nodeType() == pagx::NodeType::Image) images++;
  }
  EXPECT_EQ(images, 1u);
  auto xml = pagx::PAGXExporter::ToXML(*document);
  auto roundTrip = pagx::PAGXImporter::FromXML(xml);
  ASSERT_NE(roundTrip, nullptr);
  EXPECT_TRUE(roundTrip->errors.empty());
  EXPECT_TRUE(roundTrip->getExternalFilePaths().empty());
  EXPECT_NE(xml.find("data:image/png;base64,"), std::string::npos);
}

PAGX_TEST(PAGXPhotoshopTest, AdjustmentReplacesBackdropInsteadOfOverlayingIt) {
  auto source = PSDocument();
  auto adjustment = Pixel("adjustment");
  adjustment.kind = photoshop::LayerKind::Adjustment;
  source.layers = {Pixel("background"), Pixel("foreground"), adjustment};
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_TRUE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].id, "composite");
  auto document = pagx::PhotoshopImporter::Convert(source, {Capture("composite")});
  EXPECT_EQ(document->layers.size(), 1u);
  EXPECT_EQ(document->layers[0]->name, source.name);
}

PAGX_TEST(PAGXPhotoshopTest, UnknownMetadataAndAdvancedBlendRequireComposite) {
  for (const auto& mode : {"multiply", "dissolve", "unknown"}) {
    auto source = PSDocument();
    auto layer = Pixel("layer");
    layer.blendMode = mode;
    source.layers = {layer};
    EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).composite);
  }
  auto source = PSDocument();
  auto layer = Pixel("layer");
  layer.compositingKnown = false;
  source.layers = {layer};
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).composite);
  layer.compositingKnown = true;
  layer.features = {"blendIf"};
  source.layers = {layer};
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).composite);
}

PAGX_TEST(PAGXPhotoshopTest, ClippingAndGroupOpacityAreNotIndependentLayerRenders) {
  auto source = PSDocument();
  auto clipped = Pixel("clipped");
  clipped.clippingBase = "base";
  source.layers = {Pixel("base"), clipped};
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).composite);
  auto group = Pixel("group");
  group.kind = photoshop::LayerKind::Group;
  group.blendMode = "passThrough";
  group.children = {Pixel("child")};
  source.layers = {group};
  EXPECT_FALSE(pagx::PhotoshopImporter::Analyze(source).composite);
  source.layers[0].opacity = 0.5f;
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).composite);
}

PAGX_TEST(PAGXPhotoshopTest, HiddenSubtreesDoNotChangeCurrentStatePlan) {
  auto source = PSDocument();
  auto group = Pixel("hidden");
  group.kind = photoshop::LayerKind::Group;
  group.visible = false;
  auto unsupported = Pixel("unsupported");
  unsupported.compositingKnown = false;
  group.children = {unsupported};
  source.layers = {group, Pixel("visible")};
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  EXPECT_FALSE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].layerId, "visible");
}

PAGX_TEST(PAGXPhotoshopTest, RejectsMalformedSourceAndCaptureContracts) {
  auto source = PSDocument();
  source.layers = {Pixel("same"), Pixel("same")};
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers = {Pixel("one")};
  source.layers[0].opacity = std::numeric_limits<float>::quiet_NaN();
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers[0].opacity = 1;
  source.layers[0].clippingBase = "missing";
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers[0].clippingBase.clear();
  EXPECT_THROW(pagx::PhotoshopImporter::Convert(source, {}), std::invalid_argument);
  EXPECT_THROW(pagx::PhotoshopImporter::Convert(source, {Capture("wrong")}), std::invalid_argument);
  auto asset = Capture("layer:one");
  asset.bounds.width = 3;
  EXPECT_THROW(pagx::PhotoshopImporter::Convert(source, {asset}), std::invalid_argument);
  asset = Capture("layer:one");
  asset.png.resize(7);
  EXPECT_THROW(pagx::PhotoshopImporter::Convert(source, {asset}), std::invalid_argument);
  source.schemaVersion = 2;
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
}

PAGX_TEST(PAGXPhotoshopTest, ConvertsNativeRectangleWithoutHostRasterAndSupportsEmptyCanvas) {
  auto source = PSDocument();
  auto shape = Pixel("shape");
  shape.kind = photoshop::LayerKind::Shape;
  photoshop::SolidRectangle rectangle;
  rectangle.bounds = {-5, -10, 20, 30};
  rectangle.red = 1;
  shape.rectangle = rectangle;
  source.layers = {shape};
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).rasters.empty());
  auto document = pagx::PhotoshopImporter::Convert(source, {});
  ASSERT_EQ(document->layers.size(), 1u);
  EXPECT_FLOAT_EQ(document->layers[0]->x, 5);
  EXPECT_FLOAT_EQ(document->layers[0]->y, 10);
  auto parsed = pagx::PAGXImporter::FromXML(pagx::PAGXExporter::ToXML(*document));
  ASSERT_NE(parsed, nullptr);
  EXPECT_TRUE(parsed->errors.empty());
  source.layers.clear();
  document = pagx::PhotoshopImporter::Convert(source, {});
  EXPECT_TRUE(document->layers.empty());
  auto empty = photoshop::RasterAsset{"composite", {}, {}};
  document = pagx::PhotoshopImporter::Convert(source, {empty}, true);
  EXPECT_TRUE(document->layers[0]->contents.empty());
}
}  // namespace pag
