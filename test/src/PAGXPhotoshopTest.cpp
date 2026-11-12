/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include <limits>
#include <nlohmann/json.hpp>
#include <stdexcept>
#include "pagx/PAGXExporter.h"
#include "pagx/PAGXImporter.h"
#include "pagx/PhotoshopImporter.h"
#include "pagx/nodes/Fill.h"
#include "pagx/nodes/Image.h"
#include "pagx/nodes/Path.h"
#include "pagx/nodes/Rectangle.h"
#include "pagx/nodes/Stroke.h"
#include "renderer/LayerBuilder.h"
#include "renderer/ToTGFX.h"
#include "tgfx/core/ImageCodec.h"
#include "tgfx/core/Pixmap.h"
#include "tgfx/core/Surface.h"
#include "tgfx/layers/DisplayList.h"
#include "tgfx/layers/Layer.h"
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
  layer.bounds = {0, 0, 2, 1};
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

static photoshop::PathComponent PSBox(float left, float top, float right, float bottom) {
  photoshop::PathComponent component;
  photoshop::Subpath path;
  for (auto point : {photoshop::Point{left, top}, {right, top}, {right, bottom}, {left, bottom}})
    path.points.push_back({point, point, point});
  component.subpaths = {path};
  return component;
}
static photoshop::Layer NativeShape(const std::string& id) {
  auto layer = Pixel(id);
  layer.kind = photoshop::LayerKind::Shape;
  photoshop::Shape shape;
  shape.components = {PSBox(0, 0, 100, 100)};
  shape.fill = photoshop::Color{0, 1, 0, 1};
  layer.shape = shape;
  return layer;
}

PAGX_TEST(PAGXPhotoshopTest, VersionedPaintRetainsSourceSemanticsAndFillsTheExportRectangle) {
  const std::string text = R"({"schemaVersion":3,"name":"fill","hostVersion":"27.4.0",
    "resolution":72,"colorMode":"RGB","colorProfile":"sRGB IEC61966-2.1","bitDepth":8,
    "bounds":{"left":-10,"top":-20,"width":200,"height":100},"features":[],"layers":[{
    "id":"fill","sourceLayerId":1,"name":"fill","kind":"fill","visible":true,
    "opacity":0.75,"fillOpacity":0.6,"blendMode":"normal","compositingKnown":true,
    "clippingBase":"","features":[],"rasterFeatures":[],"text":"","children":[],
    "bounds":{"left":50,"top":60,"width":2,"height":3},
    "fill":{"kind":"solid","color":{"red":1,"green":0,"blue":0,"alpha":1}}}]})";
  auto source = photoshop::ParseSnapshot(text);
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).rasters.empty());
  auto output = pagx::PhotoshopImporter::Convert(source, {});
  auto layer = output->layers.front();
  auto rect = static_cast<pagx::Rectangle*>(layer->contents[0]);
  EXPECT_FLOAT_EQ(layer->x + rect->position.x - rect->size.width / 2, 0);
  EXPECT_FLOAT_EQ(layer->y + rect->position.y - rect->size.height / 2, 0);
  EXPECT_FLOAT_EQ(rect->size.width, 200);
  EXPECT_FLOAT_EQ(rect->size.height, 100);
  EXPECT_FLOAT_EQ(layer->alpha, 0.75);
  EXPECT_FLOAT_EQ(static_cast<pagx::Fill*>(layer->contents[1])->alpha, 0.6);
  auto json = nlohmann::json::parse(text);
  json["schemaVersion"] = 2;
  EXPECT_THROW(photoshop::ParseSnapshot(json.dump()), std::invalid_argument);
  json["schemaVersion"] = 3;
  json["layers"][0]["fill"]["kind"] = "future";
  EXPECT_THROW(photoshop::ParseSnapshot(json.dump()), std::invalid_argument);
  json = nlohmann::json::parse(text);
  json["layers"][0]["vectorMask"] = {{"enabled", false},
                                     {"linked", true},
                                     {"inverted", false},
                                     {"initialFill", false},
                                     {"density", 1},
                                     {"feather", 0},
                                     {"components", nlohmann::json::array()}};
  EXPECT_THROW(photoshop::ParseSnapshot(json.dump()), std::invalid_argument);
  json["schemaVersion"] = 4;
  EXPECT_FALSE(photoshop::ParseSnapshot(json.dump()).layers[0].vectorMask->enabled);
  json["layers"][0]["vectorMask"]["density"] = -1;
  EXPECT_THROW(photoshop::ParseSnapshot(json.dump()), std::invalid_argument);
  photoshop::Gradient gradient;
  gradient.interpolation = "perceptual";
  gradient.colors = {{0, 0.5f, {1, 0, 0, 1}}, {1, 0.5f, {0, 0, 1, 1}}};
  gradient.opacity = {{0, 0.3f, 1}, {1, 0.5f, 0}};
  source.layers[0].fill->gradient = gradient;
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].layerId, "fill");
  source.layers[0].fill->gradient->matrix = {0, 0, 0, 0, 0, 0};
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers[0].fill->gradient = gradient;
  source.layers[0].fill->gradient->colors[1].position = -1;
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers = {NativeShape("stroke-opacity")};
  source.layers[0].shape->stroke = photoshop::Stroke{};
  source.layers[0].shape->stroke->opacity = 89.0f / 255;
  EXPECT_EQ(pagx::PhotoshopImporter::Analyze(source).rasters.size(), 1u);
  source.layers[0].shape->stroke->opacity = 1;
  source.layers[0].shape->stroke->cap = "round";
  source.layers[0].shape->stroke->width = 6;
  source.layers[0].shape->stroke->dashes = {3, 2};
  EXPECT_EQ(pagx::PhotoshopImporter::Analyze(source).rasters.size(), 1u);
  source.layers[0].shape->stroke->dashes = {12, 8};
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).rasters.empty());
}

PAGX_TEST(PAGXPhotoshopTest, CompoundStrokeAlignmentSurvivesSerializationAtMultipleScales) {
  DeviceLockScope deviceLock;
  ASSERT_TRUE(static_cast<bool>(deviceLock));
  for (bool booleanHole : {false, true}) {
    for (const auto& alignment : {"inner", "outer"}) {
      auto source = PSDocument();
      source.bounds = {0, 0, 100, 100};
      auto layer = NativeShape("compound");
      auto outer = PSBox(10, 10, 90, 90);
      auto hole = PSBox(30, 30, 70, 70);
      if (booleanHole) {
        hole.operation = photoshop::PathOperation::Subtract;
        layer.shape->components = {outer, hole};
      } else {
        outer.subpaths.push_back(hole.subpaths.front());
        layer.shape->components = {outer};
      }
      layer.shape->fill.reset();
      layer.shape->stroke = photoshop::Stroke{};
      layer.shape->stroke->width = 4;
      layer.shape->stroke->align = alignment;
      layer.shape->stroke->paint = photoshop::Color{1, 0, 0, 1};
      source.layers = {layer};
      auto converted = pagx::PhotoshopImporter::Convert(source, {});
      auto xml = pagx::PAGXExporter::ToXML(*converted);
      EXPECT_EQ(xml.find("<Image"), std::string::npos);
      auto parsed = pagx::PAGXImporter::FromXML(xml);
      ASSERT_NE(parsed, nullptr);
      ASSERT_TRUE(parsed->errors.empty());
      parsed->applyLayout();
      for (int scale : {1, 4}) {
        auto root = pagx::LayerBuilder::Build(parsed.get());
        root->setMatrix(tgfx::Matrix::MakeScale(scale));
        auto surface = tgfx::Surface::Make(deviceLock.context(), 100 * scale, 100 * scale);
        ASSERT_NE(surface, nullptr);
        tgfx::DisplayList display;
        display.root()->addChild(root);
        display.render(surface.get(), false);
        std::vector<uint8_t> pixels(100 * 100 * scale * scale * 4);
        ASSERT_TRUE(surface->readPixels(
            tgfx::ImageInfo::Make(100 * scale, 100 * scale, tgfx::ColorType::RGBA_8888,
                                  tgfx::AlphaType::Unpremultiplied),
            pixels.data()));
        for (int x : {5, 7, 9, 11, 13, 15, 25, 27, 29, 31, 33, 35, 50}) {
          const bool inner = std::string(alignment) == "inner";
          const bool painted = inner ? (x > 10 && x < 14) || (x > 26 && x < 30)
                                     : (x > 6 && x < 10) || (x > 30 && x < 34);
          const auto offset = ((50 * scale) * (100 * scale) + x * scale) * 4;
          EXPECT_EQ(pixels[offset + 3], painted ? 255 : 0)
              << "x=" << x << " scale=" << scale << " " << alignment << " boolean=" << booleanHole;
        }
      }
    }
  }
}

PAGX_TEST(PAGXPhotoshopTest, IndependentVectorMasksPreserveHolesCoordinatesAndSingleOpacity) {
  DeviceLockScope deviceLock;
  ASSERT_TRUE(static_cast<bool>(deviceLock));
  for (int variant = 0; variant < 4; ++variant) {
    const bool groupOwner = variant & 1;
    const bool inverted = variant & 2;
    auto source = PSDocument();
    source.bounds = {-10, -20, 100, 100};
    auto shape = NativeShape("shape");
    shape.shape->components = {PSBox(-10, -20, 90, 80)};
    shape.opacity = 0.5f;
    auto owner = groupOwner ? Pixel("group") : shape;
    if (groupOwner) {
      owner.kind = photoshop::LayerKind::Group;
      owner.opacity = 0.5f;
      owner.children = {shape};
    } else
      owner.fillOpacity = 0.5f;
    photoshop::VectorMask mask;
    mask.inverted = inverted;
    auto hole = PSBox(20, 20, 40, 40);
    hole.operation = photoshop::PathOperation::Subtract;
    mask.components = {PSBox(0, 0, 80, 60), hole};
    owner.vectorMask = mask;
    source.layers = {owner};
    EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).rasters.empty());
    auto converted = pagx::PhotoshopImporter::Convert(source, {});
    EXPECT_NE(converted->layers[0]->mask, nullptr);
    auto xml = pagx::PAGXExporter::ToXML(*converted);
    EXPECT_EQ(xml.find("<Image"), std::string::npos);
    auto parsed = pagx::PAGXImporter::FromXML(xml);
    ASSERT_NE(parsed, nullptr);
    ASSERT_TRUE(parsed->errors.empty());
    parsed->applyLayout();
    for (int scale : {1, 4}) {
      auto root = pagx::LayerBuilder::Build(parsed.get());
      root->setMatrix(tgfx::Matrix::MakeScale(scale));
      auto surface = tgfx::Surface::Make(deviceLock.context(), 100 * scale, 100 * scale);
      ASSERT_NE(surface, nullptr);
      tgfx::DisplayList display;
      display.root()->addChild(root);
      display.render(surface.get(), false);
      std::vector<uint8_t> pixels(100 * 100 * scale * scale * 4);
      ASSERT_TRUE(surface->readPixels(
          tgfx::ImageInfo::Make(100 * scale, 100 * scale, tgfx::ColorType::RGBA_8888,
                                tgfx::AlphaType::Unpremultiplied),
          pixels.data()));
      for (auto point : {photoshop::Point{15, 25}, {35, 45}, {3, 5}, {65, 65}, {95, 95}}) {
        const auto offset =
            (static_cast<int>(point.y) * scale * 100 * scale + static_cast<int>(point.x) * scale) *
            4;
        const bool painted = (point.x == 15 || point.x == 65) != inverted;
        EXPECT_NEAR(pixels[offset + 3], painted ? 64 : 0, 1)
            << "group=" << groupOwner << " scale=" << scale;
      }
    }
    source.layers[0].vectorMask->feather = 3;
    auto plan = pagx::PhotoshopImporter::Analyze(source);
    ASSERT_EQ(plan.rasters.size(), 1u);
    EXPECT_EQ(plan.rasters[0].layerId, owner.id);
    source.layers[0].vectorMask->enabled = false;
    EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).rasters.empty());
    converted = pagx::PhotoshopImporter::Convert(source, {});
    EXPECT_EQ(converted->layers[0]->mask, nullptr);
    source.schemaVersion = 3;
    EXPECT_THROW(photoshop::Validate(source), std::invalid_argument);
  }
  auto source = PSDocument();
  auto group = Pixel("pass");
  group.kind = photoshop::LayerKind::Group;
  group.blendMode = "passThrough";
  group.children = {NativeShape("inside")};
  group.vectorMask = photoshop::VectorMask{};
  group.vectorMask->components = {PSBox(0, 0, 100, 100)};
  source.layers = {NativeShape("backdrop"), group, NativeShape("above")};
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].layerIds, (std::vector<std::string>{"backdrop", "pass"}));
}

PAGX_TEST(PAGXPhotoshopTest, UnqualifiedDocumentCompositionBakesNormalVectorsOnce) {
  auto source = PSDocument();
  source.layers = {NativeShape("normal")};
  source.features = {"compositing-environment:rgb-blend-gamma"};
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  EXPECT_TRUE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].id, "composite");
  auto output = pagx::PhotoshopImporter::Convert(source, {Capture("composite")});
  ASSERT_EQ(output->layers.size(), 1u);
  EXPECT_FLOAT_EQ(output->layers[0]->alpha, 1);
}

PAGX_TEST(PAGXPhotoshopTest, LocalBackdropReplacementStopsAtAnIsolatedGroup) {
  auto source = PSDocument();
  auto group = Pixel("group");
  group.kind = photoshop::LayerKind::Group;
  group.opacity = 0.5f;
  auto adjustment = Pixel("adjustment");
  adjustment.kind = photoshop::LayerKind::Adjustment;
  group.children = {NativeShape("background"), adjustment, NativeShape("above")};
  source.layers = {NativeShape("outside"), group};
  const auto plan = pagx::PhotoshopImporter::Analyze(source);
  EXPECT_FALSE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].parentId, "group");
  EXPECT_EQ(plan.rasters[0].layerIds, (std::vector<std::string>{"background", "adjustment"}));
  EXPECT_EQ(plan.statistics.vectorShapes, 2u);
  auto document = pagx::PhotoshopImporter::Convert(source, {Capture(plan.rasters[0].id)});
  ASSERT_EQ(document->layers.size(), 2u);
  auto outputGroup = document->layers[1];
  EXPECT_FLOAT_EQ(outputGroup->alpha, 0.5f);
  ASSERT_EQ(outputGroup->children.size(), 2u);
  EXPECT_FLOAT_EQ(outputGroup->children[0]->alpha, 1.0f);
  EXPECT_EQ(outputGroup->children[1]->name, "above");
}

PAGX_TEST(PAGXPhotoshopTest, NativeBlendsKeepTheirIsolationAndCapturedClipsIncludeTheBackdrop) {
  auto source = PSDocument();
  auto multiply = NativeShape("multiply");
  multiply.blendMode = "multiply";
  auto group = Pixel("group");
  group.kind = photoshop::LayerKind::Group;
  group.children = {multiply};
  group.blendMode = "passThrough";
  source.layers = {NativeShape("backdrop"), group};
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).rasters.empty());
  auto output = pagx::PhotoshopImporter::Convert(source, {});
  EXPECT_TRUE(output->layers[1]->passThroughBackground);
  EXPECT_EQ(output->layers[1]->children[0]->blendMode, pagx::BlendMode::Multiply);
  source.layers[1].blendMode = "normal";
  source.layers[1].opacity = 0.5f;
  output = pagx::PhotoshopImporter::Convert(source, {});
  EXPECT_FALSE(output->layers[1]->passThroughBackground);
  EXPECT_FLOAT_EQ(output->layers[1]->alpha, 0.5f);
  auto clip = NativeShape("clipped");
  clip.clippingBase = multiply.id;
  source.layers = {NativeShape("backdrop"), multiply, clip, NativeShape("above")};
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_FALSE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].layerIds,
            (std::vector<std::string>{"backdrop", "multiply", "clipped"}));
  group.children = {multiply, clip};
  source.layers = {NativeShape("backdrop"), group, NativeShape("above")};
  plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_TRUE(plan.rasters[0].parentId.empty());
  EXPECT_EQ(plan.rasters[0].layerIds, (std::vector<std::string>{"backdrop", "group"}));
  multiply.blendMode = "softLight";
  source.layers = {NativeShape("backdrop"), multiply, NativeShape("above")};
  plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].layerIds, (std::vector<std::string>{"backdrop", "multiply"}));
}

PAGX_TEST(PAGXPhotoshopTest, PassThroughDependenciesIncludeTheOutsideBackdrop) {
  auto source = PSDocument();
  auto group = Pixel("pass");
  group.kind = photoshop::LayerKind::Group;
  group.blendMode = "passThrough";
  auto multiply = Pixel("multiply");
  multiply.blendMode = "multiply";
  group.children = {NativeShape("inside"), multiply};
  source.layers = {NativeShape("behind"), group, NativeShape("above")};
  const auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_FALSE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_TRUE(plan.rasters[0].parentId.empty());
  EXPECT_EQ(plan.rasters[0].layerIds, (std::vector<std::string>{"behind", "pass"}));
  EXPECT_EQ(plan.statistics.vectorShapes, 1u);
  auto document = pagx::PhotoshopImporter::Convert(source, {Capture(plan.rasters[0].id)});
  ASSERT_EQ(document->layers.size(), 2u);
  EXPECT_EQ(document->layers[1]->name, "above");
}

PAGX_TEST(PAGXPhotoshopTest, OverlappingClippingRangesMergeWithoutDuplicatingPaint) {
  auto source = PSDocument();
  auto first = NativeShape("clip1");
  first.clippingBase = "base";
  auto second = NativeShape("clip2");
  second.clippingBase = "base";
  source.layers = {NativeShape("unrelated"), NativeShape("base"), first, second,
                   NativeShape("above")};
  const auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_FALSE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].layerIds, (std::vector<std::string>{"base", "clip1", "clip2"}));
  auto document = pagx::PhotoshopImporter::Convert(source, {Capture(plan.rasters[0].id)});
  ASSERT_EQ(document->layers.size(), 3u);
  EXPECT_EQ(document->layers[0]->name, "unrelated");
  EXPECT_EQ(document->layers[2]->name, "above");
}

PAGX_TEST(PAGXPhotoshopTest, GroupMasksBakeGroupOpacityOnceAndKnockoutEscapesIsolation) {
  auto source = PSDocument();
  auto group = Pixel("masked");
  group.kind = photoshop::LayerKind::Group;
  group.opacity = 0.3f;
  group.rasterFeatures = {"pixel-mask"};
  group.children = {NativeShape("inside")};
  source.layers = {group, NativeShape("outside")};
  const auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_FALSE(plan.composite);
  ASSERT_EQ(plan.rasters.size(), 1u);
  EXPECT_EQ(plan.rasters[0].layerIds, (std::vector<std::string>{"masked"}));
  auto document = pagx::PhotoshopImporter::Convert(source, {Capture(plan.rasters[0].id)});
  EXPECT_FLOAT_EQ(document->layers[0]->alpha, 1.0f);
  source.layers[0].children[0].features = {"blend:knockout"};
  EXPECT_TRUE(pagx::PhotoshopImporter::Analyze(source).composite);
}

PAGX_TEST(PAGXPhotoshopTest, RejectsAmbiguousClippingChainsAndInvalidNativePathOperations) {
  auto source = PSDocument();
  auto clipped = NativeShape("clipped");
  clipped.clippingBase = "base";
  source.layers = {NativeShape("base"), NativeShape("intervening"), clipped};
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers = {NativeShape("base"), clipped, NativeShape("second")};
  source.layers.back().clippingBase = "clipped";
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers.back().clippingBase = "base";
  EXPECT_NO_THROW(pagx::PhotoshopImporter::Analyze(source));
  source.layers[0].shape->components[0].operation = static_cast<photoshop::PathOperation>(99);
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
}

PAGX_TEST(PAGXPhotoshopTest, PreservesVectorBooleanGeometryAndOpacityWithoutImages) {
  auto source = PSDocument();
  auto layer = Pixel("shape");
  layer.kind = photoshop::LayerKind::Shape;
  layer.opacity = 0.5f;
  layer.fillOpacity = 0.4f;
  photoshop::Shape shape;
  auto outer = PSBox(-5.5f, 10, 100, 80);
  auto hole = PSBox(10, 20, 50, 60);
  hole.operation = photoshop::PathOperation::Subtract;
  shape.components = {outer, hole};
  shape.fill = photoshop::Color{1, 0, 0, 1};
  shape.stroke = photoshop::Stroke{};
  layer.shape = shape;
  source.layers = {layer};
  auto plan = pagx::PhotoshopImporter::Analyze(source);
  ASSERT_FALSE(plan.composite);
  EXPECT_TRUE(plan.rasters.empty());
  auto document = pagx::PhotoshopImporter::Convert(source, {});
  auto output = document->layers[0];
  EXPECT_FLOAT_EQ(output->alpha, 0.5f);
  EXPECT_FLOAT_EQ(static_cast<pagx::Fill*>(output->contents[1])->alpha, 0.4f);
  EXPECT_FLOAT_EQ(output->x, 10);
  EXPECT_FLOAT_EQ(output->y, 20);
  ASSERT_EQ(output->contents[0]->nodeType(), pagx::NodeType::Path);
  auto path = ToTGFX(*static_cast<pagx::Path*>(output->contents[0])->data);
  if (static_cast<pagx::Fill*>(output->contents[1])->fillRule == pagx::FillRule::EvenOdd)
    path.setFillType(tgfx::PathFillType::EvenOdd);
  EXPECT_TRUE(path.contains(0, 15));
  EXPECT_FALSE(path.contains(20, 30));
  EXPECT_TRUE(path.contains(60, 30));
  auto xml = pagx::PAGXExporter::ToXML(*document);
  EXPECT_EQ(xml.find("<Image"), std::string::npos);
  auto parsed = pagx::PAGXImporter::FromXML(xml);
  ASSERT_NE(parsed, nullptr);
  EXPECT_TRUE(parsed->errors.empty());
}

PAGX_TEST(PAGXPhotoshopTest, PreservesClosingBezierHandlesAndRejectsInvalidGeometry) {
  auto source = PSDocument();
  auto layer = Pixel("curve");
  layer.kind = photoshop::LayerKind::Shape;
  photoshop::Shape shape;
  auto component = PSBox(0, 0, 40, 40);
  component.subpaths[0].points[0].incoming = {-20, 0};
  component.subpaths[0].points.back().outgoing = {-20, 40};
  shape.components = {component};
  shape.fill = photoshop::Color{0, 0, 1, 1};
  layer.shape = shape;
  source.layers = {layer};
  auto document = pagx::PhotoshopImporter::Convert(source, {});
  auto path = ToTGFX(*static_cast<pagx::Path*>(document->layers[0]->contents[0])->data);
  EXPECT_TRUE(path.contains(-10, 20));
  EXPECT_FALSE(path.contains(-21, 20));
  source.layers[0].shape->components[0].subpaths[0].points[0].incoming.x =
      std::numeric_limits<float>::infinity();
  EXPECT_THROW(pagx::PhotoshopImporter::Analyze(source), std::invalid_argument);
  source.layers[0].shape = shape;
  source.layers[0].shape->components[0].operation = photoshop::PathOperation::Subtract;
  EXPECT_EQ(pagx::PhotoshopImporter::Analyze(source).rasters.size(), 1u);
  source.layers[0].shape = shape;
  source.layers[0].rasterFeatures = {"pixel-mask"};
  EXPECT_EQ(pagx::PhotoshopImporter::Analyze(source).rasters.size(), 1u);
}

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
  source.schemaVersion = 5;
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

PAGX_TEST(PAGXPhotoshopTest, ParsesVersionedSnapshotAndRejectsAmbiguousInput) {
  const std::string json = R"({"schemaVersion":1,"name":"中文 & <document>",
    "hostVersion":"27.4.0","resolution":72,"colorMode":"RGB",
    "colorProfile":"sRGB IEC61966-2.1","bitDepth":8,
    "bounds":{"left":0,"top":0,"width":200,"height":100},"features":[],"layers":[]})";
  auto source = photoshop::ParseSnapshot(json);
  EXPECT_EQ(source.name, "中文 & <document>");
  EXPECT_FLOAT_EQ(source.bounds.width, 200);
  EXPECT_TRUE(source.layers.empty());
  auto future = json;
  future.replace(future.find("\"schemaVersion\":1"), 17, "\"schemaVersion\":5");
  EXPECT_THROW(photoshop::ParseSnapshot(future), std::invalid_argument);
  auto fractional = json;
  fractional.replace(fractional.find("\"bitDepth\":8"), 12, "\"bitDepth\":8.5");
  EXPECT_THROW(photoshop::ParseSnapshot(fractional), std::invalid_argument);
  auto unknown = json;
  unknown.insert(1, "\"futureField\":true,");
  EXPECT_THROW(photoshop::ParseSnapshot(unknown), std::invalid_argument);
  auto huge = json;
  huge.replace(huge.find("\"bitDepth\":8"), 12, "\"bitDepth\":18446744073709551615");
  EXPECT_THROW(photoshop::ParseSnapshot(huge), std::invalid_argument);
  for (const auto& replacement : {"true", "false", "null", "\"72\"", "1e100"}) {
    auto invalidNumber = json;
    invalidNumber.replace(invalidNumber.find("\"resolution\":72"), 15,
                          std::string("\"resolution\":") + replacement);
    EXPECT_THROW(photoshop::ParseSnapshot(invalidNumber), std::invalid_argument);
  }
  EXPECT_THROW(photoshop::ParseSnapshot("{}"), std::invalid_argument);
}
}  // namespace pag
