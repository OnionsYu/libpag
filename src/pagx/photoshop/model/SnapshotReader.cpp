/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include <nlohmann/json.hpp>
#include <stdexcept>
#include <unordered_map>
#include <unordered_set>
#include "photoshop/Document.h"

namespace photoshop {
namespace {
using JSON = nlohmann::json;
void Keys(const JSON& value, const std::unordered_set<std::string>& allowed) {
  if (!value.is_object()) throw std::invalid_argument("Expected a snapshot object");
  for (auto it = value.begin(); it != value.end(); ++it) {
    if (!allowed.count(it.key()))
      throw std::invalid_argument("Unknown snapshot field: " + it.key());
  }
}
int Integer(const JSON& value) {
  if (!value.is_number_integer() ||
      (value.is_number_unsigned() && value.get<uint64_t>() > INT32_MAX) ||
      value.get<int64_t>() < -1 || value.get<int64_t>() > INT32_MAX) {
    throw std::invalid_argument("Invalid snapshot integer");
  }
  return value.get<int>();
}
Rect Bounds(const JSON& value) {
  Keys(value, {"left", "top", "width", "height"});
  return {value.at("left").get<float>(), value.at("top").get<float>(),
          value.at("width").get<float>(), value.at("height").get<float>()};
}
LayerKind Kind(const std::string& value) {
  static const std::unordered_map<std::string, LayerKind> kinds = {
      {"unknown", LayerKind::Unknown},
      {"pixel", LayerKind::Pixel},
      {"group", LayerKind::Group},
      {"text", LayerKind::Text},
      {"shape", LayerKind::Shape},
      {"fill", LayerKind::Fill},
      {"smartObject", LayerKind::SmartObject},
      {"adjustment", LayerKind::Adjustment},
      {"artboard", LayerKind::Artboard}};
  auto found = kinds.find(value);
  if (found == kinds.end()) throw std::invalid_argument("Unknown layer kind: " + value);
  return found->second;
}
std::vector<Layer> Layers(const JSON& value) {
  if (!value.is_array()) throw std::invalid_argument("Expected layers array");
  std::vector<Layer> result;
  for (const auto& item : value) {
    Keys(item, {"id", "sourceLayerId", "name", "kind", "visible", "opacity", "fillOpacity",
                "blendMode", "compositingKnown", "clippingBase", "features", "bounds", "text",
                "rectangle", "children"});
    Layer layer;
    layer.id = item.at("id").get<std::string>();
    layer.sourceLayerId = Integer(item.at("sourceLayerId"));
    layer.name = item.at("name").get<std::string>();
    layer.kind = Kind(item.at("kind").get<std::string>());
    layer.visible = item.at("visible").get<bool>();
    layer.opacity = item.at("opacity").get<float>();
    layer.fillOpacity = item.at("fillOpacity").get<float>();
    layer.blendMode = item.at("blendMode").get<std::string>();
    layer.compositingKnown = item.at("compositingKnown").get<bool>();
    layer.clippingBase = item.at("clippingBase").get<std::string>();
    layer.features = item.at("features").get<std::vector<std::string>>();
    layer.bounds = Bounds(item.at("bounds"));
    layer.text = item.at("text").get<std::string>();
    if (item.contains("rectangle")) {
      const auto& v = item.at("rectangle");
      Keys(v, {"bounds", "red", "green", "blue", "alpha", "roundness"});
      SolidRectangle rect;
      rect.bounds = Bounds(v.at("bounds"));
      rect.red = v.at("red").get<float>();
      rect.green = v.at("green").get<float>();
      rect.blue = v.at("blue").get<float>();
      rect.alpha = v.at("alpha").get<float>();
      rect.roundness = v.at("roundness").get<float>();
      layer.rectangle = rect;
    }
    layer.children = Layers(item.at("children"));
    result.push_back(std::move(layer));
  }
  return result;
}
}  // namespace

Document ParseSnapshot(const std::string& text) {
  if (text.size() > 32 * 1024 * 1024) throw std::invalid_argument("Snapshot exceeds 32 MiB");
  try {
    auto root = JSON::parse(text, [](int depth, JSON::parse_event_t, JSON&) {
      if (depth > 260) throw std::invalid_argument("Snapshot nesting exceeds limit");
      return true;
    });
    Keys(root, {"schemaVersion", "name", "hostVersion", "resolution", "colorMode", "colorProfile",
                "bitDepth", "bounds", "features", "layers"});
    Document doc;
    doc.schemaVersion = Integer(root.at("schemaVersion"));
    doc.name = root.at("name").get<std::string>();
    doc.hostVersion = root.at("hostVersion").get<std::string>();
    doc.resolution = root.at("resolution").get<float>();
    doc.colorMode = root.at("colorMode").get<std::string>();
    doc.colorProfile = root.at("colorProfile").get<std::string>();
    doc.bitDepth = Integer(root.at("bitDepth"));
    doc.bounds = Bounds(root.at("bounds"));
    doc.features = root.at("features").get<std::vector<std::string>>();
    doc.layers = Layers(root.at("layers"));
    Validate(doc);
    return doc;
  } catch (const JSON::exception& error) {
    throw std::invalid_argument(std::string("Invalid Photoshop snapshot JSON: ") + error.what());
  }
}
}  // namespace photoshop
