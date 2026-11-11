/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include <cmath>
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
float Number(const JSON& value) {
  if (!value.is_number()) throw std::invalid_argument("Expected a snapshot number");
  auto number = value.get<float>();
  if (!std::isfinite(number)) throw std::invalid_argument("Non-finite snapshot number");
  return number;
}
Rect Bounds(const JSON& value) {
  Keys(value, {"left", "top", "width", "height"});
  return {Number(value.at("left")), Number(value.at("top")), Number(value.at("width")),
          Number(value.at("height"))};
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
Point ReadPoint(const JSON& value) {
  Keys(value, {"x", "y"});
  return {Number(value.at("x")), Number(value.at("y"))};
}
Color ReadColor(const JSON& value) {
  Keys(value, {"red", "green", "blue", "alpha"});
  return {Number(value.at("red")), Number(value.at("green")), Number(value.at("blue")),
          Number(value.at("alpha"))};
}
Shape ReadShape(const JSON& value) {
  Keys(value, {"components", "fill", "stroke"});
  Shape shape;
  const auto& components = value.at("components");
  if (!components.is_array()) throw std::invalid_argument("Expected path components array");
  for (const auto& c : components) {
    Keys(c, {"operation", "winding", "subpaths"});
    PathComponent component;
    const auto operation = c.at("operation").get<std::string>();
    if (operation == "add") component.operation = PathOperation::Add;
    else if (operation == "subtract")
      component.operation = PathOperation::Subtract;
    else if (operation == "intersect")
      component.operation = PathOperation::Intersect;
    else if (operation == "exclude")
      component.operation = PathOperation::Exclude;
    else
      throw std::invalid_argument("Invalid path operation");
    component.winding = c.at("winding").get<bool>();
    if (!c.at("subpaths").is_array()) throw std::invalid_argument("Expected subpaths array");
    for (const auto& p : c.at("subpaths")) {
      Keys(p, {"closed", "points"});
      Subpath path;
      path.closed = p.at("closed").get<bool>();
      if (!p.at("points").is_array()) throw std::invalid_argument("Expected points array");
      for (const auto& v : p.at("points")) {
        Keys(v, {"anchor", "incoming", "outgoing"});
        path.points.push_back(
            {ReadPoint(v.at("anchor")), ReadPoint(v.at("incoming")), ReadPoint(v.at("outgoing"))});
      }
      component.subpaths.push_back(std::move(path));
    }
    shape.components.push_back(std::move(component));
  }
  if (value.contains("fill")) shape.fill = ReadColor(value.at("fill"));
  if (value.contains("stroke")) {
    const auto& v = value.at("stroke");
    Keys(v, {"color", "width", "opacity", "miterLimit", "cap", "join", "align", "dashes",
             "dashOffset"});
    Stroke stroke;
    stroke.color = ReadColor(v.at("color"));
    stroke.width = Number(v.at("width"));
    stroke.opacity = Number(v.at("opacity"));
    stroke.miterLimit = Number(v.at("miterLimit"));
    stroke.cap = v.at("cap").get<std::string>();
    stroke.join = v.at("join").get<std::string>();
    stroke.align = v.at("align").get<std::string>();
    stroke.dashOffset = Number(v.at("dashOffset"));
    if (!v.at("dashes").is_array()) throw std::invalid_argument("Expected dashes array");
    for (const auto& dash : v.at("dashes")) stroke.dashes.push_back(Number(dash));
    shape.stroke = std::move(stroke);
  }
  return shape;
}
std::vector<Layer> Layers(const JSON& value) {
  if (!value.is_array()) throw std::invalid_argument("Expected layers array");
  std::vector<Layer> result;
  for (const auto& item : value) {
    Keys(item, {"id", "sourceLayerId", "name", "kind", "visible", "opacity", "fillOpacity",
                "blendMode", "compositingKnown", "clippingBase", "features", "bounds", "text",
                "rectangle", "children", "shape", "rasterFeatures"});
    Layer layer;
    layer.id = item.at("id").get<std::string>();
    layer.sourceLayerId = Integer(item.at("sourceLayerId"));
    layer.name = item.at("name").get<std::string>();
    layer.kind = Kind(item.at("kind").get<std::string>());
    layer.visible = item.at("visible").get<bool>();
    layer.opacity = Number(item.at("opacity"));
    layer.fillOpacity = Number(item.at("fillOpacity"));
    layer.blendMode = item.at("blendMode").get<std::string>();
    layer.compositingKnown = item.at("compositingKnown").get<bool>();
    layer.clippingBase = item.at("clippingBase").get<std::string>();
    layer.features = item.at("features").get<std::vector<std::string>>();
    if (item.contains("rasterFeatures"))
      layer.rasterFeatures = item.at("rasterFeatures").get<std::vector<std::string>>();
    layer.bounds = Bounds(item.at("bounds"));
    layer.text = item.at("text").get<std::string>();
    if (item.contains("rectangle")) {
      const auto& v = item.at("rectangle");
      Keys(v, {"bounds", "red", "green", "blue", "alpha", "roundness"});
      SolidRectangle rect;
      rect.bounds = Bounds(v.at("bounds"));
      rect.red = Number(v.at("red"));
      rect.green = Number(v.at("green"));
      rect.blue = Number(v.at("blue"));
      rect.alpha = Number(v.at("alpha"));
      rect.roundness = Number(v.at("roundness"));
      layer.rectangle = rect;
    }
    if (item.contains("shape")) layer.shape = ReadShape(item.at("shape"));
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
    doc.resolution = Number(root.at("resolution"));
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
