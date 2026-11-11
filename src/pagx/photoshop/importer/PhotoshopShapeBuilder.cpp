/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include "PhotoshopShapeBuilder.h"
#include <stdexcept>
#include "pagx/nodes/Fill.h"
#include "pagx/nodes/Path.h"
#include "pagx/nodes/SolidColor.h"
#include "pagx/nodes/Stroke.h"
#include "tgfx/core/Path.h"

namespace pagx {
namespace {
tgfx::Path MakeComponent(const photoshop::PathComponent& component) {
  tgfx::Path path;
  path.setFillType(component.winding ? tgfx::PathFillType::Winding : tgfx::PathFillType::EvenOdd);
  for (const auto& subpath : component.subpaths) {
    const auto& points = subpath.points;
    path.moveTo(points[0].anchor.x, points[0].anchor.y);
    const auto count = points.size() + (subpath.closed ? 1 : 0);
    for (size_t i = 1; i < count; ++i) {
      const auto& previous = points[(i - 1) % points.size()];
      const auto& next = points[i % points.size()];
      if (previous.anchor.x == previous.outgoing.x && previous.anchor.y == previous.outgoing.y &&
          next.anchor.x == next.incoming.x && next.anchor.y == next.incoming.y) {
        path.lineTo(next.anchor.x, next.anchor.y);
      } else {
        path.cubicTo(previous.outgoing.x, previous.outgoing.y, next.incoming.x, next.incoming.y,
                     next.anchor.x, next.anchor.y);
      }
    }
    if (subpath.closed) path.close();
  }
  return path;
}
SolidColor* MakeColor(PAGXDocument* document, const photoshop::Color& value) {
  auto color = document->makeNode<SolidColor>();
  color->color = {value.red, value.green, value.blue, value.alpha};
  return color;
}
}  // namespace

bool CanBuildShape(const photoshop::Shape& shape) {
  // Initial subtract/intersect can depend on Photoshop's implicit canvas fill. Keep that in host.
  if (shape.components.empty() ||
      shape.components.front().operation != photoshop::PathOperation::Add)
    return false;
  // Boolean operations close open contours. Do not change open-path stroke semantics.
  if (shape.components.size() > 1) {
    for (const auto& component : shape.components) {
      for (const auto& path : component.subpaths)
        if (!path.closed) return false;
    }
  }
  return true;
}

void BuildPhotoshopShape(PAGXDocument* document, Layer* layer, const photoshop::Shape& shape,
                         float fillOpacity) {
  auto combined = MakeComponent(shape.components.front());
  for (size_t i = 1; i < shape.components.size(); ++i) {
    const auto& component = shape.components[i];
    tgfx::PathOp operation;
    switch (component.operation) {
      case photoshop::PathOperation::Add:
        operation = tgfx::PathOp::Union;
        break;
      case photoshop::PathOperation::Subtract:
        operation = tgfx::PathOp::Difference;
        break;
      case photoshop::PathOperation::Intersect:
        operation = tgfx::PathOp::Intersect;
        break;
      case photoshop::PathOperation::Exclude:
        operation = tgfx::PathOp::XOR;
        break;
      default:
        throw std::invalid_argument("Invalid path operation");
    }
    combined.addPath(MakeComponent(component), operation);
  }
  auto data = document->makeNode<PathData>();
  for (const auto& segment : combined) {
    const auto* p = segment.points;
    switch (segment.verb) {
      case tgfx::PathVerb::Move:
        data->moveTo(p[0].x, p[0].y);
        break;
      case tgfx::PathVerb::Line:
        data->lineTo(p[1].x, p[1].y);
        break;
      case tgfx::PathVerb::Quad:
        data->quadTo(p[1].x, p[1].y, p[2].x, p[2].y);
        break;
      case tgfx::PathVerb::Cubic:
        data->cubicTo(p[1].x, p[1].y, p[2].x, p[2].y, p[3].x, p[3].y);
        break;
      case tgfx::PathVerb::Close:
        data->close();
        break;
      case tgfx::PathVerb::Done:
        break;
      default:
        throw std::runtime_error("Unexpected rational curve in Photoshop path");
    }
  }
  auto path = document->makeNode<Path>();
  path->data = data;
  layer->contents.push_back(path);
  if (shape.fill) {
    auto fill = document->makeNode<Fill>();
    fill->alpha = fillOpacity;
    fill->color = MakeColor(document, *shape.fill);
    fill->fillRule = combined.getFillType() == tgfx::PathFillType::EvenOdd ? FillRule::EvenOdd
                                                                           : FillRule::Winding;
    layer->contents.push_back(fill);
  }
  if (shape.stroke) {
    const auto& value = *shape.stroke;
    auto stroke = document->makeNode<Stroke>();
    stroke->color = MakeColor(document, value.color);
    stroke->width = value.width;
    stroke->alpha = value.opacity;
    stroke->miterLimit = value.miterLimit;
    stroke->cap = value.cap == "round"    ? LineCap::Round
                  : value.cap == "square" ? LineCap::Square
                                          : LineCap::Butt;
    stroke->join = value.join == "round"   ? LineJoin::Round
                   : value.join == "bevel" ? LineJoin::Bevel
                                           : LineJoin::Miter;
    stroke->align = value.align == "inner"   ? StrokeAlign::Inside
                    : value.align == "outer" ? StrokeAlign::Outside
                                             : StrokeAlign::Center;
    stroke->placement = LayerPlacement::Foreground;
    stroke->dashes = value.dashes;
    stroke->dashOffset = value.dashOffset;
    layer->contents.push_back(stroke);
  }
}
}  // namespace pagx
