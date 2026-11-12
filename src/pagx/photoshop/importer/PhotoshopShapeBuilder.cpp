/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include "PhotoshopShapeBuilder.h"
#include <algorithm>
#include <stdexcept>
#include "PhotoshopPaintBuilder.h"
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
tgfx::Path MakeGeometry(const std::vector<photoshop::PathComponent>& components) {
  auto combined = MakeComponent(components.front());
  for (size_t i = 1; i < components.size(); ++i) {
    const auto& component = components[i];
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
  return combined;
}
Path* MakePath(PAGXDocument* document, const tgfx::Path& geometry) {
  auto data = document->makeNode<PathData>();
  for (const auto& segment : geometry) {
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
  return path;
}
Fill* MakeFill(PAGXDocument* document, const tgfx::Path& geometry, const photoshop::Paint& color,
               float opacity) {
  auto fill = document->makeNode<Fill>();
  fill->alpha = opacity;
  fill->color = BuildPhotoshopPaint(document, color);
  fill->fillRule =
      geometry.getFillType() == tgfx::PathFillType::EvenOdd ? FillRule::EvenOdd : FillRule::Winding;
  return fill;
}
}  // namespace

bool CanBuildShape(const photoshop::Shape& shape) {
  if ((shape.fill && !CanBuildPhotoshopPaint(*shape.fill)) ||
      (shape.stroke && !CanBuildPhotoshopPaint(shape.stroke->paint)))
    return false;
  // PS 27.4 can report a non-default strokeStyleOpacity while its composite still paints an
  // opaque stroke. Preserve the source value, but do not equate it with PAGX painter alpha.
  if (shape.stroke && shape.stroke->opacity != 1) return false;
  if (shape.stroke && !shape.stroke->dashes.empty() &&
      (shape.stroke->align != "center" || shape.components.size() > 1))
    return false;
  if (shape.stroke && shape.stroke->cap != "butt") {
    const auto& dashes = shape.stroke->dashes;
    // Overlapping extended caps have host-specific coverage. Odd patterns repeat with the
    // opposite parity, so every interval is a gap in one of the two repetitions.
    for (size_t i = dashes.size() % 2 ? 0 : 1; i < dashes.size(); i += dashes.size() % 2 ? 1 : 2) {
      if (dashes[i] < shape.stroke->width) return false;
    }
  }
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

bool CanBuildVectorMask(const std::optional<photoshop::VectorMask>& mask) {
  if (!mask || !mask->enabled) return true;
  if (mask->density != 1 || mask->feather != 0 || mask->initialFill) return false;
  for (const auto& component : mask->components) {
    for (const auto& path : component.subpaths)
      if (!path.closed) return false;
  }
  photoshop::Shape shape;
  shape.components = mask->components;
  return CanBuildShape(shape);
}

void BuildPhotoshopVectorMask(PAGXDocument* document, Layer* layer,
                              const photoshop::VectorMask& value, const photoshop::Rect& canvas) {
  if (!value.enabled) return;
  auto mask = document->makeNode<Layer>(layer->id + "_vector_mask");
  mask->visible = false;
  // Every source path is in document space. The owner can be a canvas-origin group or a
  // translated shape/rectangle; account for that translation exactly once.
  mask->x = -canvas.left - layer->x;
  mask->y = -canvas.top - layer->y;
  auto geometry = MakeGeometry(value.components);
  if (value.inverted) {
    tgfx::Path complement;
    complement.addRect(tgfx::Rect::MakeXYWH(canvas.left, canvas.top, canvas.width, canvas.height));
    complement.addPath(geometry, tgfx::PathOp::Difference);
    geometry = complement;
  }
  mask->contents = {MakePath(document, geometry),
                    MakeFill(document, geometry, photoshop::Color{1, 1, 1, 1}, 1)};
  layer->mask = mask;
  layer->maskType = MaskType::Contour;
  layer->children.push_back(mask);
}

void BuildPhotoshopShape(PAGXDocument* document, Layer* layer, const photoshop::Shape& shape,
                         float fillOpacity) {
  auto combined = MakeGeometry(shape.components);
  auto path = MakePath(document, combined);
  layer->contents.push_back(path);
  if (shape.fill) {
    layer->contents.push_back(MakeFill(document, combined, *shape.fill, fillOpacity));
  }
  if (shape.stroke) {
    const auto& value = *shape.stroke;
    auto stroke = document->makeNode<Stroke>();
    stroke->color = BuildPhotoshopPaint(document, value.paint);
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
    if (stroke->align != StrokeAlign::Center &&
        combined.getFillType() == tgfx::PathFillType::EvenOdd) {
      // PAGX Fill owns the fill rule; PathData and Stroke do not. Clip a doubled centered
      // stroke with an explicit vector fill so holes, overlaps and self-intersections keep
      // their original inside/outside meaning. Do not reverse contours heuristically.
      auto clip = combined;
      if (stroke->align == StrokeAlign::Outside) {
        // A finite complement only needs to contain the complete stroked footprint.
        auto bounds = combined.getBounds();
        bounds.outset(value.width * std::max(2.0f, value.miterLimit * 2) + 1,
                      value.width * std::max(2.0f, value.miterLimit * 2) + 1);
        clip = tgfx::Path();
        clip.addRect(bounds);
        clip.addPath(combined, tgfx::PathOp::Difference);
      }
      auto mask = document->makeNode<Layer>(layer->id + "_stroke_mask");
      mask->visible = false;
      mask->contents = {MakePath(document, clip),
                        MakeFill(document, clip, photoshop::Color{1, 1, 1, 1}, 1)};
      auto outline = document->makeNode<Layer>();
      outline->mask = mask;
      outline->maskType = MaskType::Contour;
      auto strokePath = document->makeNode<Path>();
      strokePath->data = path->data;
      outline->contents = {strokePath, stroke};
      stroke->width *= 2;
      stroke->align = StrokeAlign::Center;
      layer->children = {mask, outline};
    } else {
      layer->contents.push_back(stroke);
    }
  }
}
}  // namespace pagx
