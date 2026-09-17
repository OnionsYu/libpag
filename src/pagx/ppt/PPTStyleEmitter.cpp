/////////////////////////////////////////////////////////////////////////////////////////////////
//
//  Tencent is pleased to support the open source community by making libpag available.
//
//  Copyright (C) 2026 Tencent. All rights reserved.
//
//  Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file
//  except in compliance with the License. You may obtain a copy of the License at
//
//      http://www.apache.org/licenses/LICENSE-2.0
//
//  unless required by applicable law or agreed to in writing, software distributed under the
//  license is distributed on an "as is" basis, without warranties or conditions of any kind,
//  either express or implied. see the license for the specific language governing permissions
//  and limitations under the license.
//
/////////////////////////////////////////////////////////////////////////////////////////////////

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <limits>
#include <string>
#include <vector>
#include "pagx/nodes/ConicGradient.h"
#include "pagx/nodes/DiamondGradient.h"
#include "pagx/nodes/Image.h"
#include "pagx/nodes/LinearGradient.h"
#include "pagx/nodes/RadialGradient.h"
#include "pagx/nodes/SolidColor.h"
#include "pagx/ppt/PPTWriter.h"
#include "pagx/utils/ColorSpaceUtils.h"

namespace pagx {

using pag::FloatNearlyZero;
using pag::RadiansToDegrees;

namespace {

struct GradientStopValue {
  float offset = 0;
  Color color = {};

  bool operator<(const GradientStopValue& other) const {
    return offset < other.offset;
  }
};

Color SampleGradientColor(const std::vector<GradientStopValue>& stops, float position) {
  if (position <= stops.front().offset) {
    return stops.front().color;
  }
  if (position >= stops.back().offset) {
    return stops.back().color;
  }
  for (size_t i = 1; i < stops.size(); ++i) {
    if (position <= stops[i].offset) {
      const auto& previous = stops[i - 1];
      const auto& current = stops[i];
      float span = current.offset - previous.offset;
      if (FloatNearlyZero(span)) {
        return current.color;
      }
      return LerpColor(previous.color, current.color, (position - previous.offset) / span);
    }
  }
  return stops.back().color;
}

}  // namespace

void PPTWriter::WriteBlip(XMLBuilder& out, const std::string& relId, float alpha) {
  out.openElement("a:blip").addRequiredAttribute("r:embed", relId);
  if (alpha < 1.0f) {
    out.closeElementStart();
    out.openElement("a:alphaModFix")
        .addRequiredAttribute("amt", AlphaToPct(alpha))
        .closeElementSelfClosing();
    out.closeElement();  // a:blip
  } else {
    out.closeElementSelfClosing();
  }
}

void PPTWriter::WriteDefaultStretch(XMLBuilder& out) {
  out.openElement("a:stretch").closeElementStart();
  out.openElement("a:fillRect").closeElementSelfClosing();
  out.closeElement();  // a:stretch
}

// ── Picture helpers ─────────────────────────────────────────────────────────

void PPTWriter::beginPicture(XMLBuilder& out, const char* name) {
  int id = _ctx->nextShapeId();
  out.openElement("p:pic").closeElementStart();

  out.openElement("p:nvPicPr").closeElementStart();
  out.openElement("p:cNvPr")
      .addRequiredAttribute("id", id)
      .addRequiredAttribute("name", name)
      .closeElementSelfClosing();
  out.openElement("p:cNvPicPr").closeElementStart();
  out.openElement("a:picLocks")
      .addRequiredAttribute("noChangeAspect", "1")
      .closeElementSelfClosing();
  out.closeElement();  // p:cNvPicPr
  out.openElement("p:nvPr").closeElementSelfClosing();
  out.closeElement();  // p:nvPicPr
}

void PPTWriter::endPicture(XMLBuilder& out, const Xform& xf) {
  out.openElement("p:spPr").closeElementStart();
  WriteXfrm(out, xf);
  out.openElement("a:prstGeom").addRequiredAttribute("prst", "rect").closeElementStart();
  out.openElement("a:avLst").closeElementSelfClosing();
  out.closeElement();  // a:prstGeom
  out.closeElement();  // p:spPr

  out.closeElement();  // p:pic
}

// ── Rasterized picture ──────────────────────────────────────────────────────

void PPTWriter::writePicture(XMLBuilder& out, const std::string& relId, int64_t offX, int64_t offY,
                             int64_t extCX, int64_t extCY) {
  beginPicture(out, "MaskedImage");

  out.openElement("p:blipFill").closeElementStart();
  WriteBlip(out, relId, 1.0f);
  WriteDefaultStretch(out);
  out.closeElement();  // p:blipFill

  endPicture(out, {offX, offY, extCX, extCY, 0});
}

// ── ImagePattern as p:pic element ──────────────────────────────────────────

bool PPTWriter::writeImagePatternAsPicture(XMLBuilder& out, const Fill* fill,
                                           const Rect& shapeBounds, const Matrix& m, float alpha,
                                           bool rectangularClip) {
  if (!fill || !fill->color || fill->color->nodeType() != NodeType::ImagePattern) {
    return false;
  }
  auto* pattern = static_cast<const ImagePattern*>(fill->color);
  if (!pattern->image) {
    return false;
  }
  // Repeat/Mirror need OOXML <a:tile> (or a baked tiled PNG); Clamp needs the
  // image's edge pixels to extend across the whole shape, which can't be
  // expressed by drawing a single <p:pic> at the natural extent.  Only Decal
  // (transparent outside the natural extent) maps cleanly onto a <p:pic>.
  if (pattern->tileModeX != TileMode::Decal || pattern->tileModeY != TileMode::Decal) {
    return false;
  }

  int imgW = 0;
  int imgH = 0;
  if (!GetImageDimensions(pattern->image, &imgW, &imgH) || shapeBounds.isEmpty()) {
    return false;
  }

  ImagePatternRect ipr = {};
  if (!ComputeImagePatternRect(pattern, imgW, imgH, shapeBounds, &ipr)) {
    return false;
  }

  // A standalone p:pic always has rectangular geometry. It is therefore safe only for a sharp
  // rectangular host. Rounded rectangles, ellipses, and custom paths must keep the image as the
  // host shape's a:blipFill so the shape geometry clips the image correctly.
  if (!rectangularClip) {
    return false;
  }

  // If the image already covers the whole shape, the host shape's a:blipFill is exact and avoids
  // an unnecessary extra picture. p:pic is only needed for a strict inner image rectangle, where
  // it preserves the transparent area surrounding a Decal image.
  bool fillsShape = ipr.visL <= shapeBounds.x + 0.5f && ipr.visT <= shapeBounds.y + 0.5f &&
                    ipr.visR >= shapeBounds.x + shapeBounds.width - 0.5f &&
                    ipr.visB >= shapeBounds.y + shapeBounds.height - 0.5f;
  if (fillsShape) {
    return false;
  }

  std::string relId = _ctx->addImage(pattern->image);
  if (relId.empty()) {
    return false;
  }

  float effectiveAlpha = fill->alpha * alpha;

  bool hasSrcRect = (ipr.srcL != 0 || ipr.srcT != 0 || ipr.srcR != 0 || ipr.srcB != 0);

  auto xf = DecomposeXform(ipr.visL, ipr.visT, ipr.visR - ipr.visL, ipr.visB - ipr.visT, m);

  beginPicture(out, "Image");

  out.openElement("p:blipFill").closeElementStart();
  WriteBlip(out, relId, effectiveAlpha);
  if (hasSrcRect) {
    out.openElement("a:srcRect");
    AddLTRBAttrs(out, {ipr.srcL, ipr.srcT, ipr.srcR, ipr.srcB});
    out.closeElementSelfClosing();
  }
  WriteDefaultStretch(out);
  out.closeElement();  // p:blipFill

  endPicture(out, xf);
  return true;
}

// ── Color source / gradient ────────────────────────────────────────────────

void PPTWriter::writeGradientStops(XMLBuilder& out, const std::vector<ColorStop*>& stops,
                                   float alpha) {
  out.openElement("a:gsLst").closeElementStart();
  for (const auto* stop : stops) {
    int pos = std::clamp(static_cast<int>(std::round(stop->offset * 100000.0f)), 0, 100000);
    out.openElement("a:gs").addRequiredAttribute("pos", pos).closeElementStart();
    WriteSrgbClr(out, stop->color, stop->color.alpha * alpha);
    out.closeElement();  // a:gs
  }
  out.closeElement();  // a:gsLst
}

void PPTWriter::writeLinearGradientStops(XMLBuilder& out, const LinearGradient* gradient,
                                         float alpha, const Rect& shapeBounds) {
  Point start = gradient->matrix.mapPoint(gradient->startPoint);
  Point end = gradient->matrix.mapPoint(gradient->endPoint);
  if (gradient->fitsToGeometry) {
    start = {shapeBounds.x + start.x * shapeBounds.width,
             shapeBounds.y + start.y * shapeBounds.height};
    end = {shapeBounds.x + end.x * shapeBounds.width, shapeBounds.y + end.y * shapeBounds.height};
  }

  float dx = end.x - start.x;
  float dy = end.y - start.y;
  float length2 = dx * dx + dy * dy;
  if (shapeBounds.isEmpty() || FloatNearlyZero(length2)) {
    writeGradientStops(out, gradient->colorStops, alpha);
    return;
  }

  // DrawingML linear gradients span the projection range of the whole shape. PAGX instead maps
  // stop 0/1 to the authored start/end points, which may cover only part of the shape. Project the
  // shape corners onto the authored gradient line and remap the stop positions into DrawingML's
  // full-shape coordinate range. Stops outside that range are clamped; duplicate boundary stops
  // preserve the endpoint colours over any area before/after the authored gradient segment.
  Point corners[] = {{shapeBounds.x, shapeBounds.y},
                     {shapeBounds.x + shapeBounds.width, shapeBounds.y},
                     {shapeBounds.x + shapeBounds.width, shapeBounds.y + shapeBounds.height},
                     {shapeBounds.x, shapeBounds.y + shapeBounds.height}};
  float minT = std::numeric_limits<float>::max();
  float maxT = std::numeric_limits<float>::lowest();
  for (const auto& corner : corners) {
    float t = ((corner.x - start.x) * dx + (corner.y - start.y) * dy) / length2;
    minT = std::min(minT, t);
    maxT = std::max(maxT, t);
  }
  float range = maxT - minT;
  if (FloatNearlyZero(range)) {
    writeGradientStops(out, gradient->colorStops, alpha);
    return;
  }

  std::vector<GradientStopValue> sourceStops;
  sourceStops.reserve(gradient->colorStops.size());
  for (const auto* stop : gradient->colorStops) {
    if (stop != nullptr) {
      sourceStops.push_back({stop->offset, stop->color});
    }
  }
  if (sourceStops.empty()) {
    writeGradientStops(out, gradient->colorStops, alpha);
    return;
  }
  std::stable_sort(sourceStops.begin(), sourceStops.end());

  struct MappedStop {
    float offset;
    Color color;
  };
  std::vector<MappedStop> mapped;
  mapped.reserve(sourceStops.size() + 2);
  mapped.push_back({0.0f, SampleGradientColor(sourceStops, minT)});
  for (const auto& stop : sourceStops) {
    if (stop.offset > minT && stop.offset < maxT) {
      mapped.push_back({(stop.offset - minT) / range, stop.color});
    }
  }
  mapped.push_back({1.0f, SampleGradientColor(sourceStops, maxT)});

  out.openElement("a:gsLst").closeElementStart();
  for (const auto& entry : mapped) {
    int pos = std::clamp(static_cast<int>(std::round(entry.offset * 100000.0f)), 0, 100000);
    out.openElement("a:gs").addRequiredAttribute("pos", pos).closeElementStart();
    WriteSrgbClr(out, entry.color, entry.color.alpha * alpha);
    out.closeElement();  // a:gs
  }
  out.closeElement();  // a:gsLst
}

void PPTWriter::writeColorSource(XMLBuilder& out, const ColorSource* source, float alpha,
                                 const Rect& shapeBounds) {
  if (!source) {
    out.openElement("a:noFill").closeElementSelfClosing();
    return;
  }

  switch (source->nodeType()) {
    case NodeType::SolidColor: {
      auto* solid = static_cast<const SolidColor*>(source);
      writeSolidColorFill(out, solid->color, alpha);
      break;
    }
    case NodeType::LinearGradient: {
      auto* grad = static_cast<const LinearGradient*>(source);
      Point sp = grad->matrix.mapPoint(grad->startPoint);
      Point ep = grad->matrix.mapPoint(grad->endPoint);
      float dx = ep.x - sp.x;
      float dy = ep.y - sp.y;
      // When fitsToGeometry is true, start/end are in normalized (0-1) space.
      // Scale by the shape aspect ratio so atan2 produces the correct visual
      // angle in document space.
      if (grad->fitsToGeometry && shapeBounds.width > 0 && shapeBounds.height > 0) {
        dx *= shapeBounds.width;
        dy *= shapeBounds.height;
      }
      float angleDeg = RadiansToDegrees(std::atan2(dy, dx));
      int ang = AngleToPPT(angleDeg);

      out.openElement("a:gradFill").closeElementStart();
      writeLinearGradientStops(out, grad, alpha, shapeBounds);
      out.openElement("a:lin")
          .addRequiredAttribute("ang", ang)
          .addRequiredAttribute("scaled", "1")
          .closeElementSelfClosing();
      out.closeElement();  // a:gradFill
      break;
    }
    case NodeType::RadialGradient: {
      // Map the gradient center through its matrix to document space and convert to
      // fillToRect insets (1/1000 percent) relative to the shape bounds.  OOXML
      // radial gradients always span the whole bounding box, so radius cannot be
      // honoured exactly; only the focus position is preserved.  When shapeBounds
      // is empty (e.g. text fills) fall back to the shape center.
      auto* grad = static_cast<const RadialGradient*>(source);
      out.openElement("a:gradFill").closeElementStart();
      writeGradientStops(out, grad->colorStops, alpha);
      out.openElement("a:path").addRequiredAttribute("path", "circle").closeElementStart();
      WriteFillToRectFromCenter(out, grad, shapeBounds);
      out.closeElement();  // a:path
      out.closeElement();  // a:gradFill
      break;
    }
    case NodeType::ConicGradient: {
      // OOXML has no conic/sweep gradient primitive. Layers containing one are
      // normally escalated to a layer-level rasterization fallback by
      // PPTFeatureProbe; this branch is only reached when the rasterizer is
      // unavailable (e.g. no GPU) and emits a circular path gradient as a
      // last-resort fallback so something visible still ends up on the slide.
      auto* grad = static_cast<const ConicGradient*>(source);
      out.openElement("a:gradFill").addRequiredAttribute("rotWithShape", "1").closeElementStart();
      writeGradientStops(out, grad->colorStops, alpha);
      out.openElement("a:path").addRequiredAttribute("path", "circle").closeElementStart();
      WriteFillToRectFromCenter(out, grad, shapeBounds);
      out.closeElement();  // a:path
      out.closeElement();  // a:gradFill
      break;
    }
    case NodeType::DiamondGradient: {
      // OOXML has no diamond gradient primitive. The closest preset is a
      // rectangular path gradient, which produces an axis-aligned diamond-like
      // pattern when the focus rect is collapsed to the center point.
      auto* grad = static_cast<const DiamondGradient*>(source);
      out.openElement("a:gradFill").addRequiredAttribute("rotWithShape", "1").closeElementStart();
      writeGradientStops(out, grad->colorStops, alpha);
      out.openElement("a:path").addRequiredAttribute("path", "rect").closeElementStart();
      WriteFillToRectFromCenter(out, grad, shapeBounds);
      out.closeElement();  // a:path
      out.closeElement();  // a:gradFill
      break;
    }
    case NodeType::ImagePattern:
      writeImagePatternFill(out, static_cast<const ImagePattern*>(source), alpha, shapeBounds);
      break;
    default:
      out.openElement("a:noFill").closeElementSelfClosing();
      break;
  }
}

void PPTWriter::writeImagePatternFill(XMLBuilder& out, const ImagePattern* pattern, float alpha,
                                      const Rect& shapeBounds) {
  if (!pattern->image) {
    out.openElement("a:noFill").closeElementSelfClosing();
    return;
  }

  bool needsNativeTile =
      (pattern->tileModeX == TileMode::Repeat || pattern->tileModeX == TileMode::Mirror ||
       pattern->tileModeY == TileMode::Repeat || pattern->tileModeY == TileMode::Mirror);
  // Clamp also needs the bake path: OOXML can't express "extend edge pixels"
  // natively, and emitting just <a:srcRect>/<a:fillRect> would either tile
  // (wrong) or stretch the whole image (also wrong).  Repeat/Mirror prefer the
  // bake too because it preserves the pattern's shader matrix exactly.
  bool needsBake = needsNativeTile || pattern->tileModeX == TileMode::Clamp ||
                   pattern->tileModeY == TileMode::Clamp;

  if (needsBake && !shapeBounds.isEmpty()) {
    int w = static_cast<int>(ceilf(shapeBounds.width));
    int h = static_cast<int>(ceilf(shapeBounds.height));
    float offsetX = pattern->matrix.tx - shapeBounds.x;
    float offsetY = pattern->matrix.ty - shapeBounds.y;
    auto tiledPng = RenderTiledPattern(&_gpu, pattern, w, h, offsetX, offsetY, _rasterScale);
    if (tiledPng) {
      auto relId = _ctx->addRawImage(std::move(tiledPng));
      out.openElement("a:blipFill").closeElementStart();
      WriteBlip(out, relId, alpha);
      WriteDefaultStretch(out);
      out.closeElement();  // a:blipFill
      return;
    }
  }

  std::string relId = _ctx->addImage(pattern->image);
  if (relId.empty()) {
    out.openElement("a:noFill").closeElementSelfClosing();
    return;
  }

  int imgW = 0;
  int imgH = 0;
  float imgDpiX = 72.0f;
  float imgDpiY = 72.0f;
  // Read the image bytes once and feed them to both the dimensions and DPI
  // probes so pattern fills don't re-open the source file. Falls back to the
  // header-only PNG path (GetImagePNGDimensions) when no byte stream is
  // available — that path still honors the 24-byte fast read from file.
  bool hasDimensions = false;
  auto imageBytes = GetImageData(pattern->image);
  if (imageBytes && imageBytes->size() > 0) {
    const uint8_t* bytes = imageBytes->bytes();
    size_t byteCount = imageBytes->size();
    hasDimensions = GetPNGDimensions(bytes, byteCount, &imgW, &imgH) ||
                    GetJPEGDimensions(bytes, byteCount, &imgW, &imgH) ||
                    GetWebPDimensions(bytes, byteCount, &imgW, &imgH);
    if (!GetPNGDPI(bytes, byteCount, &imgDpiX, &imgDpiY)) {
      GetJPEGDPI(bytes, byteCount, &imgDpiX, &imgDpiY);
    }
  } else {
    hasDimensions = GetImagePNGDimensions(pattern->image, &imgW, &imgH);
  }

  out.openElement("a:blipFill").closeElementStart();
  WriteBlip(out, relId, alpha);

  if (needsNativeTile && hasDimensions && !shapeBounds.isEmpty()) {
    const auto& M = pattern->matrix;
    double dpiCorrX = static_cast<double>(imgDpiX) / 96.0;
    double dpiCorrY = static_cast<double>(imgDpiY) / 96.0;
    auto sx = static_cast<int>(std::round(std::sqrt(M.a * M.a + M.b * M.b) * dpiCorrX * 100000.0));
    auto sy = static_cast<int>(std::round(std::sqrt(M.c * M.c + M.d * M.d) * dpiCorrY * 100000.0));
    auto tx = PxToEMU(M.tx - shapeBounds.x);
    auto ty = PxToEMU(M.ty - shapeBounds.y);
    bool flipX = (pattern->tileModeX == TileMode::Mirror);
    bool flipY = (pattern->tileModeY == TileMode::Mirror);
    const char* flip = "none";
    if (flipX && flipY) {
      flip = "xy";
    } else if (flipX) {
      flip = "x";
    } else if (flipY) {
      flip = "y";
    }
    out.openElement("a:tile")
        .addRequiredAttribute("tx", tx)
        .addRequiredAttribute("ty", ty)
        .addRequiredAttribute("sx", sx)
        .addRequiredAttribute("sy", sy)
        .addRequiredAttribute("flip", flip)
        .addRequiredAttribute("algn", "tl")
        .closeElementSelfClosing();
  } else {
    bool hasPlacement = hasDimensions && !shapeBounds.isEmpty();
    ImagePatternRect ipr = {};
    LTRBInsets srcInsets;
    LTRBInsets fillInsets;
    bool placed = false;
    if (hasPlacement && ComputeImagePatternRect(pattern, imgW, imgH, shapeBounds, &ipr)) {
      srcInsets.l = ipr.srcL;
      srcInsets.t = ipr.srcT;
      srcInsets.r = ipr.srcR;
      srcInsets.b = ipr.srcB;
      fillInsets.l =
          static_cast<int>(std::round((ipr.visL - shapeBounds.x) / shapeBounds.width * 100000.0f));
      fillInsets.t =
          static_cast<int>(std::round((ipr.visT - shapeBounds.y) / shapeBounds.height * 100000.0f));
      fillInsets.r = static_cast<int>(std::round((shapeBounds.x + shapeBounds.width - ipr.visR) /
                                                 shapeBounds.width * 100000.0f));
      fillInsets.b = static_cast<int>(std::round((shapeBounds.y + shapeBounds.height - ipr.visB) /
                                                 shapeBounds.height * 100000.0f));
      placed = srcInsets.any() || fillInsets.any();
    }
    if (placed) {
      out.openElement("a:srcRect");
      AddLTRBAttrs(out, srcInsets);
      out.closeElementSelfClosing();
      out.openElement("a:stretch").closeElementStart();
      out.openElement("a:fillRect");
      AddLTRBAttrs(out, fillInsets);
      out.closeElementSelfClosing();
      out.closeElement();  // a:stretch
    } else {
      WriteDefaultStretch(out);
    }
  }

  out.closeElement();  // a:blipFill
}

// ── Fill / stroke ──────────────────────────────────────────────────────────

void PPTWriter::writeSolidColorFill(XMLBuilder& out, const Color& color, float alpha) {
  out.openElement("a:solidFill").closeElementStart();
  WriteSrgbClr(out, color, color.alpha * alpha);
  out.closeElement();  // a:solidFill
}

void PPTWriter::writeFill(XMLBuilder& out, const Fill* fill, float alpha, const Rect& shapeBounds) {
  if (!fill) {
    out.openElement("a:noFill").closeElementSelfClosing();
    return;
  }
  float effectiveAlpha = fill->alpha * alpha;
  if (!fill->color) {
    // Per PAGX spec, Fill defaults to opaque black (#000000) when no color is specified.
    writeSolidColorFill(out, Color{}, effectiveAlpha);
    return;
  }
  writeColorSource(out, fill->color, effectiveAlpha, shapeBounds);
}

void PPTWriter::writeStroke(XMLBuilder& out, const Stroke* stroke, float alpha) {
  if (!stroke) {
    out.openElement("a:ln").closeElementStart();
    out.openElement("a:noFill").closeElementSelfClosing();
    out.closeElement();
    return;
  }

  int64_t w = PxToEMU(stroke->width);
  out.openElement("a:ln").addRequiredAttribute("w", w);
  if (stroke->cap == LineCap::Round) {
    out.addRequiredAttribute("cap", "rnd");
  } else if (stroke->cap == LineCap::Square) {
    out.addRequiredAttribute("cap", "sq");
  }
  out.closeElementStart();

  float effectiveAlpha = stroke->alpha * alpha;
  if (stroke->color) {
    writeColorSource(out, stroke->color, effectiveAlpha);
  } else {
    // Per PAGX spec, Stroke defaults to opaque black (#000000) when no color is specified.
    writeSolidColorFill(out, Color{}, effectiveAlpha);
  }

  if (!stroke->dashes.empty()) {
    float sw = (stroke->width > 0) ? stroke->width : 1.0f;
    const char* preset = MatchPresetDash(stroke->dashes, sw);
    if (preset) {
      out.openElement("a:prstDash").addRequiredAttribute("val", preset).closeElementSelfClosing();
    } else {
      out.openElement("a:custDash").closeElementStart();
      for (size_t i = 0; i + 1 < stroke->dashes.size(); i += 2) {
        int d = std::max(1, static_cast<int>(std::round(stroke->dashes[i] / sw * 100000.0)));
        int sp = std::max(1, static_cast<int>(std::round(stroke->dashes[i + 1] / sw * 100000.0)));
        out.openElement("a:ds")
            .addRequiredAttribute("d", d)
            .addRequiredAttribute("sp", sp)
            .closeElementSelfClosing();
      }
      if (stroke->dashes.size() % 2 != 0) {
        int d = std::max(1, static_cast<int>(std::round(stroke->dashes.back() / sw * 100000.0)));
        out.openElement("a:ds")
            .addRequiredAttribute("d", d)
            .addRequiredAttribute("sp", d)
            .closeElementSelfClosing();
      }
      out.closeElement();  // a:custDash
    }
  }

  if (stroke->join == LineJoin::Round) {
    out.openElement("a:round").closeElementSelfClosing();
  } else if (stroke->join == LineJoin::Bevel) {
    out.openElement("a:bevel").closeElementSelfClosing();
  } else {
    int lim = static_cast<int>(std::round(stroke->miterLimit * 100000.0f));
    out.openElement("a:miter").addRequiredAttribute("lim", lim).closeElementSelfClosing();
  }

  out.closeElement();  // a:ln
}

// ── Effects (shadow) ───────────────────────────────────────────────────────

void PPTWriter::writeShadowElement(XMLBuilder& out, const char* tag, float blurX, float blurY,
                                   float offsetX, float offsetY, const Color& color,
                                   bool includeAlign) {
  float blur = (blurX + blurY) / 2.0f;
  float dist = std::sqrt(offsetX * offsetX + offsetY * offsetY);
  // OOXML `dir` measures clockwise from the +x axis in screen space (Y-down),
  // matching atan2(offsetY, offsetX) directly — no axis swap needed.
  float dir = RadiansToDegrees(std::atan2(offsetY, offsetX));
  auto& builder = out.openElement(tag)
                      .addRequiredAttribute("blurRad", PxToEMU(blur))
                      .addRequiredAttribute("dist", PxToEMU(dist))
                      .addRequiredAttribute("dir", AngleToPPT(dir));
  if (includeAlign) {
    builder.addRequiredAttribute("algn", "ctr").addRequiredAttribute("rotWithShape", "0");
  }
  builder.closeElementStart();
  out.openElement("a:srgbClr").addRequiredAttribute("val", ColorToHex6(color)).closeElementStart();
  out.openElement("a:alpha")
      .addRequiredAttribute("val", AlphaToPct(color.alpha))
      .closeElementSelfClosing();
  out.closeElement();  // a:srgbClr
  out.closeElement();  // tag
}

void PPTWriter::writeEffects(XMLBuilder& out, const std::vector<LayerFilter*>& filters,
                             const std::vector<LayerStyle*>& styles) {
  auto sources = CollectEffectSources(filters, styles);
  if (sources.empty()) {
    return;
  }

  out.openElement("a:effectLst").closeElementStart();

  // OOXML effectLst requires this element order (§20.1.8.20):
  // blur, fillOverlay, glow, innerShdw, outerShdw, prstShdw, reflection, softEdge

  // BlurFilter blurs the shape itself, which is exactly what OOXML's a:blur does, so
  // the mapping is direct. grow="1" allows the blur to extend beyond the original
  // shape bounds — required for solid-filled shapes, otherwise PowerPoint clips the
  // blurred edges back to the original bounds and the effect becomes invisible.
  //
  // BackgroundBlurStyle and GlassStyle are intentionally NOT mapped to a:blur. They
  // are backdrop-sampling effects that operate on the pixels behind the shape while
  // leaving the shape's own fill and stroke crisp. OOXML has no native backdrop
  // primitive and using a:blur as a stand-in produces visibly wrong output: the
  // card's stroke, rounded corners, and fill all become fuzzy. Authors who need a
  // faithful render of either style should enable `bakeUnsupported` so the layer is
  // baked with its real backdrop via rasterizeLayerAsPicture (probed in
  // PPTFeatureProbe).
  if (sources.blur) {
    float avgBlur = (sources.blur->blurX + sources.blur->blurY) / 2.0f;
    if (avgBlur > 0) {
      out.openElement("a:blur")
          .addRequiredAttribute("rad", PxToEMU(avgBlur))
          .addRequiredAttribute("grow", "true")
          .closeElementSelfClosing();
    }
  }

  if (sources.blend) {
    const char* blendStr = BlendModeToPPT(sources.blend->blendMode);
    if (blendStr) {
      out.openElement("a:fillOverlay").addRequiredAttribute("blend", blendStr).closeElementStart();
      out.openElement("a:solidFill").closeElementStart();
      WriteSrgbClr(out, sources.blend->color, sources.blend->color.alpha);
      out.closeElement();  // a:solidFill
      out.closeElement();  // a:fillOverlay
    }
  }

  // OOXML §20.1.8.20 allows at most one <a:innerShdw> and one <a:outerShdw>
  // per <a:effectLst>. Filters take precedence over styles.
  if (sources.innerShadowFilter) {
    auto* s = sources.innerShadowFilter;
    writeShadowElement(out, "a:innerShdw", s->blurX, s->blurY, s->offsetX, s->offsetY, s->color,
                       false);
  } else if (sources.innerShadowStyle) {
    auto* s = sources.innerShadowStyle;
    writeShadowElement(out, "a:innerShdw", s->blurX, s->blurY, s->offsetX, s->offsetY, s->color,
                       false);
  }
  if (sources.dropShadowFilter) {
    auto* s = sources.dropShadowFilter;
    writeShadowElement(out, "a:outerShdw", s->blurX, s->blurY, s->offsetX, s->offsetY, s->color,
                       true);
  } else if (sources.dropShadowStyle) {
    auto* s = sources.dropShadowStyle;
    writeShadowElement(out, "a:outerShdw", s->blurX, s->blurY, s->offsetX, s->offsetY, s->color,
                       true);
  }

  out.closeElement();  // a:effectLst
}

}  // namespace pagx
