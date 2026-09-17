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

#include "pagx/ppt/PPTFeatureProbe.h"
#include <cmath>
#include "base/utils/MathUtil.h"
#include "pagx/nodes/BackgroundBlurStyle.h"
#include "pagx/nodes/BlendFilter.h"
#include "pagx/nodes/BlurFilter.h"
#include "pagx/nodes/ColorMatrixFilter.h"
#include "pagx/nodes/ColorSource.h"
#include "pagx/nodes/ColorStop.h"
#include "pagx/nodes/ConicGradient.h"
#include "pagx/nodes/DiamondGradient.h"
#include "pagx/nodes/DropShadowFilter.h"
#include "pagx/nodes/DropShadowStyle.h"
#include "pagx/nodes/Fill.h"
#include "pagx/nodes/Group.h"
#include "pagx/nodes/ImagePattern.h"
#include "pagx/nodes/InnerShadowFilter.h"
#include "pagx/nodes/InnerShadowStyle.h"
#include "pagx/nodes/LayerFilter.h"
#include "pagx/nodes/LinearGradient.h"
#include "pagx/nodes/NoiseFilter.h"
#include "pagx/nodes/NoiseStyle.h"
#include "pagx/nodes/RadialGradient.h"
#include "pagx/nodes/SolidColor.h"
#include "pagx/nodes/Stroke.h"
#include "pagx/nodes/TextBox.h"
#include "pagx/types/BlendMode.h"
#include "pagx/types/Color.h"
#include "pagx/types/Matrix.h"
#include "pagx/types/StrokeStyle.h"
#include "pagx/types/TileMode.h"

namespace pagx {

using pag::FloatNearlyZero;

// A 2D affine matrix is representable by OOXML's <a:xfrm> (translation +
// rotation + axis-aligned scale/flip) iff its column basis vectors are
// orthogonal -- i.e. (a, b) . (c, d) == 0 after normalization. Any non-zero
// shear breaks that, and the matrix decomposition in PPTExporter silently
// drops the shear component, producing visibly wrong geometry.
static bool MatrixHasShear(const Matrix& m) {
  float lenA = std::sqrt(m.a * m.a + m.b * m.b);
  float lenC = std::sqrt(m.c * m.c + m.d * m.d);
  if (lenA <= 0 || lenC <= 0) {
    return false;
  }
  float dot = (m.a * m.c + m.b * m.d) / (lenA * lenC);
  return std::fabs(dot) > 1e-3f;
}

// BlendFilter emits via <a:fillOverlay>, which OOXML natively supports for a
// small set of modes. Any other mode on a BlendFilter requires rasterization.
static bool IsSupportedBlendFilterMode(BlendMode mode) {
  switch (mode) {
    case BlendMode::Normal:
    case BlendMode::Multiply:
    case BlendMode::Screen:
    case BlendMode::Darken:
    case BlendMode::Lighten:
      return true;
    default:
      return false;
  }
}

// Layer.blendMode, Fill.blendMode and Stroke.blendMode have no editable OOXML
// encoding at all -- the writer emits plain fills/strokes and silently drops
// the blend mode. Only Normal can survive the vector path; anything else must
// be rasterized so the composite against the backdrop is preserved.
static bool IsSupportedPaintBlendMode(BlendMode mode) {
  return mode == BlendMode::Normal;
}

static bool ColorIsWideGamut(const Color& c) {
  return c.colorSpace != ColorSpace::SRGB;
}

static void Merge(PPTFeatureFlags* dst, const PPTFeatureFlags& src) {
  dst->hasTextPath |= src.hasTextPath;
  dst->hasTextModifier |= src.hasTextModifier;
  dst->hasUnsupportedBlend |= src.hasUnsupportedBlend;
  dst->hasColorMatrix |= src.hasColorMatrix;
  dst->hasWideGamutColor |= src.hasWideGamutColor;
  dst->hasDiamondGradient |= src.hasDiamondGradient;
  dst->hasConicGradient |= src.hasConicGradient;
  dst->hasShearTransform |= src.hasShearTransform;
  dst->hasBackdropStyle |= src.hasBackdropStyle;
  dst->hasProceduralNoise |= src.hasProceduralNoise;
  dst->hasUnsupportedImagePattern |= src.hasUnsupportedImagePattern;
  dst->hasInexactNativeMapping |= src.hasInexactNativeMapping;
}

static bool GradientHasWideGamutStop(const std::vector<ColorStop*>& stops) {
  for (const auto* stop : stops) {
    if (stop && ColorIsWideGamut(stop->color)) {
      return true;
    }
  }
  return false;
}

static void ProbeColorSource(const ColorSource* source, PPTFeatureFlags* out) {
  if (source == nullptr) {
    return;
  }
  switch (source->nodeType()) {
    case NodeType::SolidColor:
      if (ColorIsWideGamut(static_cast<const SolidColor*>(source)->color)) {
        out->hasWideGamutColor = true;
      }
      break;
    case NodeType::LinearGradient:
      if (GradientHasWideGamutStop(static_cast<const LinearGradient*>(source)->colorStops)) {
        out->hasWideGamutColor = true;
      }
      break;
    case NodeType::RadialGradient:
      if (GradientHasWideGamutStop(static_cast<const RadialGradient*>(source)->colorStops)) {
        out->hasWideGamutColor = true;
      }
      {
        auto* gradient = static_cast<const RadialGradient*>(source);
        bool hasLinearTransform =
            !FloatNearlyZero(gradient->matrix.a - 1.0f) || !FloatNearlyZero(gradient->matrix.b) ||
            !FloatNearlyZero(gradient->matrix.c) || !FloatNearlyZero(gradient->matrix.d - 1.0f);
        // DrawingML path gradients can preserve the focus point, but they do not expose PAGX's
        // independent radius or an arbitrary transformed radial coordinate system. The default
        // normalized radius is the one exact common case; everything else is a fidelity bake.
        if (!gradient->fitsToGeometry || !FloatNearlyZero(gradient->radius - 0.5f) ||
            hasLinearTransform) {
          out->hasInexactNativeMapping = true;
        }
      }
      break;
    case NodeType::ConicGradient:
      out->hasConicGradient = true;
      if (GradientHasWideGamutStop(static_cast<const ConicGradient*>(source)->colorStops)) {
        out->hasWideGamutColor = true;
      }
      break;
    case NodeType::DiamondGradient:
      out->hasDiamondGradient = true;
      if (GradientHasWideGamutStop(static_cast<const DiamondGradient*>(source)->colorStops)) {
        out->hasWideGamutColor = true;
      }
      break;
    case NodeType::ImagePattern: {
      auto* pattern = static_cast<const ImagePattern*>(source);
      bool nonTiling =
          pattern->tileModeX == TileMode::Decal && pattern->tileModeY == TileMode::Decal;
      if (nonTiling &&
          (!FloatNearlyZero(pattern->matrix.b) || !FloatNearlyZero(pattern->matrix.c) ||
           pattern->matrix.a <= 0 || pattern->matrix.d <= 0)) {
        out->hasUnsupportedImagePattern = true;
      }
      break;
    }
    default:
      break;
  }
}

PPTFeatureFlags ProbeElementsFeatures(const std::vector<Element*>& elements) {
  PPTFeatureFlags out;
  for (const auto* el : elements) {
    if (el == nullptr) {
      continue;
    }
    switch (el->nodeType()) {
      case NodeType::TextPath:
        out.hasTextPath = true;
        break;
      case NodeType::TextModifier:
      case NodeType::RangeSelector:
        out.hasTextModifier = true;
        break;
      case NodeType::Fill: {
        auto* fill = static_cast<const Fill*>(el);
        if (!IsSupportedPaintBlendMode(fill->blendMode)) {
          out.hasUnsupportedBlend = true;
        }
        ProbeColorSource(fill->color, &out);
        break;
      }
      case NodeType::Stroke: {
        auto* stroke = static_cast<const Stroke*>(el);
        if (!IsSupportedPaintBlendMode(stroke->blendMode)) {
          out.hasUnsupportedBlend = true;
        }
        if (stroke->align != StrokeAlign::Center || !FloatNearlyZero(stroke->dashOffset) ||
            stroke->dashAdaptive) {
          out.hasInexactNativeMapping = true;
        }
        if (stroke->color != nullptr && stroke->color->nodeType() == NodeType::ImagePattern) {
          // DrawingML line fills do not provide the same image-pattern placement and clipping
          // model as PAGX strokes. Bake instead of emitting a misleading rectangular blip fill.
          out.hasInexactNativeMapping = true;
        }
        ProbeColorSource(stroke->color, &out);
        break;
      }
      case NodeType::Group: {
        auto* g = static_cast<const Group*>(el);
        if (!FloatNearlyZero(g->skew)) {
          out.hasShearTransform = true;
        }
        Merge(&out, ProbeElementsFeatures(g->elements));
        break;
      }
      case NodeType::TextBox: {
        auto* tb = static_cast<const TextBox*>(el);
        Merge(&out, ProbeElementsFeatures(tb->elements));
        break;
      }
      default:
        break;
    }
  }
  return out;
}

PPTFeatureFlags ProbeLayerFeatures(const Layer* layer) {
  PPTFeatureFlags out;
  if (layer == nullptr || !layer->visible) {
    return out;
  }
  if (!IsSupportedPaintBlendMode(layer->blendMode)) {
    out.hasUnsupportedBlend = true;
  }
  if (!layer->matrix.isIdentity() && MatrixHasShear(layer->matrix)) {
    out.hasShearTransform = true;
  }
  int blurCount = 0;
  int blendCount = 0;
  int innerShadowCount = 0;
  int dropShadowCount = 0;
  for (const auto* filter : layer->filters) {
    if (filter == nullptr) {
      continue;
    }
    auto type = filter->nodeType();
    if (type == NodeType::ColorMatrixFilter) {
      out.hasColorMatrix = true;
    } else if (type == NodeType::NoiseFilter) {
      out.hasProceduralNoise = true;
    } else if (type == NodeType::BlurFilter) {
      auto* blur = static_cast<const BlurFilter*>(filter);
      blurCount++;
      if (!FloatNearlyZero(blur->blurX - blur->blurY) || blur->tileMode != TileMode::Decal) {
        out.hasInexactNativeMapping = true;
      }
    } else if (type == NodeType::BlendFilter) {
      auto* blend = static_cast<const BlendFilter*>(filter);
      blendCount++;
      if (!IsSupportedBlendFilterMode(blend->blendMode)) {
        out.hasUnsupportedBlend = true;
      }
      if (ColorIsWideGamut(blend->color)) {
        out.hasWideGamutColor = true;
      }
    } else if (type == NodeType::InnerShadowFilter) {
      auto* shadow = static_cast<const InnerShadowFilter*>(filter);
      innerShadowCount++;
      if (!FloatNearlyZero(shadow->blurX - shadow->blurY) || shadow->shadowOnly) {
        out.hasInexactNativeMapping = true;
      }
      if (ColorIsWideGamut(shadow->color)) {
        out.hasWideGamutColor = true;
      }
    } else if (type == NodeType::DropShadowFilter) {
      auto* shadow = static_cast<const DropShadowFilter*>(filter);
      dropShadowCount++;
      if (!FloatNearlyZero(shadow->blurX - shadow->blurY) || shadow->shadowOnly) {
        out.hasInexactNativeMapping = true;
      }
      if (ColorIsWideGamut(shadow->color)) {
        out.hasWideGamutColor = true;
      }
    }
  }
  for (const auto* style : layer->styles) {
    if (style == nullptr) {
      continue;
    }
    if (style->nodeType() == NodeType::BackgroundBlurStyle) {
      auto* bg = static_cast<const BackgroundBlurStyle*>(style);
      // Zero-radius BackgroundBlur is a true no-op, so only flag when blur is non-zero.
      if (bg->blurX > 0 || bg->blurY > 0) {
        out.hasBackdropStyle = true;
      }
    } else if (style->nodeType() == NodeType::GlassStyle) {
      // Unlike BackgroundBlurStyle, a GlassStyle always composites the backdrop through the
      // layer's shape (frost / refraction / chromatic dispersion / edge lighting) as soon as
      // it is present, so flag it unconditionally.
      out.hasBackdropStyle = true;
    } else if (style->nodeType() == NodeType::NoiseStyle) {
      out.hasProceduralNoise = true;
    } else if (style->nodeType() == NodeType::InnerShadowStyle) {
      auto* shadow = static_cast<const InnerShadowStyle*>(style);
      innerShadowCount++;
      if (!FloatNearlyZero(shadow->blurX - shadow->blurY) ||
          shadow->blendMode != BlendMode::Normal || shadow->excludeChildEffects) {
        out.hasInexactNativeMapping = true;
      }
      if (ColorIsWideGamut(shadow->color)) {
        out.hasWideGamutColor = true;
      }
    } else if (style->nodeType() == NodeType::DropShadowStyle) {
      auto* shadow = static_cast<const DropShadowStyle*>(style);
      dropShadowCount++;
      if (!FloatNearlyZero(shadow->blurX - shadow->blurY) ||
          shadow->blendMode != BlendMode::Normal || shadow->excludeChildEffects ||
          !shadow->showBehindLayer) {
        out.hasInexactNativeMapping = true;
      }
      if (ColorIsWideGamut(shadow->color)) {
        out.hasWideGamutColor = true;
      }
    }
  }
  // DrawingML effectLst permits only one blur, one inner shadow, and one outer shadow. The
  // native writer intentionally selects the first of each kind, so multiple authored instances
  // must bake in fidelity mode instead of silently losing the remaining effects.
  if (blurCount > 1 || blendCount > 1 || innerShadowCount > 1 || dropShadowCount > 1) {
    out.hasInexactNativeMapping = true;
  }

  // Only probe the layer's own contents (groups / text boxes are emitted as
  // part of this layer in writeElements). Composition layers and child layers
  // are visited separately by writeLayer, so each one gets its own probe and
  // can be rasterized at its own scope. Aggregating descendant flags here
  // would force the smallest enclosing layer that has any unsupported feature
  // anywhere in its sub-tree to bake the entire sub-tree into one PNG, which
  // both blows up the output size and turns surrounding native content (e.g.
  // an underlying gradient) into a non-editable raster.
  Merge(&out, ProbeElementsFeatures(layer->contents));
  return out;
}

}  // namespace pagx
