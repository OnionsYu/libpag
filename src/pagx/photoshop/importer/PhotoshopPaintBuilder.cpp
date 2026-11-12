/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////
#include "PhotoshopPaintBuilder.h"
#include <stdexcept>
#include "pagx/nodes/SolidColor.h"

namespace pagx {
bool CanBuildPhotoshopPaint(const photoshop::Paint& paint) {
  return !paint.gradient;
}
ColorSource* BuildPhotoshopPaint(PAGXDocument* document, const photoshop::Paint& paint) {
  if (!CanBuildPhotoshopPaint(paint)) throw std::invalid_argument("Unqualified Photoshop paint");
  auto color = document->makeNode<SolidColor>();
  const auto& value = paint.color;
  color->color = {value.red, value.green, value.blue, value.alpha};
  return color;
}
}  // namespace pagx
