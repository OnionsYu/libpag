/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////
#pragma once
#include "pagx/PAGXDocument.h"
#include "photoshop/Document.h"

namespace pagx {
bool CanBuildShape(const photoshop::Shape& shape);
bool CanBuildVectorMask(const std::optional<photoshop::VectorMask>& mask);
void BuildPhotoshopVectorMask(PAGXDocument* document, Layer* layer,
                              const photoshop::VectorMask& mask, const photoshop::Rect& canvas);
void BuildPhotoshopShape(PAGXDocument* document, Layer* layer, const photoshop::Shape& shape,
                         float fillOpacity);
}  // namespace pagx
