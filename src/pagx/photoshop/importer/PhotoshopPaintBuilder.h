/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////
#pragma once
#include "pagx/PAGXDocument.h"
#include "pagx/nodes/ColorSource.h"
#include "photoshop/Document.h"

namespace pagx {
bool CanBuildPhotoshopPaint(const photoshop::Paint& paint);
ColorSource* BuildPhotoshopPaint(PAGXDocument* document, const photoshop::Paint& paint);
}  // namespace pagx
