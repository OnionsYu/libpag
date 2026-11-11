/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////
#pragma once
#include "pagx/PhotoshopImporter.h"

namespace pagx {
bool IsPhotoshopGroup(const photoshop::Layer& layer);
bool IsNativePhotoshopLayer(const photoshop::Layer& layer);
bool IsEmptyPhotoshopLayer(const photoshop::Layer& layer);
PhotoshopConversionPlan PlanPhotoshopComposition(const photoshop::Document& source,
                                                 bool forceComposite);
}  // namespace pagx
