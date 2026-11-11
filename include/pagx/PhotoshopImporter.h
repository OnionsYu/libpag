/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once

#include "pagx/PAGXDocument.h"
#include "photoshop/Document.h"

namespace pagx {

struct PhotoshopRasterRequest {
  std::string id;
  // Empty for the complete document composite; otherwise a source model layer ID.
  std::string layerId;
};

struct PhotoshopConversionPlan {
  bool composite = false;
  std::vector<PhotoshopRasterRequest> rasters = {};
  std::vector<photoshop::Diagnostic> diagnostics = {};
};

class PhotoshopImporter {
 public:
  // Analysis is conservative: uncertain backdrop dependencies collapse to the entire document.
  // The host executes returned requests; this library never invokes Photoshop or opens PSD files.
  static PhotoshopConversionPlan Analyze(const photoshop::Document& source,
                                         bool forceComposite = false);

  // Recomputes the plan from source/options, so stale or tampered plans cannot skip dependencies.
  // Throws std::invalid_argument on missing, duplicate, unexpected or invalid capture resources.
  static std::shared_ptr<PAGXDocument> Convert(const photoshop::Document& source,
                                               const std::vector<photoshop::RasterAsset>& assets,
                                               bool forceComposite = false);
};

}  // namespace pagx
