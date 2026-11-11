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
  // First replaced sibling. Empty only for the complete document composite.
  std::string layerId;
  // Contiguous source siblings in paint order, including intervening hidden/empty nodes.
  // Each root includes its complete subtree. The output replaces these roots exactly once.
  std::vector<std::string> layerIds = {};
  // Parent source group whose opacity remains in PAGX, or empty at the document root.
  std::string parentId = {};
};

struct PhotoshopExportStatistics {
  size_t vectorShapes = 0;
  size_t outlinedTexts = 0;
  size_t sourceImages = 0;
  size_t fallbackScopes = 0;
};

struct PhotoshopConversionPlan {
  bool composite = false;
  std::vector<PhotoshopRasterRequest> rasters = {};
  std::vector<photoshop::Diagnostic> diagnostics = {};
  PhotoshopExportStatistics statistics = {};
};

class PhotoshopImporter {
 public:
  // Backdrop dependencies collapse to a contiguous sibling range inside the nearest proven
  // isolation boundary. Unknown composition and escaping knockout still require the document.
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
