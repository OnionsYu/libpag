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

#pragma once

#include <cmath>
#include <memory>
#include <string>
#include "pagx/PAGXDocument.h"

namespace pagx {

/**
 * HTMLImporter converts a constrained subset of HTML/CSS into a PAGX Document.
 *
 * This converter is intentionally lossy: it accepts only a flex-based, well-formed XHTML subset
 * that maps cleanly onto PAGX's container layout model (Layer + layout/gap/padding/flex/alignment/
 * arrangement). Anything outside the subset (CSS Grid, float, pseudo-elements, animations, complex
 * cascade, non-px/% units, etc.) is dropped or approximated and a warning is appended to
 * `PAGXDocument::errors`. Inline `<svg>` content is preserved as an unresolved import directive on
 * the enclosing Layer and can be expanded later via `pagx resolve`.
 *
 * The importer does not run layout itself — it only translates the element tree into the PAGX node
 * graph. Call `PAGXDocument::applyLayout()` afterwards (or use the `pagx` CLI) to position content.
 */
class HTMLImporter {
 public:
  struct Options {
    /**
     * If true, unsupported elements are preserved as empty Layers carrying their original tag in
     * `customData["html-tag"]`. If false (default), unsupported elements are dropped and a warning
     * is emitted.
     */
    bool preserveUnknownElements = false;

    /**
     * Target width for the output document. When not NaN, overrides the canvas width inferred from
     * the root `<html>` element's style. Both targetWidth and targetHeight must be set (non-NaN) to
     * take effect.
     */
    float targetWidth = NAN;

    /**
     * Target height for the output document. When not NaN, overrides the canvas height inferred
     * from the root `<html>` element's style. Both targetWidth and targetHeight must be set
     * (non-NaN) to take effect.
     */
    float targetHeight = NAN;

    /**
     * Default canvas width used when neither the root `<html>` element specifies a width nor
     * `targetWidth` is provided. The default value is 800.
     */
    float defaultWidth = 800.0f;

    /**
     * Default canvas height used when neither the root `<html>` element specifies a height nor
     * `targetHeight` is provided. The default value is 600.
     */
    float defaultHeight = 600.0f;

    Options() {
    }
  };

  /**
   * Parses an HTML file and creates a PAGX Document. Returns nullptr if the file cannot be loaded
   * or parsing fails fatally. Non-fatal issues are appended to `PAGXDocument::errors`.
   */
  static std::shared_ptr<PAGXDocument> Parse(const std::string& filePath,
                                             const Options& options = Options());

  /**
   * Parses HTML data and creates a PAGX Document.
   */
  static std::shared_ptr<PAGXDocument> Parse(const uint8_t* data, size_t length,
                                             const Options& options = Options());

  /**
   * Parses an HTML string and creates a PAGX Document.
   */
  static std::shared_ptr<PAGXDocument> ParseString(const std::string& htmlContent,
                                                   const Options& options = Options());
};

}  // namespace pagx
