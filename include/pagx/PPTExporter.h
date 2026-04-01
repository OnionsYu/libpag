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

#include <string>
#include "pagx/PAGXDocument.h"

namespace pagx {

/**
 * Export options for PPTExporter.
 */
struct PPTExportOptions {
  /**
   * Indentation spaces for the generated XML inside the PPTX. The default value is 2.
   */
  int indent = 2;
};

/**
 * PPTExporter converts a PAGXDocument into PPTX (PowerPoint) format.
 * Each top-level layer in the document becomes a separate slide.
 * Supported elements: Rectangle, Ellipse, Path, Text, Group, Image.
 * Supported fills: SolidColor, LinearGradient, RadialGradient, ImagePattern.
 * Supported strokes: SolidColor with width, cap, join, dashes.
 */
class PPTExporter {
 public:
  using Options = PPTExportOptions;

  /**
   * Exports a PAGXDocument to a PPTX file.
   * Returns true on success.
   */
  static bool ToFile(const PAGXDocument& document, const std::string& filePath,
                     const Options& options = {});
};

}  // namespace pagx
