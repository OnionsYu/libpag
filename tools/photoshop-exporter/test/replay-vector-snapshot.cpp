/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include <fstream>
#include <iostream>
#include <stdexcept>
#include "pagx/PAGXExporter.h"
#include "pagx/PAGXImporter.h"
#include "pagx/PhotoshopImporter.h"

// Offline replay of captured vector-only source snapshots against the CURRENT conversion core.
// A missing raster is a failure; never substitute an old PAGX image or pretend it is a host render.
int main(int argc, char** argv) {
  try {
    if (argc != 3)
      throw std::runtime_error("Usage: photoshop-vector-replay snapshot.json output.pagx");
    std::ifstream input(argv[1], std::ios::binary | std::ios::ate);
    if (!input || input.tellg() < 0 || input.tellg() > 32 * 1024 * 1024)
      throw std::runtime_error("Snapshot unavailable or larger than 32 MiB");
    std::string json(static_cast<size_t>(input.tellg()), '\0');
    input.seekg(0);
    if (!input.read(json.data(), static_cast<std::streamsize>(json.size())))
      throw std::runtime_error("Cannot read snapshot");
    const auto source = photoshop::ParseSnapshot(json);
    const auto plan = pagx::PhotoshopImporter::Analyze(source);
    if (plan.composite || !plan.rasters.empty())
      throw std::runtime_error(
          "Snapshot needs fresh Photoshop raster captures; vector replay cannot verify it");
    auto document = pagx::PhotoshopImporter::Convert(source, {});
    const auto xml = pagx::PAGXExporter::ToXML(*document);
    const auto parsed = pagx::PAGXImporter::FromXML(xml);
    if (!parsed || !parsed->errors.empty() || parsed->hasUnresolvedImports() ||
        !parsed->getExternalFilePaths().empty())
      throw std::runtime_error("Replay failed PAGX validation");
    std::ofstream output(argv[2], std::ios::binary | std::ios::trunc);
    if (!output.write(xml.data(), static_cast<std::streamsize>(xml.size())))
      throw std::runtime_error("Cannot write replay output");
    output.close();
    if (!output) throw std::runtime_error("Cannot finish replay output");
    std::cout << "Replayed " << plan.statistics.vectorShapes << " vector shapes and "
              << plan.statistics.outlinedTexts << " outlined texts.\n";
    return 0;
  } catch (const std::exception& error) {
    std::cerr << error.what() << '\n';
    return 1;
  }
}
