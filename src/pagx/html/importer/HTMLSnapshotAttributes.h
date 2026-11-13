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
//  Unless required by applicable law or agreed to in writing, software distributed under the
//  License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
//  either express or implied. See the License for the specific language governing permissions
//  and limitations under the License.
//
/////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once

#include <string>
#include "pagx/xml/XMLDOM.h"

namespace pagx::html {

// Serialized by tools/html-snapshot/lib/snapshot-scene.ts. The preview markup's generated
// elements are paint primitives, not independent source elements. Keep this distinction through
// normalization instead of inferring ownership from tag names or the number of children.
inline bool IsSnapshotAttribute(const std::string& name) {
  return name == "data-pagx-node-role" || name == "data-pagx-node-source" ||
         name == "data-pagx-node-parent";
}

inline bool HasSnapshotRole(const DOMNode* element, const std::string& role) {
  if (element == nullptr) return false;
  const auto* value = element->findAttribute("data-pagx-node-role");
  return value != nullptr && *value == role;
}

}  // namespace pagx::html
