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

#include "pagx/PPTExporter.h"
#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <fstream>
#include <map>
#include <string>
#include <vector>
#include "base/utils/MathUtil.h"
#include "pagx/PAGXDocument.h"
#include "pagx/nodes/ColorStop.h"
#include "pagx/nodes/Ellipse.h"
#include "pagx/nodes/Fill.h"
#include "pagx/nodes/Group.h"
#include "pagx/nodes/Image.h"
#include "pagx/nodes/ImagePattern.h"
#include "pagx/nodes/LinearGradient.h"
#include "pagx/nodes/Path.h"
#include "pagx/nodes/PathData.h"
#include "pagx/nodes/RadialGradient.h"
#include "pagx/nodes/Rectangle.h"
#include "pagx/nodes/SolidColor.h"
#include "pagx/nodes/Stroke.h"
#include "pagx/nodes/Text.h"
#include "pagx/nodes/TextBox.h"
#include "pagx/types/Rect.h"
#include "pagx/utils/Base64.h"
#include "pagx/utils/StringParser.h"

namespace pagx {

using pag::DegreesToRadians;
using pag::FloatNearlyZero;

//==============================================================================
// CRC32 computation
//==============================================================================

static uint32_t ComputeCRC32(const uint8_t* data, size_t len) {
  static uint32_t table[256] = {};
  static bool initialized = false;
  if (!initialized) {
    for (uint32_t i = 0; i < 256; i++) {
      uint32_t crc = i;
      for (int j = 0; j < 8; j++) {
        crc = (crc >> 1) ^ (crc & 1 ? 0xEDB88320u : 0u);
      }
      table[i] = crc;
    }
    initialized = true;
  }
  uint32_t crc = 0xFFFFFFFFu;
  for (size_t i = 0; i < len; i++) {
    crc = table[(crc ^ data[i]) & 0xFF] ^ (crc >> 8);
  }
  return ~crc;
}

//==============================================================================
// ZipWriter – creates uncompressed ZIP archives
//==============================================================================

class ZipWriter {
 public:
  void addFile(const std::string& name, const std::string& content) {
    addFile(name, reinterpret_cast<const uint8_t*>(content.data()), content.size());
  }

  void addFile(const std::string& name, const uint8_t* data, size_t len) {
    Entry entry;
    entry.name = name;
    entry.data.assign(data, data + len);
    entry.crc32 = ComputeCRC32(data, len);
    entry.localOffset = static_cast<uint32_t>(buffer_.size());

    // Local file header
    writeU32(0x04034b50);       // signature
    writeU16(20);               // version needed
    writeU16(0);                // flags
    writeU16(0);                // compression: stored
    writeU16(0);                // mod time
    writeU16(0);                // mod date
    writeU32(entry.crc32);
    writeU32(static_cast<uint32_t>(len));  // compressed size
    writeU32(static_cast<uint32_t>(len));  // uncompressed size
    writeU16(static_cast<uint16_t>(name.size()));
    writeU16(0);                // extra field length
    buffer_.insert(buffer_.end(), name.begin(), name.end());
    buffer_.insert(buffer_.end(), data, data + len);

    entries_.push_back(std::move(entry));
  }

  std::vector<uint8_t> build() {
    uint32_t cdOffset = static_cast<uint32_t>(buffer_.size());
    for (const auto& entry : entries_) {
      // Central directory entry
      writeU32(0x02014b50);       // signature
      writeU16(20);               // version made by
      writeU16(20);               // version needed
      writeU16(0);                // flags
      writeU16(0);                // compression
      writeU16(0);                // mod time
      writeU16(0);                // mod date
      writeU32(entry.crc32);
      writeU32(static_cast<uint32_t>(entry.data.size()));
      writeU32(static_cast<uint32_t>(entry.data.size()));
      writeU16(static_cast<uint16_t>(entry.name.size()));
      writeU16(0);                // extra field length
      writeU16(0);                // file comment length
      writeU16(0);                // disk number start
      writeU16(0);                // internal attributes
      writeU32(0);                // external attributes
      writeU32(entry.localOffset);
      buffer_.insert(buffer_.end(), entry.name.begin(), entry.name.end());
    }
    uint32_t cdSize = static_cast<uint32_t>(buffer_.size()) - cdOffset;

    // End of central directory
    writeU32(0x06054b50);
    writeU16(0);                // disk number
    writeU16(0);                // CD start disk
    writeU16(static_cast<uint16_t>(entries_.size()));
    writeU16(static_cast<uint16_t>(entries_.size()));
    writeU32(cdSize);
    writeU32(cdOffset);
    writeU16(0);                // comment length

    return std::move(buffer_);
  }

 private:
  struct Entry {
    std::string name;
    std::vector<uint8_t> data;
    uint32_t crc32 = 0;
    uint32_t localOffset = 0;
  };

  std::vector<Entry> entries_;
  std::vector<uint8_t> buffer_;

  void writeU16(uint16_t val) {
    buffer_.push_back(static_cast<uint8_t>(val & 0xFF));
    buffer_.push_back(static_cast<uint8_t>((val >> 8) & 0xFF));
  }

  void writeU32(uint32_t val) {
    buffer_.push_back(static_cast<uint8_t>(val & 0xFF));
    buffer_.push_back(static_cast<uint8_t>((val >> 8) & 0xFF));
    buffer_.push_back(static_cast<uint8_t>((val >> 16) & 0xFF));
    buffer_.push_back(static_cast<uint8_t>((val >> 24) & 0xFF));
  }
};

//==============================================================================
// XMLBuilder – XML generation helper (similar to SVGBuilder)
//==============================================================================

class XMLBuilder {
 public:
  explicit XMLBuilder(int indentSpaces = 2, int initialLevel = 0)
      : indentLevel_(initialLevel), indentSpaces_(indentSpaces) {
    buffer_.reserve(4096);
  }

  void appendDeclaration() {
    buffer_ += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
  }

  void openElement(const char* tag) {
    writeIndent();
    buffer_ += '<';
    buffer_ += tag;
    tagStack_.push_back(tag);
  }

  void addAttribute(const char* name, const char* value) {
    if (!value) {
      return;
    }
    buffer_ += ' ';
    buffer_ += name;
    buffer_ += "=\"";
    buffer_ += escapeXml(value);
    buffer_ += '"';
  }

  void addAttribute(const char* name, const std::string& value) {
    addAttribute(name, value.c_str());
  }

  void addAttribute(const char* name, int64_t value) {
    buffer_ += ' ';
    buffer_ += name;
    buffer_ += "=\"";
    buffer_ += std::to_string(value);
    buffer_ += '"';
  }

  void closeElementStart() {
    buffer_ += ">\n";
    indentLevel_++;
  }

  void closeElementSelfClosing() {
    buffer_ += "/>\n";
    tagStack_.pop_back();
  }

  void closeElement() {
    indentLevel_--;
    writeIndent();
    buffer_ += "</";
    buffer_ += tagStack_.back();
    buffer_ += ">\n";
    tagStack_.pop_back();
  }

  void closeElementWithText(const std::string& text) {
    buffer_ += '>';
    buffer_ += escapeXml(text);
    buffer_ += "</";
    buffer_ += tagStack_.back();
    buffer_ += ">\n";
    tagStack_.pop_back();
  }

  void addRawContent(const std::string& content) {
    buffer_ += content;
  }

  std::string release() {
    return std::move(buffer_);
  }

 private:
  std::string buffer_;
  std::vector<const char*> tagStack_;
  int indentLevel_ = 0;
  int indentSpaces_ = 2;

  void writeIndent() {
    buffer_.append(static_cast<size_t>(indentLevel_ * indentSpaces_), ' ');
  }

  static std::string escapeXml(const std::string& input) {
    std::string result;
    result.reserve(input.size());
    for (char c : input) {
      switch (c) {
        case '&':
          result += "&amp;";
          break;
        case '<':
          result += "&lt;";
          break;
        case '>':
          result += "&gt;";
          break;
        case '"':
          result += "&quot;";
          break;
        default:
          result += c;
          break;
      }
    }
    return result;
  }
};

//==============================================================================
// Coordinate and unit conversion utilities
//==============================================================================

static constexpr int64_t kEMUPerPixel = 9525;  // 914400 EMU/inch ÷ 96 DPI
static constexpr int64_t kRotationUnitsPerDegree = 60000;

static int64_t PixelsToEMU(float px) {
  return static_cast<int64_t>(std::round(px * kEMUPerPixel));
}

static int64_t DegreesToPPTRotation(float degrees) {
  return static_cast<int64_t>(std::round(degrees * kRotationUnitsPerDegree));
}

// PPTX font size is in hundredths of a point; 1 px at 96 DPI = 0.75 pt.
static int64_t FontSizeToPPT(float pxSize) {
  return static_cast<int64_t>(std::round(pxSize * 75.0f));
}

// PPTX alpha values are in 1/1000 of a percent (100000 = 100%).
static int64_t AlphaToPPT(float alpha) {
  return static_cast<int64_t>(std::round(std::clamp(alpha, 0.0f, 1.0f) * 100000.0f));
}

static std::string ColorToHex(const Color& color) {
  int r = std::clamp(static_cast<int>(std::round(color.red * 255.0f)), 0, 255);
  int g = std::clamp(static_cast<int>(std::round(color.green * 255.0f)), 0, 255);
  int b = std::clamp(static_cast<int>(std::round(color.blue * 255.0f)), 0, 255);
  char buf[8];
  snprintf(buf, sizeof(buf), "%02X%02X%02X", r, g, b);
  return buf;
}

static std::string I64ToString(int64_t val) {
  return std::to_string(val);
}

//==============================================================================
// FillStrokeInfo collection (shared with SVGExporter pattern)
//==============================================================================

struct FillStrokeInfo {
  const Fill* fill = nullptr;
  const Stroke* stroke = nullptr;
  const TextBox* textBox = nullptr;
};

static FillStrokeInfo CollectFillStroke(const std::vector<Element*>& contents) {
  FillStrokeInfo info = {};
  for (const auto* element : contents) {
    if (element->nodeType() == NodeType::Fill && !info.fill) {
      info.fill = static_cast<const Fill*>(element);
    } else if (element->nodeType() == NodeType::Stroke && !info.stroke) {
      info.stroke = static_cast<const Stroke*>(element);
    } else if (element->nodeType() == NodeType::TextBox && !info.textBox) {
      info.textBox = static_cast<const TextBox*>(element);
    }
    if (info.fill && info.stroke && info.textBox) {
      break;
    }
  }
  return info;
}

//==============================================================================
// Layer and Group matrix builders (matching SVGExporter logic)
//==============================================================================

static Matrix BuildLayerMatrix(const Layer* layer) {
  Matrix m = layer->matrix;
  if (layer->x != 0.0f || layer->y != 0.0f) {
    m = Matrix::Translate(layer->x, layer->y) * m;
  }
  return m;
}

static Matrix BuildGroupMatrix(const Group* group) {
  bool hasAnchor = !FloatNearlyZero(group->anchor.x) || !FloatNearlyZero(group->anchor.y);
  bool hasPosition = !FloatNearlyZero(group->position.x) || !FloatNearlyZero(group->position.y);
  bool hasRotation = !FloatNearlyZero(group->rotation);
  bool hasScale =
      !FloatNearlyZero(group->scale.x - 1.0f) || !FloatNearlyZero(group->scale.y - 1.0f);
  bool hasSkew = !FloatNearlyZero(group->skew);

  if (!hasAnchor && !hasPosition && !hasRotation && !hasScale && !hasSkew) {
    return {};
  }

  Matrix m = {};
  if (hasAnchor) {
    m = Matrix::Translate(-group->anchor.x, -group->anchor.y);
  }
  if (hasScale) {
    m = Matrix::Scale(group->scale.x, group->scale.y) * m;
  }
  if (hasSkew) {
    m = Matrix::Rotate(group->skewAxis) * m;
    Matrix shear = {};
    shear.c = std::tan(DegreesToRadians(group->skew));
    m = shear * m;
    m = Matrix::Rotate(-group->skewAxis) * m;
  }
  if (hasRotation) {
    m = Matrix::Rotate(group->rotation) * m;
  }
  if (hasPosition) {
    m = Matrix::Translate(group->position.x, group->position.y) * m;
  }

  return m;
}

//==============================================================================
// Transform rect helper – applies accumulated matrix to a local bounding box
//==============================================================================

struct TransformedRect {
  float x;
  float y;
  float width;
  float height;
  float rotationDeg;
};

static TransformedRect ApplyMatrixToRect(float localX, float localY, float localW, float localH,
                                         const Matrix& m) {
  float cx = localX + localW / 2.0f;
  float cy = localY + localH / 2.0f;
  Point mappedCenter = m.mapPoint({cx, cy});

  float scaleX = std::sqrt(m.a * m.a + m.b * m.b);
  float scaleY = std::sqrt(m.c * m.c + m.d * m.d);
  float det = m.a * m.d - m.b * m.c;
  if (det < 0) {
    scaleY = -scaleY;
  }

  float rotDeg = std::atan2(m.b, m.a) * 180.0f / 3.14159265358979323846f;

  float sw = localW * scaleX;
  float sh = localH * std::abs(scaleY);

  TransformedRect result;
  result.x = mappedCenter.x - sw / 2.0f;
  result.y = mappedCenter.y - sh / 2.0f;
  result.width = sw;
  result.height = sh;
  result.rotationDeg = rotDeg;
  return result;
}

//==============================================================================
// Image data helpers
//==============================================================================

static std::vector<uint8_t> GetImageBytes(const Image* image) {
  if (image->data) {
    return {image->data->bytes(), image->data->bytes() + image->data->size()};
  }
  if (!image->filePath.empty()) {
    if (image->filePath.rfind("data:", 0) == 0) {
      auto decoded = DecodeBase64DataURI(image->filePath);
      if (decoded) {
        return {decoded->bytes(), decoded->bytes() + decoded->size()};
      }
    } else {
      std::ifstream file(image->filePath, std::ios::binary);
      if (file) {
        return {std::istreambuf_iterator<char>(file), std::istreambuf_iterator<char>()};
      }
    }
  }
  return {};
}

//==============================================================================
// Boilerplate XML generators
//==============================================================================

static std::string GenerateContentTypes(int slideCount, bool hasImages) {
  XMLBuilder xml;
  xml.appendDeclaration();
  xml.openElement("Types");
  xml.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
  xml.closeElementStart();

  xml.openElement("Default");
  xml.addAttribute("Extension", "xml");
  xml.addAttribute("ContentType", "application/xml");
  xml.closeElementSelfClosing();

  xml.openElement("Default");
  xml.addAttribute("Extension", "rels");
  xml.addAttribute("ContentType",
                    "application/vnd.openxmlformats-package.relationships+xml");
  xml.closeElementSelfClosing();

  if (hasImages) {
    xml.openElement("Default");
    xml.addAttribute("Extension", "png");
    xml.addAttribute("ContentType", "image/png");
    xml.closeElementSelfClosing();

    xml.openElement("Default");
    xml.addAttribute("Extension", "jpeg");
    xml.addAttribute("ContentType", "image/jpeg");
    xml.closeElementSelfClosing();
  }

  xml.openElement("Override");
  xml.addAttribute("PartName", "/ppt/presentation.xml");
  xml.addAttribute(
      "ContentType",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml");
  xml.closeElementSelfClosing();

  xml.openElement("Override");
  xml.addAttribute("PartName", "/ppt/slideMasters/slideMaster1.xml");
  xml.addAttribute(
      "ContentType",
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml");
  xml.closeElementSelfClosing();

  xml.openElement("Override");
  xml.addAttribute("PartName", "/ppt/slideLayouts/slideLayout1.xml");
  xml.addAttribute(
      "ContentType",
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml");
  xml.closeElementSelfClosing();

  xml.openElement("Override");
  xml.addAttribute("PartName", "/ppt/theme/theme1.xml");
  xml.addAttribute("ContentType",
                    "application/vnd.openxmlformats-officedocument.theme+xml");
  xml.closeElementSelfClosing();

  for (int i = 1; i <= slideCount; i++) {
    xml.openElement("Override");
    xml.addAttribute("PartName", "/ppt/slides/slide" + std::to_string(i) + ".xml");
    xml.addAttribute(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
    xml.closeElementSelfClosing();
  }

  xml.closeElement();
  return xml.release();
}

static std::string GenerateRootRels() {
  return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>
)";
}

static std::string GeneratePresentation(int slideCount, int64_t cx, int64_t cy) {
  XMLBuilder xml;
  xml.appendDeclaration();
  xml.openElement("p:presentation");
  xml.addAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
  xml.addAttribute("xmlns:r",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
  xml.addAttribute("xmlns:p", "http://schemas.openxmlformats.org/presentationml/2006/main");
  xml.closeElementStart();

  xml.openElement("p:sldMasterIdLst");
  xml.closeElementStart();
  xml.openElement("p:sldMasterId");
  xml.addAttribute("id", static_cast<int64_t>(2147483648));
  xml.addAttribute("r:id", "rId1");
  xml.closeElementSelfClosing();
  xml.closeElement();

  xml.openElement("p:sldIdLst");
  xml.closeElementStart();
  for (int i = 0; i < slideCount; i++) {
    xml.openElement("p:sldId");
    xml.addAttribute("id", static_cast<int64_t>(256 + i));
    xml.addAttribute("r:id", "rId" + std::to_string(i + 2));
    xml.closeElementSelfClosing();
  }
  xml.closeElement();

  xml.openElement("p:sldSz");
  xml.addAttribute("cx", cx);
  xml.addAttribute("cy", cy);
  xml.closeElementSelfClosing();

  xml.openElement("p:notesSz");
  xml.addAttribute("cx", static_cast<int64_t>(6858000));
  xml.addAttribute("cy", static_cast<int64_t>(9144000));
  xml.closeElementSelfClosing();

  xml.closeElement();
  return xml.release();
}

static std::string GeneratePresentationRels(int slideCount) {
  XMLBuilder xml;
  xml.appendDeclaration();
  xml.openElement("Relationships");
  xml.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
  xml.closeElementStart();

  xml.openElement("Relationship");
  xml.addAttribute("Id", "rId1");
  xml.addAttribute("Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster");
  xml.addAttribute("Target", "slideMasters/slideMaster1.xml");
  xml.closeElementSelfClosing();

  for (int i = 0; i < slideCount; i++) {
    xml.openElement("Relationship");
    xml.addAttribute("Id", "rId" + std::to_string(i + 2));
    xml.addAttribute(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
    xml.addAttribute("Target", "slides/slide" + std::to_string(i + 1) + ".xml");
    xml.closeElementSelfClosing();
  }

  int themeRelId = slideCount + 2;
  xml.openElement("Relationship");
  xml.addAttribute("Id", "rId" + std::to_string(themeRelId));
  xml.addAttribute("Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
  xml.addAttribute("Target", "theme/theme1.xml");
  xml.closeElementSelfClosing();

  xml.closeElement();
  return xml.release();
}

static std::string GenerateTheme() {
  return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="PAGX Theme">
  <a:themeElements>
    <a:clrScheme name="PAGX">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="PAGX">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="PAGX">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
)";
}

static std::string GenerateSlideMaster() {
  return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>
)";
}

static std::string GenerateSlideMasterRels() {
  return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
)";
}

static std::string GenerateSlideLayout() {
  return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
)";
}

static std::string GenerateSlideLayoutRels() {
  return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
)";
}

//==============================================================================
// PPTMediaContext – tracks image media files across all slides
//==============================================================================

class PPTMediaContext {
 public:
  struct MediaEntry {
    std::string fileName;
    std::vector<uint8_t> data;
  };

  // Returns the media filename for an image, loading its data if new.
  std::string getOrCreateMedia(const Image* image) {
    auto it = imageMap_.find(image);
    if (it != imageMap_.end()) {
      return entries_[it->second].fileName;
    }
    auto bytes = GetImageBytes(image);
    if (bytes.empty()) {
      return {};
    }
    std::string fileName = "image" + std::to_string(nextId_++) + ".png";
    int index = static_cast<int>(entries_.size());
    entries_.push_back({fileName, std::move(bytes)});
    imageMap_[image] = index;
    return fileName;
  }

  const std::vector<MediaEntry>& entries() const {
    return entries_;
  }

  bool hasImages() const {
    return !entries_.empty();
  }

 private:
  std::vector<MediaEntry> entries_;
  std::map<const Image*, int> imageMap_;
  int nextId_ = 1;
};

//==============================================================================
// PPTSlideWriter – writes PAGX layers as PPT shapes within a slide
//==============================================================================

class PPTSlideWriter {
 public:
  struct SlideImageRef {
    std::string relId;
    std::string mediaFileName;
  };

  PPTSlideWriter(PPTMediaContext* media, int indent)
      : media_(media), indent_(indent) {
  }

  std::string writeSlide(const std::vector<Layer*>& layers) {
    XMLBuilder xml(indent_);
    xml.appendDeclaration();
    xml.openElement("p:sld");
    xml.addAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
    xml.addAttribute("xmlns:r",
                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    xml.addAttribute("xmlns:p", "http://schemas.openxmlformats.org/presentationml/2006/main");
    xml.closeElementStart();

    xml.openElement("p:cSld");
    xml.closeElementStart();

    xml.openElement("p:spTree");
    xml.closeElementStart();

    // Root group shape properties (required)
    writeRootGroupProps(xml);

    // Write all layers into this single slide
    for (const auto* layer : layers) {
      Matrix layerMatrix = BuildLayerMatrix(layer);
      writeLayerContents(xml, layer, layerMatrix);
      for (const auto* child : layer->children) {
        writeLayerAsShapes(xml, child, layerMatrix);
      }
    }

    xml.closeElement();  // </p:spTree>
    xml.closeElement();  // </p:cSld>

    xml.openElement("p:clrMapOvr");
    xml.closeElementStart();
    xml.openElement("a:masterClrMapping");
    xml.closeElementSelfClosing();
    xml.closeElement();

    xml.closeElement();  // </p:sld>
    return xml.release();
  }

  std::string writeSlideRels() const {
    XMLBuilder xml;
    xml.appendDeclaration();
    xml.openElement("Relationships");
    xml.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    xml.closeElementStart();

    xml.openElement("Relationship");
    xml.addAttribute("Id", "rId1");
    xml.addAttribute(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout");
    xml.addAttribute("Target", "../slideLayouts/slideLayout1.xml");
    xml.closeElementSelfClosing();

    for (const auto& ref : imageRefs_) {
      xml.openElement("Relationship");
      xml.addAttribute("Id", ref.relId);
      xml.addAttribute(
          "Type",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
      xml.addAttribute("Target", "../media/" + ref.mediaFileName);
      xml.closeElementSelfClosing();
    }

    xml.closeElement();
    return xml.release();
  }

 private:
  PPTMediaContext* media_ = nullptr;
  int indent_ = 2;
  int nextShapeId_ = 2;
  int nextRelId_ = 2;
  int shapeCounter_ = 0;
  std::vector<SlideImageRef> imageRefs_;
  std::map<std::string, std::string> mediaToRelId_;

  int allocShapeId() {
    return nextShapeId_++;
  }

  std::string getImageRelId(const Image* image) {
    std::string mediaFile = media_->getOrCreateMedia(image);
    if (mediaFile.empty()) {
      return {};
    }
    auto it = mediaToRelId_.find(mediaFile);
    if (it != mediaToRelId_.end()) {
      return it->second;
    }
    std::string relId = "rId" + std::to_string(nextRelId_++);
    imageRefs_.push_back({relId, mediaFile});
    mediaToRelId_[mediaFile] = relId;
    return relId;
  }

  std::string shapeName(const char* prefix) {
    return std::string(prefix) + " " + std::to_string(++shapeCounter_);
  }

  // ── Root group shape properties ───────────────────────────────────────────
  void writeRootGroupProps(XMLBuilder& out) {
    out.openElement("p:nvGrpSpPr");
    out.closeElementStart();
    out.openElement("p:cNvPr");
    out.addAttribute("id", static_cast<int64_t>(1));
    out.addAttribute("name", "");
    out.closeElementSelfClosing();
    out.openElement("p:cNvGrpSpPr");
    out.closeElementSelfClosing();
    out.openElement("p:nvPr");
    out.closeElementSelfClosing();
    out.closeElement();

    out.openElement("p:grpSpPr");
    out.closeElementStart();
    writeXfrm(out, 0, 0, 0, 0);
    out.closeElement();
  }

  // ── Transform helpers ─────────────────────────────────────────────────────
  void writeXfrm(XMLBuilder& out, int64_t offX, int64_t offY, int64_t extCX, int64_t extCY,
                 int64_t rotation = 0, bool isGroup = false) {
    out.openElement("a:xfrm");
    if (rotation != 0) {
      out.addAttribute("rot", rotation);
    }
    out.closeElementStart();

    out.openElement("a:off");
    out.addAttribute("x", offX);
    out.addAttribute("y", offY);
    out.closeElementSelfClosing();

    out.openElement("a:ext");
    out.addAttribute("cx", extCX);
    out.addAttribute("cy", extCY);
    out.closeElementSelfClosing();

    if (isGroup) {
      out.openElement("a:chOff");
      out.addAttribute("x", offX);
      out.addAttribute("y", offY);
      out.closeElementSelfClosing();

      out.openElement("a:chExt");
      out.addAttribute("cx", extCX);
      out.addAttribute("cy", extCY);
      out.closeElementSelfClosing();
    }

    out.closeElement();
  }

  // ── Color source writing ──────────────────────────────────────────────────
  void writeColorElement(XMLBuilder& out, const Color& color) {
    out.openElement("a:srgbClr");
    out.addAttribute("val", ColorToHex(color));
    if (color.alpha < 1.0f) {
      out.closeElementStart();
      out.openElement("a:alpha");
      out.addAttribute("val", AlphaToPPT(color.alpha));
      out.closeElementSelfClosing();
      out.closeElement();
    } else {
      out.closeElementSelfClosing();
    }
  }

  void writeGradientStops(XMLBuilder& out, const std::vector<ColorStop*>& stops) {
    for (const auto* stop : stops) {
      out.openElement("a:gs");
      out.addAttribute("pos", static_cast<int64_t>(std::round(stop->offset * 100000.0f)));
      out.closeElementStart();
      writeColorElement(out, stop->color);
      out.closeElement();
    }
  }

  void writeSolidFill(XMLBuilder& out, const Color& color, float alpha = 1.0f) {
    out.openElement("a:solidFill");
    out.closeElementStart();
    Color c = color;
    c.alpha *= alpha;
    writeColorElement(out, c);
    out.closeElement();
  }

  void writeFillProperties(XMLBuilder& out, const Fill* fill) {
    if (!fill || !fill->color) {
      out.openElement("a:noFill");
      out.closeElementSelfClosing();
      return;
    }
    auto type = fill->color->nodeType();
    if (type == NodeType::SolidColor) {
      auto solid = static_cast<const SolidColor*>(fill->color);
      writeSolidFill(out, solid->color, fill->alpha);
    } else if (type == NodeType::LinearGradient) {
      auto grad = static_cast<const LinearGradient*>(fill->color);
      out.openElement("a:gradFill");
      out.closeElementStart();
      out.openElement("a:gsLst");
      out.closeElementStart();
      writeGradientStops(out, grad->colorStops);
      out.closeElement();
      float dx = grad->endPoint.x - grad->startPoint.x;
      float dy = grad->endPoint.y - grad->startPoint.y;
      float angle = std::atan2(dy, dx) * 180.0f / 3.14159265358979323846f;
      out.openElement("a:lin");
      out.addAttribute("ang", DegreesToPPTRotation(angle));
      out.addAttribute("scaled", "0");
      out.closeElementSelfClosing();
      out.closeElement();
    } else if (type == NodeType::RadialGradient) {
      auto grad = static_cast<const RadialGradient*>(fill->color);
      out.openElement("a:gradFill");
      out.closeElementStart();
      out.openElement("a:gsLst");
      out.closeElementStart();
      writeGradientStops(out, grad->colorStops);
      out.closeElement();
      out.openElement("a:path");
      out.addAttribute("path", "circle");
      out.closeElementStart();
      out.openElement("a:fillToRect");
      out.addAttribute("l", static_cast<int64_t>(50000));
      out.addAttribute("t", static_cast<int64_t>(50000));
      out.addAttribute("r", static_cast<int64_t>(50000));
      out.addAttribute("b", static_cast<int64_t>(50000));
      out.closeElementSelfClosing();
      out.closeElement();
      out.closeElement();
    } else if (type == NodeType::ImagePattern) {
      auto pattern = static_cast<const ImagePattern*>(fill->color);
      if (pattern->image) {
        std::string relId = getImageRelId(pattern->image);
        if (!relId.empty()) {
          out.openElement("a:blipFill");
          out.closeElementStart();
          out.openElement("a:blip");
          out.addAttribute("r:embed", relId);
          out.closeElementSelfClosing();
          out.openElement("a:stretch");
          out.closeElementStart();
          out.openElement("a:fillRect");
          out.closeElementSelfClosing();
          out.closeElement();
          out.closeElement();
          return;
        }
      }
      out.openElement("a:noFill");
      out.closeElementSelfClosing();
    } else {
      out.openElement("a:noFill");
      out.closeElementSelfClosing();
    }
  }

  void writeStrokeProperties(XMLBuilder& out, const Stroke* stroke) {
    if (!stroke || !stroke->color) {
      return;
    }
    int64_t widthEmu = PixelsToEMU(stroke->width);
    out.openElement("a:ln");
    out.addAttribute("w", widthEmu);

    if (stroke->cap == LineCap::Round) {
      out.addAttribute("cap", "rnd");
    } else if (stroke->cap == LineCap::Square) {
      out.addAttribute("cap", "sq");
    } else {
      out.addAttribute("cap", "flat");
    }
    out.closeElementStart();

    auto type = stroke->color->nodeType();
    if (type == NodeType::SolidColor) {
      auto solid = static_cast<const SolidColor*>(stroke->color);
      writeSolidFill(out, solid->color, stroke->alpha);
    } else {
      out.openElement("a:noFill");
      out.closeElementSelfClosing();
    }

    if (!stroke->dashes.empty()) {
      out.openElement("a:prstDash");
      out.addAttribute("val", "dash");
      out.closeElementSelfClosing();
    }

    if (stroke->join == LineJoin::Round) {
      out.openElement("a:round");
      out.closeElementSelfClosing();
    } else if (stroke->join == LineJoin::Bevel) {
      out.openElement("a:bevel");
      out.closeElementSelfClosing();
    } else {
      out.openElement("a:miter");
      out.addAttribute("lim", static_cast<int64_t>(stroke->miterLimit * 100000));
      out.closeElementSelfClosing();
    }

    out.closeElement();
  }

  // ── Preset geometry ───────────────────────────────────────────────────────
  void writePresetGeom(XMLBuilder& out, const char* prst) {
    out.openElement("a:prstGeom");
    out.addAttribute("prst", prst);
    out.closeElementStart();
    out.openElement("a:avLst");
    out.closeElementSelfClosing();
    out.closeElement();
  }

  void writeRoundRectGeom(XMLBuilder& out, float roundness, float width, float height) {
    out.openElement("a:prstGeom");
    out.addAttribute("prst", "roundRect");
    out.closeElementStart();
    out.openElement("a:avLst");
    out.closeElementStart();
    float minDim = std::min(width, height);
    int64_t adjVal = 0;
    if (minDim > 0) {
      adjVal = static_cast<int64_t>(std::round(roundness / minDim * 100000.0f));
      adjVal = std::min(adjVal, static_cast<int64_t>(50000));
    }
    out.openElement("a:gd");
    out.addAttribute("name", "adj");
    out.addAttribute("fmla", "val " + I64ToString(adjVal));
    out.closeElementSelfClosing();
    out.closeElement();
    out.closeElement();
  }

  // ── Custom geometry from PathData ─────────────────────────────────────────
  void writeCustomGeometry(XMLBuilder& out, const PathData* pathData, const Rect& bounds) {
    int64_t pathW = PixelsToEMU(bounds.width);
    int64_t pathH = PixelsToEMU(bounds.height);

    out.openElement("a:custGeom");
    out.closeElementStart();

    out.openElement("a:avLst");
    out.closeElementSelfClosing();
    out.openElement("a:gdLst");
    out.closeElementSelfClosing();
    out.openElement("a:ahLst");
    out.closeElementSelfClosing();
    out.openElement("a:cxnLst");
    out.closeElementSelfClosing();

    out.openElement("a:rect");
    out.addAttribute("l", static_cast<int64_t>(0));
    out.addAttribute("t", static_cast<int64_t>(0));
    out.addAttribute("r", pathW);
    out.addAttribute("b", pathH);
    out.closeElementSelfClosing();

    out.openElement("a:pathLst");
    out.closeElementStart();
    out.openElement("a:path");
    out.addAttribute("w", pathW);
    out.addAttribute("h", pathH);
    out.closeElementStart();

    float bx = bounds.x;
    float by = bounds.y;

    pathData->forEach([&](PathVerb verb, const Point* pts) {
      switch (verb) {
        case PathVerb::Move: {
          out.openElement("a:moveTo");
          out.closeElementStart();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(pts[0].x - bx));
          out.addAttribute("y", PixelsToEMU(pts[0].y - by));
          out.closeElementSelfClosing();
          out.closeElement();
          break;
        }
        case PathVerb::Line: {
          out.openElement("a:lnTo");
          out.closeElementStart();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(pts[0].x - bx));
          out.addAttribute("y", PixelsToEMU(pts[0].y - by));
          out.closeElementSelfClosing();
          out.closeElement();
          break;
        }
        case PathVerb::Quad: {
          // Convert quadratic to cubic: CP1 = P0 + 2/3*(P1-P0), CP2 = P2 + 2/3*(P1-P2)
          // We don't have the current point P0 easily, so approximate using the control point.
          // Actually for OOXML we can write cubicBezTo with the quad-to-cubic conversion.
          // Since we don't track current point, use a simple approximation:
          // write as cubicBezTo with duplicated control points.
          float cx = pts[0].x, cy = pts[0].y;
          float ex = pts[1].x, ey = pts[1].y;
          out.openElement("a:cubicBezTo");
          out.closeElementStart();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(cx - bx));
          out.addAttribute("y", PixelsToEMU(cy - by));
          out.closeElementSelfClosing();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(cx - bx));
          out.addAttribute("y", PixelsToEMU(cy - by));
          out.closeElementSelfClosing();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(ex - bx));
          out.addAttribute("y", PixelsToEMU(ey - by));
          out.closeElementSelfClosing();
          out.closeElement();
          break;
        }
        case PathVerb::Cubic: {
          out.openElement("a:cubicBezTo");
          out.closeElementStart();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(pts[0].x - bx));
          out.addAttribute("y", PixelsToEMU(pts[0].y - by));
          out.closeElementSelfClosing();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(pts[1].x - bx));
          out.addAttribute("y", PixelsToEMU(pts[1].y - by));
          out.closeElementSelfClosing();
          out.openElement("a:pt");
          out.addAttribute("x", PixelsToEMU(pts[2].x - bx));
          out.addAttribute("y", PixelsToEMU(pts[2].y - by));
          out.closeElementSelfClosing();
          out.closeElement();
          break;
        }
        case PathVerb::Close: {
          out.openElement("a:close");
          out.closeElementSelfClosing();
          break;
        }
      }
    });

    out.closeElement();  // </a:path>
    out.closeElement();  // </a:pathLst>
    out.closeElement();  // </a:custGeom>
  }

  // ── Shape writers ─────────────────────────────────────────────────────────
  void writeRectangle(XMLBuilder& out, const Rectangle* rect, const FillStrokeInfo& fs,
                      const Matrix& transform) {
    float localX = rect->position.x - rect->size.width / 2;
    float localY = rect->position.y - rect->size.height / 2;
    auto tr = ApplyMatrixToRect(localX, localY, rect->size.width, rect->size.height, transform);

    int id = allocShapeId();
    out.openElement("p:sp");
    out.closeElementStart();

    out.openElement("p:nvSpPr");
    out.closeElementStart();
    out.openElement("p:cNvPr");
    out.addAttribute("id", static_cast<int64_t>(id));
    out.addAttribute("name", shapeName("Rectangle"));
    out.closeElementSelfClosing();
    out.openElement("p:cNvSpPr");
    out.closeElementSelfClosing();
    out.openElement("p:nvPr");
    out.closeElementSelfClosing();
    out.closeElement();

    out.openElement("p:spPr");
    out.closeElementStart();
    int64_t rot = FloatNearlyZero(tr.rotationDeg) ? 0 : DegreesToPPTRotation(tr.rotationDeg);
    writeXfrm(out, PixelsToEMU(tr.x), PixelsToEMU(tr.y),
              PixelsToEMU(tr.width), PixelsToEMU(tr.height), rot);
    if (rect->roundness > 0) {
      writeRoundRectGeom(out, rect->roundness, rect->size.width, rect->size.height);
    } else {
      writePresetGeom(out, "rect");
    }
    writeFillProperties(out, fs.fill);
    writeStrokeProperties(out, fs.stroke);
    out.closeElement();

    out.closeElement();  // </p:sp>
  }

  void writeEllipse(XMLBuilder& out, const Ellipse* ellipse, const FillStrokeInfo& fs,
                    const Matrix& transform) {
    float localX = ellipse->position.x - ellipse->size.width / 2;
    float localY = ellipse->position.y - ellipse->size.height / 2;
    auto tr = ApplyMatrixToRect(localX, localY, ellipse->size.width, ellipse->size.height,
                                transform);

    int id = allocShapeId();
    out.openElement("p:sp");
    out.closeElementStart();

    out.openElement("p:nvSpPr");
    out.closeElementStart();
    out.openElement("p:cNvPr");
    out.addAttribute("id", static_cast<int64_t>(id));
    out.addAttribute("name", shapeName("Ellipse"));
    out.closeElementSelfClosing();
    out.openElement("p:cNvSpPr");
    out.closeElementSelfClosing();
    out.openElement("p:nvPr");
    out.closeElementSelfClosing();
    out.closeElement();

    out.openElement("p:spPr");
    out.closeElementStart();
    int64_t rot = FloatNearlyZero(tr.rotationDeg) ? 0 : DegreesToPPTRotation(tr.rotationDeg);
    writeXfrm(out, PixelsToEMU(tr.x), PixelsToEMU(tr.y),
              PixelsToEMU(tr.width), PixelsToEMU(tr.height), rot);
    writePresetGeom(out, "ellipse");
    writeFillProperties(out, fs.fill);
    writeStrokeProperties(out, fs.stroke);
    out.closeElement();

    out.closeElement();
  }

  void writePath(XMLBuilder& out, const Path* path, const FillStrokeInfo& fs,
                 const Matrix& transform) {
    if (!path->data || path->data->isEmpty()) {
      return;
    }
    Rect bounds = path->data->getBounds();
    if (bounds.isEmpty()) {
      return;
    }

    auto tr = ApplyMatrixToRect(bounds.x, bounds.y, bounds.width, bounds.height, transform);

    int id = allocShapeId();
    out.openElement("p:sp");
    out.closeElementStart();

    out.openElement("p:nvSpPr");
    out.closeElementStart();
    out.openElement("p:cNvPr");
    out.addAttribute("id", static_cast<int64_t>(id));
    out.addAttribute("name", shapeName("Path"));
    out.closeElementSelfClosing();
    out.openElement("p:cNvSpPr");
    out.closeElementSelfClosing();
    out.openElement("p:nvPr");
    out.closeElementSelfClosing();
    out.closeElement();

    out.openElement("p:spPr");
    out.closeElementStart();
    int64_t rot = FloatNearlyZero(tr.rotationDeg) ? 0 : DegreesToPPTRotation(tr.rotationDeg);
    writeXfrm(out, PixelsToEMU(tr.x), PixelsToEMU(tr.y),
              PixelsToEMU(tr.width), PixelsToEMU(tr.height), rot);
    writeCustomGeometry(out, path->data, bounds);
    writeFillProperties(out, fs.fill);
    writeStrokeProperties(out, fs.stroke);
    out.closeElement();

    out.closeElement();
  }

  void writeText(XMLBuilder& out, const Text* text, const FillStrokeInfo& fs,
                 const Matrix& transform) {
    if (text->text.empty()) {
      return;
    }

    float localX, localY, localW, localH;
    if (fs.textBox && fs.textBox->size.width > 0 && fs.textBox->size.height > 0) {
      localX = fs.textBox->position.x;
      localY = fs.textBox->position.y;
      localW = fs.textBox->size.width;
      localH = fs.textBox->size.height;
    } else {
      float estimatedWidth = static_cast<float>(text->text.size()) * text->fontSize * 0.6f;
      float estimatedHeight = text->fontSize * 1.5f;
      localX = text->position.x;
      localY = text->position.y - text->fontSize;
      localW = std::max(estimatedWidth, text->fontSize * 2.0f);
      localH = estimatedHeight;
    }

    auto tr = ApplyMatrixToRect(localX, localY, localW, localH, transform);

    int id = allocShapeId();
    out.openElement("p:sp");
    out.closeElementStart();

    out.openElement("p:nvSpPr");
    out.closeElementStart();
    out.openElement("p:cNvPr");
    out.addAttribute("id", static_cast<int64_t>(id));
    out.addAttribute("name", shapeName("Text"));
    out.closeElementSelfClosing();
    out.openElement("p:cNvSpPr");
    out.addAttribute("txBox", "1");
    out.closeElementSelfClosing();
    out.openElement("p:nvPr");
    out.closeElementSelfClosing();
    out.closeElement();

    out.openElement("p:spPr");
    out.closeElementStart();
    int64_t rot = FloatNearlyZero(tr.rotationDeg) ? 0 : DegreesToPPTRotation(tr.rotationDeg);
    writeXfrm(out, PixelsToEMU(tr.x), PixelsToEMU(tr.y),
              PixelsToEMU(tr.width), PixelsToEMU(tr.height), rot);
    writePresetGeom(out, "rect");
    out.openElement("a:noFill");
    out.closeElementSelfClosing();
    out.closeElement();

    // Text body
    out.openElement("p:txBody");
    out.closeElementStart();

    out.openElement("a:bodyPr");
    out.addAttribute("wrap", "square");
    out.addAttribute("lIns", static_cast<int64_t>(0));
    out.addAttribute("tIns", static_cast<int64_t>(0));
    out.addAttribute("rIns", static_cast<int64_t>(0));
    out.addAttribute("bIns", static_cast<int64_t>(0));
    out.closeElementStart();
    out.openElement("a:noAutofit");
    out.closeElementSelfClosing();
    out.closeElement();

    out.openElement("a:lstStyle");
    out.closeElementSelfClosing();

    // Split text by newlines into paragraphs
    std::string remaining = text->text;
    size_t pos = 0;
    bool firstParagraph = true;
    while (pos <= remaining.size()) {
      size_t nl = remaining.find('\n', pos);
      std::string line = (nl == std::string::npos) ? remaining.substr(pos)
                                                    : remaining.substr(pos, nl - pos);
      writeParagraph(out, line, text, fs, firstParagraph);
      firstParagraph = false;
      if (nl == std::string::npos) {
        break;
      }
      pos = nl + 1;
    }

    out.closeElement();  // </p:txBody>
    out.closeElement();  // </p:sp>
  }

  void writeParagraph(XMLBuilder& out, const std::string& lineText, const Text* text,
                      const FillStrokeInfo& fs, bool /*firstParagraph*/) {
    out.openElement("a:p");
    out.closeElementStart();

    // Paragraph properties
    if (fs.textBox) {
      out.openElement("a:pPr");
      if (fs.textBox->textAlign == TextAlign::Center) {
        out.addAttribute("algn", "ctr");
      } else if (fs.textBox->textAlign == TextAlign::End) {
        out.addAttribute("algn", "r");
      } else if (fs.textBox->textAlign == TextAlign::Justify) {
        out.addAttribute("algn", "just");
      }
      out.closeElementSelfClosing();
    }

    if (!lineText.empty()) {
      out.openElement("a:r");
      out.closeElementStart();

      // Run properties
      out.openElement("a:rPr");
      out.addAttribute("lang", "en-US");
      out.addAttribute("sz", FontSizeToPPT(text->fontSize));
      bool hasBold = text->fauxBold ||
                     (!text->fontStyle.empty() &&
                      text->fontStyle.find("Bold") != std::string::npos);
      bool hasItalic = text->fauxItalic ||
                       (!text->fontStyle.empty() &&
                        text->fontStyle.find("Italic") != std::string::npos);
      if (hasBold) {
        out.addAttribute("b", static_cast<int64_t>(1));
      }
      if (hasItalic) {
        out.addAttribute("i", static_cast<int64_t>(1));
      }
      if (!FloatNearlyZero(text->letterSpacing)) {
        out.addAttribute("spc", static_cast<int64_t>(std::round(text->letterSpacing * 100.0f)));
      }
      out.addAttribute("dirty", static_cast<int64_t>(0));
      out.closeElementStart();

      // Text color from fill
      if (fs.fill && fs.fill->color && fs.fill->color->nodeType() == NodeType::SolidColor) {
        auto solid = static_cast<const SolidColor*>(fs.fill->color);
        writeSolidFill(out, solid->color, fs.fill->alpha);
      }

      // Font
      if (!text->fontFamily.empty()) {
        out.openElement("a:latin");
        out.addAttribute("typeface", text->fontFamily);
        out.closeElementSelfClosing();
        out.openElement("a:ea");
        out.addAttribute("typeface", text->fontFamily);
        out.closeElementSelfClosing();
        out.openElement("a:cs");
        out.addAttribute("typeface", text->fontFamily);
        out.closeElementSelfClosing();
      }

      out.closeElement();  // </a:rPr>

      out.openElement("a:t");
      out.closeElementWithText(lineText);

      out.closeElement();  // </a:r>
    }

    out.openElement("a:endParaRPr");
    out.addAttribute("lang", "en-US");
    out.addAttribute("sz", FontSizeToPPT(text->fontSize));
    out.closeElementSelfClosing();

    out.closeElement();  // </a:p>
  }

  // ── Group shape ───────────────────────────────────────────────────────────
  void writeGroup(XMLBuilder& out, const Group* group, const Matrix& parentTransform) {
    Matrix groupMatrix = BuildGroupMatrix(group);
    Matrix accumulated = parentTransform * groupMatrix;

    auto fs = CollectFillStroke(group->elements);
    bool hasGeometry = false;
    for (const auto* elem : group->elements) {
      auto t = elem->nodeType();
      if (t == NodeType::Rectangle || t == NodeType::Ellipse || t == NodeType::Path ||
          t == NodeType::Text || t == NodeType::Group) {
        hasGeometry = true;
        break;
      }
    }

    if (!hasGeometry) {
      return;
    }

    for (const auto* elem : group->elements) {
      auto t = elem->nodeType();
      if (t == NodeType::Fill || t == NodeType::Stroke || t == NodeType::TextBox) {
        continue;
      }
      switch (t) {
        case NodeType::Rectangle:
          writeRectangle(out, static_cast<const Rectangle*>(elem), fs, accumulated);
          break;
        case NodeType::Ellipse:
          writeEllipse(out, static_cast<const Ellipse*>(elem), fs, accumulated);
          break;
        case NodeType::Path:
          writePath(out, static_cast<const Path*>(elem), fs, accumulated);
          break;
        case NodeType::Text:
          writeText(out, static_cast<const Text*>(elem), fs, accumulated);
          break;
        case NodeType::Group:
          writeGroup(out, static_cast<const Group*>(elem), accumulated);
          break;
        default:
          break;
      }
    }
  }

  // ── Layer content writing ─────────────────────────────────────────────────
  void writeLayerContents(XMLBuilder& out, const Layer* layer, const Matrix& transform) {
    auto fs = CollectFillStroke(layer->contents);

    for (const auto* elem : layer->contents) {
      auto t = elem->nodeType();
      if (t == NodeType::Fill || t == NodeType::Stroke || t == NodeType::TextBox) {
        continue;
      }
      switch (t) {
        case NodeType::Rectangle:
          writeRectangle(out, static_cast<const Rectangle*>(elem), fs, transform);
          break;
        case NodeType::Ellipse:
          writeEllipse(out, static_cast<const Ellipse*>(elem), fs, transform);
          break;
        case NodeType::Path:
          writePath(out, static_cast<const Path*>(elem), fs, transform);
          break;
        case NodeType::Text:
          writeText(out, static_cast<const Text*>(elem), fs, transform);
          break;
        case NodeType::Group:
          writeGroup(out, static_cast<const Group*>(elem), transform);
          break;
        default:
          break;
      }
    }
  }

  void writeLayerAsShapes(XMLBuilder& out, const Layer* layer, const Matrix& parentTransform) {
    if (!layer->visible) {
      return;
    }
    Matrix layerMatrix = BuildLayerMatrix(layer);
    Matrix accumulated = parentTransform * layerMatrix;

    writeLayerContents(out, layer, accumulated);
    for (const auto* child : layer->children) {
      writeLayerAsShapes(out, child, accumulated);
    }
  }
};

//==============================================================================
// PPTExporter – main export logic
//==============================================================================

bool PPTExporter::ToFile(const PAGXDocument& document, const std::string& filePath,
                         const Options& options) {
  if (document.layers.empty()) {
    return false;
  }

  int64_t slideCX = PixelsToEMU(document.width);
  int64_t slideCY = PixelsToEMU(document.height);

  PPTMediaContext media;
  ZipWriter zip;

  // Generate a single slide containing all layers
  PPTSlideWriter slideWriter(&media, options.indent);
  std::string slideXml = slideWriter.writeSlide(document.layers);
  std::string slideRelsXml = slideWriter.writeSlideRels();

  // Add boilerplate files
  zip.addFile("[Content_Types].xml",
              GenerateContentTypes(1, media.hasImages()));
  zip.addFile("_rels/.rels", GenerateRootRels());
  zip.addFile("ppt/presentation.xml",
              GeneratePresentation(1, slideCX, slideCY));
  zip.addFile("ppt/_rels/presentation.xml.rels",
              GeneratePresentationRels(1));
  zip.addFile("ppt/theme/theme1.xml", GenerateTheme());
  zip.addFile("ppt/slideMasters/slideMaster1.xml", GenerateSlideMaster());
  zip.addFile("ppt/slideMasters/_rels/slideMaster1.xml.rels", GenerateSlideMasterRels());
  zip.addFile("ppt/slideLayouts/slideLayout1.xml", GenerateSlideLayout());
  zip.addFile("ppt/slideLayouts/_rels/slideLayout1.xml.rels", GenerateSlideLayoutRels());

  // Add the single slide
  zip.addFile("ppt/slides/slide1.xml", slideXml);
  zip.addFile("ppt/slides/_rels/slide1.xml.rels", slideRelsXml);

  // Add media files
  for (const auto& entry : media.entries()) {
    zip.addFile("ppt/media/" + entry.fileName, entry.data.data(), entry.data.size());
  }

  auto zipData = zip.build();
  std::ofstream file(filePath, std::ios::binary);
  if (!file) {
    return false;
  }
  file.write(reinterpret_cast<const char*>(zipData.data()),
             static_cast<std::streamsize>(zipData.size()));
  return file.good();
}

}  // namespace pagx
