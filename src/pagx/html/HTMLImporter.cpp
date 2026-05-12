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

#include "pagx/HTMLImporter.h"
#include <cctype>
#include <cmath>
#include <cstdlib>
#include <cstring>
#include <sstream>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include "pagx/PAGXDocument.h"
#include "pagx/nodes/DropShadowStyle.h"
#include "pagx/nodes/Fill.h"
#include "pagx/nodes/Image.h"
#include "pagx/nodes/ImagePattern.h"
#include "pagx/nodes/Layer.h"
#include "pagx/nodes/Rectangle.h"
#include "pagx/nodes/RoundCorner.h"
#include "pagx/nodes/SolidColor.h"
#include "pagx/nodes/Stroke.h"
#include "pagx/nodes/Text.h"
#include "pagx/nodes/TextBox.h"
#include "pagx/utils/StringParser.h"
#include "pagx/xml/XMLDOM.h"

namespace pagx {

namespace {

constexpr int MAX_HTML_RECURSION_DEPTH = 128;
constexpr float DEFAULT_FONT_SIZE = 16.0f;

//==============================================================================
// String helpers
//==============================================================================

std::string ToLower(std::string s) {
  for (auto& c : s) {
    c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
  }
  return s;
}

std::string Trim(const std::string& s) {
  size_t start = 0;
  while (start < s.size() && std::isspace(static_cast<unsigned char>(s[start]))) {
    start++;
  }
  size_t end = s.size();
  while (end > start && std::isspace(static_cast<unsigned char>(s[end - 1]))) {
    end--;
  }
  return s.substr(start, end - start);
}

//==============================================================================
// CSS color parsing
//==============================================================================

const std::unordered_map<std::string, uint32_t>& CSSNamedColors() {
  // Limited CSS Color 3 set: 16 base + extended common values.
  // Hex is 0xRRGGBB.
  static const std::unordered_map<std::string, uint32_t> table = {
      {"aliceblue", 0xF0F8FF},
      {"antiquewhite", 0xFAEBD7},
      {"aqua", 0x00FFFF},
      {"aquamarine", 0x7FFFD4},
      {"azure", 0xF0FFFF},
      {"beige", 0xF5F5DC},
      {"bisque", 0xFFE4C4},
      {"black", 0x000000},
      {"blanchedalmond", 0xFFEBCD},
      {"blue", 0x0000FF},
      {"blueviolet", 0x8A2BE2},
      {"brown", 0xA52A2A},
      {"burlywood", 0xDEB887},
      {"cadetblue", 0x5F9EA0},
      {"chartreuse", 0x7FFF00},
      {"chocolate", 0xD2691E},
      {"coral", 0xFF7F50},
      {"cornflowerblue", 0x6495ED},
      {"cornsilk", 0xFFF8DC},
      {"crimson", 0xDC143C},
      {"cyan", 0x00FFFF},
      {"darkblue", 0x00008B},
      {"darkcyan", 0x008B8B},
      {"darkgoldenrod", 0xB8860B},
      {"darkgray", 0xA9A9A9},
      {"darkgreen", 0x006400},
      {"darkgrey", 0xA9A9A9},
      {"darkkhaki", 0xBDB76B},
      {"darkmagenta", 0x8B008B},
      {"darkolivegreen", 0x556B2F},
      {"darkorange", 0xFF8C00},
      {"darkorchid", 0x9932CC},
      {"darkred", 0x8B0000},
      {"darksalmon", 0xE9967A},
      {"darkseagreen", 0x8FBC8F},
      {"darkslateblue", 0x483D8B},
      {"darkslategray", 0x2F4F4F},
      {"darkslategrey", 0x2F4F4F},
      {"darkturquoise", 0x00CED1},
      {"darkviolet", 0x9400D3},
      {"deeppink", 0xFF1493},
      {"deepskyblue", 0x00BFFF},
      {"dimgray", 0x696969},
      {"dimgrey", 0x696969},
      {"dodgerblue", 0x1E90FF},
      {"firebrick", 0xB22222},
      {"floralwhite", 0xFFFAF0},
      {"forestgreen", 0x228B22},
      {"fuchsia", 0xFF00FF},
      {"gainsboro", 0xDCDCDC},
      {"ghostwhite", 0xF8F8FF},
      {"gold", 0xFFD700},
      {"goldenrod", 0xDAA520},
      {"gray", 0x808080},
      {"green", 0x008000},
      {"greenyellow", 0xADFF2F},
      {"grey", 0x808080},
      {"honeydew", 0xF0FFF0},
      {"hotpink", 0xFF69B4},
      {"indianred", 0xCD5C5C},
      {"indigo", 0x4B0082},
      {"ivory", 0xFFFFF0},
      {"khaki", 0xF0E68C},
      {"lavender", 0xE6E6FA},
      {"lavenderblush", 0xFFF0F5},
      {"lawngreen", 0x7CFC00},
      {"lemonchiffon", 0xFFFACD},
      {"lightblue", 0xADD8E6},
      {"lightcoral", 0xF08080},
      {"lightcyan", 0xE0FFFF},
      {"lightgoldenrodyellow", 0xFAFAD2},
      {"lightgray", 0xD3D3D3},
      {"lightgreen", 0x90EE90},
      {"lightgrey", 0xD3D3D3},
      {"lightpink", 0xFFB6C1},
      {"lightsalmon", 0xFFA07A},
      {"lightseagreen", 0x20B2AA},
      {"lightskyblue", 0x87CEFA},
      {"lightslategray", 0x778899},
      {"lightslategrey", 0x778899},
      {"lightsteelblue", 0xB0C4DE},
      {"lightyellow", 0xFFFFE0},
      {"lime", 0x00FF00},
      {"limegreen", 0x32CD32},
      {"linen", 0xFAF0E6},
      {"magenta", 0xFF00FF},
      {"maroon", 0x800000},
      {"mediumaquamarine", 0x66CDAA},
      {"mediumblue", 0x0000CD},
      {"mediumorchid", 0xBA55D3},
      {"mediumpurple", 0x9370DB},
      {"mediumseagreen", 0x3CB371},
      {"mediumslateblue", 0x7B68EE},
      {"mediumspringgreen", 0x00FA9A},
      {"mediumturquoise", 0x48D1CC},
      {"mediumvioletred", 0xC71585},
      {"midnightblue", 0x191970},
      {"mintcream", 0xF5FFFA},
      {"mistyrose", 0xFFE4E1},
      {"moccasin", 0xFFE4B5},
      {"navajowhite", 0xFFDEAD},
      {"navy", 0x000080},
      {"oldlace", 0xFDF5E6},
      {"olive", 0x808000},
      {"olivedrab", 0x6B8E23},
      {"orange", 0xFFA500},
      {"orangered", 0xFF4500},
      {"orchid", 0xDA70D6},
      {"palegoldenrod", 0xEEE8AA},
      {"palegreen", 0x98FB98},
      {"paleturquoise", 0xAFEEEE},
      {"palevioletred", 0xDB7093},
      {"papayawhip", 0xFFEFD5},
      {"peachpuff", 0xFFDAB9},
      {"peru", 0xCD853F},
      {"pink", 0xFFC0CB},
      {"plum", 0xDDA0DD},
      {"powderblue", 0xB0E0E6},
      {"purple", 0x800080},
      {"rebeccapurple", 0x663399},
      {"red", 0xFF0000},
      {"rosybrown", 0xBC8F8F},
      {"royalblue", 0x4169E1},
      {"saddlebrown", 0x8B4513},
      {"salmon", 0xFA8072},
      {"sandybrown", 0xF4A460},
      {"seagreen", 0x2E8B57},
      {"seashell", 0xFFF5EE},
      {"sienna", 0xA0522D},
      {"silver", 0xC0C0C0},
      {"skyblue", 0x87CEEB},
      {"slateblue", 0x6A5ACD},
      {"slategray", 0x708090},
      {"slategrey", 0x708090},
      {"snow", 0xFFFAFA},
      {"springgreen", 0x00FF7F},
      {"steelblue", 0x4682B4},
      {"tan", 0xD2B48C},
      {"teal", 0x008080},
      {"thistle", 0xD8BFD8},
      {"tomato", 0xFF6347},
      {"transparent", 0x000000},
      {"turquoise", 0x40E0D0},
      {"violet", 0xEE82EE},
      {"wheat", 0xF5DEB3},
      {"white", 0xFFFFFF},
      {"whitesmoke", 0xF5F5F5},
      {"yellow", 0xFFFF00},
      {"yellowgreen", 0x9ACD32},
  };
  return table;
}

bool ParseHexDigit(char c, int* outValue) {
  if (c >= '0' && c <= '9') {
    *outValue = c - '0';
    return true;
  }
  if (c >= 'a' && c <= 'f') {
    *outValue = 10 + (c - 'a');
    return true;
  }
  if (c >= 'A' && c <= 'F') {
    *outValue = 10 + (c - 'A');
    return true;
  }
  return false;
}

bool ParseHexColor(const std::string& value, Color* outColor) {
  if (value.empty() || value[0] != '#') {
    return false;
  }
  size_t len = value.size();
  if (len != 4 && len != 5 && len != 7 && len != 9) {
    return false;
  }
  int digits[8] = {0};
  size_t hexLen = len - 1;
  for (size_t i = 0; i < hexLen; i++) {
    if (!ParseHexDigit(value[i + 1], &digits[i])) {
      return false;
    }
  }
  Color color = {};
  color.colorSpace = ColorSpace::SRGB;
  if (hexLen == 3 || hexLen == 4) {
    color.red = static_cast<float>(digits[0] * 17) / 255.0f;
    color.green = static_cast<float>(digits[1] * 17) / 255.0f;
    color.blue = static_cast<float>(digits[2] * 17) / 255.0f;
    color.alpha = (hexLen == 4) ? static_cast<float>(digits[3] * 17) / 255.0f : 1.0f;
  } else {
    color.red = static_cast<float>(digits[0] * 16 + digits[1]) / 255.0f;
    color.green = static_cast<float>(digits[2] * 16 + digits[3]) / 255.0f;
    color.blue = static_cast<float>(digits[4] * 16 + digits[5]) / 255.0f;
    color.alpha = (hexLen == 8) ? static_cast<float>(digits[6] * 16 + digits[7]) / 255.0f : 1.0f;
  }
  *outColor = color;
  return true;
}

bool ParseRgbFunction(const std::string& value, Color* outColor) {
  std::string lower = ToLower(value);
  bool isRgba = lower.compare(0, 5, "rgba(") == 0;
  bool isRgb = lower.compare(0, 4, "rgb(") == 0;
  if (!isRgb && !isRgba) {
    return false;
  }
  size_t lp = value.find('(');
  size_t rp = value.find(')');
  if (lp == std::string::npos || rp == std::string::npos || rp <= lp + 1) {
    return false;
  }
  std::string inner = value.substr(lp + 1, rp - lp - 1);
  // Accept commas or whitespace as separators; alpha may be after a slash.
  // Split on ',', '/', and whitespace.
  std::vector<std::string> parts;
  {
    std::string cur;
    for (char c : inner) {
      if (c == ',' || c == '/' || std::isspace(static_cast<unsigned char>(c))) {
        if (!cur.empty()) {
          parts.push_back(cur);
          cur.clear();
        }
      } else {
        cur.push_back(c);
      }
    }
    if (!cur.empty()) {
      parts.push_back(cur);
    }
  }
  if (parts.size() < 3) {
    return false;
  }
  auto parseChannel = [](const std::string& token) -> float {
    if (token.empty()) {
      return 0.0f;
    }
    if (token.back() == '%') {
      float v = static_cast<float>(std::strtod(token.c_str(), nullptr));
      return std::max(0.0f, std::min(1.0f, v / 100.0f));
    }
    float v = static_cast<float>(std::strtod(token.c_str(), nullptr));
    return std::max(0.0f, std::min(1.0f, v / 255.0f));
  };
  auto parseAlpha = [](const std::string& token) -> float {
    if (token.empty()) {
      return 1.0f;
    }
    if (token.back() == '%') {
      float v = static_cast<float>(std::strtod(token.c_str(), nullptr));
      return std::max(0.0f, std::min(1.0f, v / 100.0f));
    }
    float v = static_cast<float>(std::strtod(token.c_str(), nullptr));
    return std::max(0.0f, std::min(1.0f, v));
  };
  Color color = {};
  color.colorSpace = ColorSpace::SRGB;
  color.red = parseChannel(parts[0]);
  color.green = parseChannel(parts[1]);
  color.blue = parseChannel(parts[2]);
  color.alpha = parts.size() >= 4 ? parseAlpha(parts[3]) : 1.0f;
  *outColor = color;
  return true;
}

bool ParseCSSColor(const std::string& rawValue, Color* outColor) {
  std::string value = Trim(rawValue);
  if (value.empty() || ToLower(value) == "none" || ToLower(value) == "currentcolor") {
    return false;
  }
  if (value[0] == '#') {
    return ParseHexColor(value, outColor);
  }
  std::string lower = ToLower(value);
  if (lower.compare(0, 3, "rgb") == 0) {
    return ParseRgbFunction(value, outColor);
  }
  const auto& table = CSSNamedColors();
  auto it = table.find(lower);
  if (it != table.end()) {
    uint32_t hex = it->second;
    Color color = {};
    color.colorSpace = ColorSpace::SRGB;
    color.red = static_cast<float>((hex >> 16) & 0xFF) / 255.0f;
    color.green = static_cast<float>((hex >> 8) & 0xFF) / 255.0f;
    color.blue = static_cast<float>(hex & 0xFF) / 255.0f;
    color.alpha = (lower == "transparent") ? 0.0f : 1.0f;
    *outColor = color;
    return true;
  }
  return false;
}

//==============================================================================
// CSS length parsing
//
// Recognized: bare number (treated as px), `Npx`, `N%`. Other units (em/rem/vh/vw)
// produce a parse failure so the caller can emit a warning.
//==============================================================================

enum class LengthUnit { Px, Percent, Unsupported };

bool ParseLength(const std::string& rawValue, float* outValue, LengthUnit* outUnit) {
  std::string value = Trim(rawValue);
  if (value.empty()) {
    return false;
  }
  size_t i = 0;
  while (i < value.size() && (std::isdigit(static_cast<unsigned char>(value[i])) ||
                              value[i] == '.' || value[i] == '-' || value[i] == '+')) {
    i++;
  }
  if (i == 0) {
    return false;
  }
  float number = static_cast<float>(std::strtod(value.c_str(), nullptr));
  std::string unit = ToLower(Trim(value.substr(i)));
  if (unit.empty() || unit == "px") {
    *outValue = number;
    *outUnit = LengthUnit::Px;
    return true;
  }
  if (unit == "%") {
    *outValue = number;
    *outUnit = LengthUnit::Percent;
    return true;
  }
  *outValue = number;
  *outUnit = LengthUnit::Unsupported;
  return true;  // Return true so the caller can decide how to warn.
}

//==============================================================================
// Style declaration parser: parses `style="k: v; k: v"` into a map.
// Property names are lowercased; values are trimmed.
//==============================================================================

std::unordered_map<std::string, std::string> ParseStyleDecl(const std::string& style) {
  std::unordered_map<std::string, std::string> out;
  size_t pos = 0;
  while (pos < style.size()) {
    size_t end = style.find(';', pos);
    if (end == std::string::npos) {
      end = style.size();
    }
    std::string decl = style.substr(pos, end - pos);
    pos = end + 1;
    size_t colon = decl.find(':');
    if (colon == std::string::npos) {
      continue;
    }
    std::string key = ToLower(Trim(decl.substr(0, colon)));
    std::string value = Trim(decl.substr(colon + 1));
    if (!key.empty()) {
      out[key] = value;
    }
  }
  return out;
}

//==============================================================================
// HTML element category helpers
//==============================================================================

bool IsHtmlContainerTag(const std::string& tag) {
  static const std::unordered_set<std::string> set = {
      "div", "section", "article", "header", "footer", "nav",   "main",   "aside",     "button",
      "a",   "ul",      "ol",      "li",     "form",   "label", "figure", "figcaption"};
  return set.count(tag) > 0;
}

bool IsHtmlTextTag(const std::string& tag) {
  // Block-level text containers we treat as Layer-with-TextBox.
  static const std::unordered_set<std::string> set = {"p",  "h1", "h2",         "h3",  "h4",
                                                      "h5", "h6", "blockquote", "pre", "code"};
  return set.count(tag) > 0;
}

bool IsHtmlInlineTag(const std::string& tag) {
  // Inline tags whose text content is flattened into the surrounding block (v1 behaviour:
  // formatting like <b>/<i> is dropped, only the text remains).
  static const std::unordered_set<std::string> set = {"span", "strong", "b",     "em",
                                                      "i",    "u",      "small", "mark"};
  return set.count(tag) > 0;
}

float HeadingFontSizePx(const std::string& tag) {
  // Browser defaults (approximate) so headings have visible hierarchy without explicit styling.
  if (tag == "h1") return 32.0f;
  if (tag == "h2") return 24.0f;
  if (tag == "h3") return 18.72f;
  if (tag == "h4") return 16.0f;
  if (tag == "h5") return 13.28f;
  if (tag == "h6") return 10.72f;
  return DEFAULT_FONT_SIZE;
}

//==============================================================================
// Inherited text styling that cascades down the element tree.
//==============================================================================

struct TextStyle {
  Color color = {0, 0, 0, 1, ColorSpace::SRGB};
  std::string fontFamily = {};
  float fontSize = DEFAULT_FONT_SIZE;
  bool bold = false;
  bool italic = false;
  float letterSpacing = 0.0f;
  float lineHeight = 0.0f;  // 0 = auto from font metrics
  TextAlign textAlign = TextAlign::Start;
};

//==============================================================================
// Parser context
//==============================================================================

class HTMLParserContext {
 public:
  explicit HTMLParserContext(const HTMLImporter::Options& options) : _options(options) {
  }

  std::shared_ptr<PAGXDocument> parse(const uint8_t* data, size_t length);
  std::shared_ptr<PAGXDocument> parseFile(const std::string& filePath);

 private:
  std::shared_ptr<PAGXDocument> parseDOM(const std::shared_ptr<XMLDOM>& dom);

  // Resolves the canvas size from <html>/options. Always returns a positive size.
  void resolveCanvasSize(const std::shared_ptr<DOMNode>& root, float* outWidth, float* outHeight);

  // Walks the body subtree and converts elements into a layer tree.
  Layer* convertElement(const std::shared_ptr<DOMNode>& element, const TextStyle& parentTextStyle,
                        int depth);
  void convertChildren(const std::shared_ptr<DOMNode>& element, Layer* parent,
                       const TextStyle& inheritedTextStyle, int depth);

  // Style application. Returns the updated text style for child elements.
  TextStyle applyStyle(Layer* layer, const std::string& tag,
                       const std::unordered_map<std::string, std::string>& props,
                       const TextStyle& parentTextStyle);

  // Builds a leaf Layer holding a single TextBox + Text from the given text.
  Layer* makeTextLayer(const std::string& text, const TextStyle& style);

  // Inline children/text: returns concatenated plain text after stripping inline tags.
  std::string collectInlineText(const std::shared_ptr<DOMNode>& element);

  // Returns true if all element children of `element` are inline tags or whitespace text.
  bool isLeafTextContainer(const std::shared_ptr<DOMNode>& element);

  // Converts an <img> element to a Layer carrying an Image content node.
  Layer* convertImage(const std::shared_ptr<DOMNode>& element,
                      const std::unordered_map<std::string, std::string>& props);

  // Converts an inline <svg> subtree to a Layer with an unresolved import directive.
  Layer* convertInlineSvg(const std::shared_ptr<DOMNode>& element,
                          const std::unordered_map<std::string, std::string>& props);

  // Helpers.
  std::string getAttr(const std::shared_ptr<DOMNode>& node, const std::string& name);
  std::unordered_map<std::string, std::string> readStyle(const std::shared_ptr<DOMNode>& node);
  void serializeNode(const std::shared_ptr<DOMNode>& node, std::string& out);
  void warn(const std::string& message);

  // After a Layer is fully populated, splits it into outer (background) + inner (content) when
  // both a background painter and structural attributes (padding/gap/children) coexist on the
  // same Layer. The PAGX layout engine otherwise insets the background by the padding.
  // Returns the input layer (possibly with structure rewritten) to keep the caller-side identity
  // stable for the parent's children list.
  void splitBackgroundIfNeeded(Layer* layer);

  HTMLImporter::Options _options;
  std::shared_ptr<PAGXDocument> _document;
};

std::string HTMLParserContext::getAttr(const std::shared_ptr<DOMNode>& node,
                                       const std::string& name) {
  if (!node) {
    return {};
  }
  auto* value = node->findAttribute(name);
  return value ? *value : std::string{};
}

std::unordered_map<std::string, std::string> HTMLParserContext::readStyle(
    const std::shared_ptr<DOMNode>& node) {
  return ParseStyleDecl(getAttr(node, "style"));
}

void HTMLParserContext::warn(const std::string& message) {
  if (_document) {
    _document->errors.push_back(message);
  }
}

//==============================================================================
// Top-level entry
//==============================================================================

std::shared_ptr<PAGXDocument> HTMLParserContext::parse(const uint8_t* data, size_t length) {
  if (!data || length == 0) {
    return nullptr;
  }
  auto dom = XMLDOM::Make(data, length);
  if (!dom) {
    return nullptr;
  }
  return parseDOM(dom);
}

std::shared_ptr<PAGXDocument> HTMLParserContext::parseFile(const std::string& filePath) {
  auto dom = XMLDOM::MakeFromFile(filePath);
  if (!dom) {
    return nullptr;
  }
  return parseDOM(dom);
}

void HTMLParserContext::resolveCanvasSize(const std::shared_ptr<DOMNode>& root, float* outWidth,
                                          float* outHeight) {
  bool hasExternal = !std::isnan(_options.targetWidth) && !std::isnan(_options.targetHeight);
  if (hasExternal && _options.targetWidth > 0 && _options.targetHeight > 0) {
    *outWidth = _options.targetWidth;
    *outHeight = _options.targetHeight;
    return;
  }
  float width = 0;
  float height = 0;
  if (root) {
    auto props = readStyle(root);
    auto wIt = props.find("width");
    auto hIt = props.find("height");
    LengthUnit unit;
    if (wIt != props.end() && ParseLength(wIt->second, &width, &unit)) {
      if (unit != LengthUnit::Px) width = 0;
    }
    if (hIt != props.end() && ParseLength(hIt->second, &height, &unit)) {
      if (unit != LengthUnit::Px) height = 0;
    }
  }
  if (width <= 0) width = _options.defaultWidth;
  if (height <= 0) height = _options.defaultHeight;
  *outWidth = width;
  *outHeight = height;
}

std::shared_ptr<PAGXDocument> HTMLParserContext::parseDOM(const std::shared_ptr<XMLDOM>& dom) {
  auto root = dom->getRootNode();
  if (!root) {
    return nullptr;
  }
  std::string rootName = ToLower(root->name);

  // Tolerate both <html> and a bare <body>/<div> root for fragments.
  std::shared_ptr<DOMNode> htmlNode;
  std::shared_ptr<DOMNode> bodyNode;
  if (rootName == "html") {
    htmlNode = root;
    auto child = root->firstChild;
    while (child) {
      if (child->type == DOMNodeType::Element && ToLower(child->name) == "body") {
        bodyNode = child;
        break;
      }
      child = child->nextSibling;
    }
    if (!bodyNode) {
      // No <body> — treat <html> children directly as body content.
      bodyNode = root;
    }
  } else if (rootName == "body") {
    bodyNode = root;
  } else {
    // Fragment: wrap in a synthetic body. We simply use the root as the body container.
    bodyNode = root;
  }

  float width = 0;
  float height = 0;
  resolveCanvasSize(htmlNode, &width, &height);
  if (width <= 0 || height <= 0) {
    return nullptr;
  }
  _document = PAGXDocument::Make(width, height);

  // Top-level body becomes a single Layer that fills the canvas. Both <html> and <body> styles
  // cascade onto this single layer; the <html> background colors the canvas, while <body>'s
  // padding/font-* defaults apply to all content.
  auto* bodyLayer = _document->makeNode<Layer>("body");
  bodyLayer->name = "body";
  bodyLayer->left = 0;
  bodyLayer->top = 0;
  bodyLayer->right = 0;
  bodyLayer->bottom = 0;
  bodyLayer->layout = LayoutMode::Vertical;
  bodyLayer->alignment = Alignment::Stretch;
  bodyLayer->arrangement = Arrangement::Start;

  TextStyle initialStyle = {};
  if (htmlNode && htmlNode != bodyNode) {
    auto htmlProps = readStyle(htmlNode);
    // Width/height on <html> are already consumed as the canvas size — applying them again on
    // the body layer would conflict with its left/right/top/bottom = 0 stretch.
    htmlProps.erase("width");
    htmlProps.erase("height");
    initialStyle = applyStyle(bodyLayer, "html", htmlProps, initialStyle);
  }
  if (bodyNode) {
    auto bodyProps = readStyle(bodyNode);
    initialStyle = applyStyle(bodyLayer, "body", bodyProps, initialStyle);
  }

  _document->layers.push_back(bodyLayer);

  if (bodyNode) {
    convertChildren(bodyNode, bodyLayer, initialStyle, 0);
  }
  splitBackgroundIfNeeded(bodyLayer);
  return _document;
}

//==============================================================================
// Element conversion
//==============================================================================

bool HTMLParserContext::isLeafTextContainer(const std::shared_ptr<DOMNode>& element) {
  bool hasText = false;
  auto child = element->firstChild;
  while (child) {
    if (child->type == DOMNodeType::Text) {
      if (!Trim(child->name).empty()) {
        hasText = true;
      }
    } else if (child->type == DOMNodeType::Element) {
      std::string tag = ToLower(child->name);
      if (tag == "br") {
        // newline marker, treated as text.
      } else if (IsHtmlInlineTag(tag)) {
        hasText = true;
      } else {
        return false;  // Has a block child — not a leaf text container.
      }
    }
    child = child->nextSibling;
  }
  return hasText;
}

// `<br/>` markers must survive whitespace collapsing. Use ASCII unit separator (US, 0x1F)
// internally as a sentinel that the final pass converts back to `\n`.
constexpr char kBrSentinel = '\x1F';

std::string HTMLParserContext::collectInlineText(const std::shared_ptr<DOMNode>& element) {
  std::string out;
  auto child = element->firstChild;
  while (child) {
    if (child->type == DOMNodeType::Text) {
      // CSS `white-space: normal`: every run of whitespace (including text-node newlines)
      // collapses to a single space.
      bool lastWasSpace = !out.empty() && (out.back() == ' ' || out.back() == kBrSentinel);
      for (char c : child->name) {
        if (std::isspace(static_cast<unsigned char>(c))) {
          if (!lastWasSpace) {
            out.push_back(' ');
            lastWasSpace = true;
          }
        } else {
          out.push_back(c);
          lastWasSpace = false;
        }
      }
    } else if (child->type == DOMNodeType::Element) {
      std::string tag = ToLower(child->name);
      if (tag == "br") {
        out.push_back(kBrSentinel);
      } else if (IsHtmlInlineTag(tag)) {
        out += collectInlineText(child);
      }
    }
    child = child->nextSibling;
  }
  std::string trimmed = Trim(out);
  std::string final;
  final.reserve(trimmed.size());
  for (char c : trimmed) {
    final.push_back(c == kBrSentinel ? '\n' : c);
  }
  return final;
}

void HTMLParserContext::convertChildren(const std::shared_ptr<DOMNode>& element, Layer* parent,
                                        const TextStyle& inheritedTextStyle, int depth) {
  if (depth >= MAX_HTML_RECURSION_DEPTH) {
    warn("Maximum HTML recursion depth exceeded; deeper content was dropped.");
    return;
  }
  // Buffer for accumulated inline text/whitespace between block children.
  std::string textBuffer;
  bool textHasContent = false;
  auto flushText = [&]() {
    if (!textHasContent) {
      textBuffer.clear();
      return;
    }
    std::string flushed = Trim(textBuffer);
    textBuffer.clear();
    textHasContent = false;
    if (flushed.empty()) {
      return;
    }
    auto* leaf = makeTextLayer(flushed, inheritedTextStyle);
    if (leaf) {
      parent->children.push_back(leaf);
    }
  };

  auto child = element->firstChild;
  while (child) {
    if (child->type == DOMNodeType::Text) {
      // Accumulate raw text; whitespace will be collapsed when flushed.
      for (char c : child->name) {
        if (std::isspace(static_cast<unsigned char>(c))) {
          if (!textBuffer.empty() && !std::isspace(static_cast<unsigned char>(textBuffer.back()))) {
            textBuffer.push_back(' ');
          }
        } else {
          textBuffer.push_back(c);
          textHasContent = true;
        }
      }
    } else if (child->type == DOMNodeType::Element) {
      std::string tag = ToLower(child->name);
      if (tag == "br") {
        textBuffer.push_back('\n');
        textHasContent = true;
      } else if (IsHtmlInlineTag(tag)) {
        // Flatten inline element's text into the surrounding text buffer.
        std::string inner = collectInlineText(child);
        if (!inner.empty()) {
          if (!textBuffer.empty() && !std::isspace(static_cast<unsigned char>(textBuffer.back()))) {
            textBuffer.push_back(' ');
          }
          textBuffer += inner;
          textHasContent = true;
        }
      } else {
        flushText();
        Layer* converted = convertElement(child, inheritedTextStyle, depth + 1);
        if (converted) {
          parent->children.push_back(converted);
        }
      }
    }
    child = child->nextSibling;
  }
  flushText();
}

Layer* HTMLParserContext::convertElement(const std::shared_ptr<DOMNode>& element,
                                         const TextStyle& parentTextStyle, int depth) {
  if (depth >= MAX_HTML_RECURSION_DEPTH) {
    warn("Maximum HTML recursion depth exceeded; deeper content was dropped.");
    return nullptr;
  }
  std::string tag = ToLower(element->name);
  auto props = readStyle(element);

  // Suppress non-renderable head-level tags entirely.
  if (tag == "head" || tag == "title" || tag == "meta" || tag == "link" || tag == "style" ||
      tag == "script" || tag == "noscript") {
    return nullptr;
  }
  if (tag == "img") {
    auto* layer = convertImage(element, props);
    if (layer) {
      applyStyle(layer, tag, props, parentTextStyle);
    }
    return layer;
  }
  if (tag == "svg") {
    auto* layer = convertInlineSvg(element, props);
    if (layer) {
      applyStyle(layer, tag, props, parentTextStyle);
    }
    return layer;
  }

  bool isContainer = IsHtmlContainerTag(tag) || IsHtmlTextTag(tag) || tag == "body";
  if (!isContainer) {
    if (_options.preserveUnknownElements) {
      auto* layer = _document->makeNode<Layer>();
      layer->customData["html-tag"] = tag;
      auto childStyle = applyStyle(layer, tag, props, parentTextStyle);
      convertChildren(element, layer, childStyle, depth);
      return layer;
    }
    warn("Unsupported HTML element <" + tag + "> dropped.");
    return nullptr;
  }

  // Block / text container.
  auto idAttr = getAttr(element, "id");
  auto* layer = _document->makeNode<Layer>(idAttr);
  if (!idAttr.empty()) {
    layer->id = idAttr;
  }
  layer->name = tag;

  // Heading defaults: bake in the visible font hierarchy before child cascade runs so authored
  // overrides in `style="..."` still win.
  TextStyle scratchStyle = parentTextStyle;
  if (IsHtmlTextTag(tag) && tag.size() == 2 && tag[0] == 'h' && tag[1] >= '1' && tag[1] <= '6') {
    scratchStyle.fontSize = HeadingFontSizePx(tag);
    scratchStyle.bold = true;
  }

  TextStyle childStyle = applyStyle(layer, tag, props, scratchStyle);

  bool isLeaf = isLeafTextContainer(element);
  if (isLeaf) {
    // A leaf text container holds only a TextBox; the TextBox is stretch-constrained
    // (left=right=top=bottom=0) so it fills the parent's content area for wrapping.
    // Layout mode would suppress those constraints, leaving the TextBox content-measured
    // (single-line). Clear the layout so the TextBox can stretch and wrap.
    layer->layout = LayoutMode::None;
    std::string text = collectInlineText(element);
    if (!text.empty()) {
      auto* textBox = _document->makeNode<TextBox>();
      textBox->left = 0;
      textBox->right = 0;
      textBox->top = 0;
      textBox->bottom = 0;
      textBox->textAlign = childStyle.textAlign;
      textBox->wordWrap = true;
      if (childStyle.lineHeight > 0) {
        textBox->lineHeight = childStyle.lineHeight;
      }

      auto* textNode = _document->makeNode<Text>();
      textNode->text = text;
      textNode->fontSize = childStyle.fontSize;
      textNode->letterSpacing = childStyle.letterSpacing;
      textNode->fauxBold = childStyle.bold;
      textNode->fauxItalic = childStyle.italic;
      if (!childStyle.fontFamily.empty()) {
        textNode->fontFamily = childStyle.fontFamily;
      }
      textBox->elements.push_back(textNode);

      // Color is applied via a Fill on the layer wrapping the TextBox: PAGX renders the
      // accumulated geometry painted by the Fill in document order. We attach the Fill *after*
      // the TextBox so it paints the glyphs.
      layer->contents.push_back(textBox);
      auto* solid = _document->makeNode<SolidColor>();
      solid->color = childStyle.color;
      auto* fill = _document->makeNode<Fill>();
      fill->color = solid;
      layer->contents.push_back(fill);
    }
  } else {
    convertChildren(element, layer, childStyle, depth);
  }
  splitBackgroundIfNeeded(layer);
  return layer;
}

void HTMLParserContext::splitBackgroundIfNeeded(Layer* layer) {
  if (!layer) return;
  // Detect a background painter — currently a stretch-fill Rectangle (left/right/top/bottom = 0)
  // in `contents`. When found together with non-zero padding (or actual content that would
  // otherwise be inset twice), introduce an inner content Layer carrying the padding+content.
  bool hasBackgroundRect = false;
  for (auto* el : layer->contents) {
    if (el->nodeType() == NodeType::Rectangle) {
      hasBackgroundRect = true;
      break;
    }
  }
  if (!hasBackgroundRect) {
    return;
  }
  bool hasPadding = !layer->padding.isZero();
  bool hasChildren = !layer->children.empty();
  // PAGX painters apply to accumulated, not-yet-painted geometry, so the order matters: the
  // background prefix must include both the geometry (Rectangle) and the painter that consumes
  // it (Fill/Stroke). The split point is the first content-bearing Element (Group/TextBox/Text);
  // everything from there on (including any later painters) belongs with the content.
  size_t splitIndex = layer->contents.size();
  for (size_t i = 0; i < layer->contents.size(); i++) {
    auto t = layer->contents[i]->nodeType();
    if (t == NodeType::Group || t == NodeType::TextBox || t == NodeType::Text) {
      splitIndex = i;
      break;
    }
  }
  std::vector<Element*> backgroundContents(layer->contents.begin(),
                                           layer->contents.begin() + splitIndex);
  std::vector<Element*> contentContents(layer->contents.begin() + splitIndex,
                                        layer->contents.end());
  bool hasContentEls = !contentContents.empty();
  if (!hasPadding && !hasChildren && !hasContentEls) {
    return;
  }
  if (!hasPadding && !hasContentEls && hasChildren) {
    // Nothing to inset against — the children's own constraints already place them within the
    // outer layer. No wrapper needed.
    return;
  }
  auto* inner = _document->makeNode<Layer>();
  inner->name = "content";
  inner->left = 0;
  inner->right = 0;
  inner->top = 0;
  inner->bottom = 0;
  inner->layout = layer->layout;
  inner->gap = layer->gap;
  inner->padding = layer->padding;
  inner->alignment = layer->alignment;
  inner->arrangement = layer->arrangement;
  inner->children = std::move(layer->children);
  layer->children.clear();
  layer->children.push_back(inner);
  inner->contents = std::move(contentContents);

  layer->contents = std::move(backgroundContents);
  layer->layout = LayoutMode::None;
  layer->gap = 0;
  layer->padding = {};
  layer->alignment = Alignment::Stretch;
  layer->arrangement = Arrangement::Start;
}

Layer* HTMLParserContext::makeTextLayer(const std::string& text, const TextStyle& style) {
  if (text.empty() || !_document) {
    return nullptr;
  }
  auto* layer = _document->makeNode<Layer>();
  layer->name = "text";
  layer->layout = LayoutMode::None;

  auto* textBox = _document->makeNode<TextBox>();
  textBox->left = 0;
  textBox->right = 0;
  textBox->top = 0;
  textBox->bottom = 0;
  textBox->textAlign = style.textAlign;
  textBox->wordWrap = true;
  if (style.lineHeight > 0) {
    textBox->lineHeight = style.lineHeight;
  }

  auto* textNode = _document->makeNode<Text>();
  textNode->text = text;
  textNode->fontSize = style.fontSize;
  textNode->letterSpacing = style.letterSpacing;
  textNode->fauxBold = style.bold;
  textNode->fauxItalic = style.italic;
  if (!style.fontFamily.empty()) {
    textNode->fontFamily = style.fontFamily;
  }
  textBox->elements.push_back(textNode);

  layer->contents.push_back(textBox);

  auto* solid = _document->makeNode<SolidColor>();
  solid->color = style.color;
  auto* fill = _document->makeNode<Fill>();
  fill->color = solid;
  layer->contents.push_back(fill);
  return layer;
}

Layer* HTMLParserContext::convertImage(
    const std::shared_ptr<DOMNode>& element,
    const std::unordered_map<std::string, std::string>& /*props*/) {
  auto src = getAttr(element, "src");
  if (src.empty()) {
    warn("<img> element missing 'src' attribute; dropped.");
    return nullptr;
  }
  auto* layer = _document->makeNode<Layer>();
  layer->name = "img";

  auto wAttr = getAttr(element, "width");
  auto hAttr = getAttr(element, "height");
  LengthUnit unit = LengthUnit::Px;
  float v = 0;
  if (!wAttr.empty() && ParseLength(wAttr, &v, &unit) && unit == LengthUnit::Px && v > 0 &&
      std::isnan(layer->width)) {
    layer->width = v;
  }
  if (!hAttr.empty() && ParseLength(hAttr, &v, &unit) && unit == LengthUnit::Px && v > 0 &&
      std::isnan(layer->height)) {
    layer->height = v;
  }

  // Paint the image as a full-bleed Rectangle with an ImagePattern fill. PAGX renders
  // the geometry then applies the fill; LetterBox scaleMode auto-fits per the layer bounds.
  auto* image = _document->makeNode<Image>();
  image->filePath = src;

  auto* rect = _document->makeNode<Rectangle>();
  rect->left = 0;
  rect->top = 0;
  rect->right = 0;
  rect->bottom = 0;

  auto* pattern = _document->makeNode<ImagePattern>();
  pattern->image = image;

  auto* fill = _document->makeNode<Fill>();
  fill->color = pattern;

  layer->contents.push_back(rect);
  layer->contents.push_back(fill);

  if (src.compare(0, 5, "data:") != 0 && image->data == nullptr) {
    warn("<img src='" + src +
         "'> references an external file; embed it before publishing or run `pagx resolve`.");
  }
  return layer;
}

void HTMLParserContext::serializeNode(const std::shared_ptr<DOMNode>& node, std::string& out) {
  if (!node) return;
  if (node->type == DOMNodeType::Text) {
    for (char c : node->name) {
      switch (c) {
        case '&':
          out += "&amp;";
          break;
        case '<':
          out += "&lt;";
          break;
        case '>':
          out += "&gt;";
          break;
        default:
          out.push_back(c);
      }
    }
    return;
  }
  out.push_back('<');
  out += node->name;
  for (auto& attr : node->attributes) {
    out.push_back(' ');
    out += attr.name;
    out += "=\"";
    for (char c : attr.value) {
      switch (c) {
        case '&':
          out += "&amp;";
          break;
        case '"':
          out += "&quot;";
          break;
        case '<':
          out += "&lt;";
          break;
        default:
          out.push_back(c);
      }
    }
    out.push_back('"');
  }
  if (!node->firstChild) {
    out += "/>";
    return;
  }
  out.push_back('>');
  auto child = node->firstChild;
  while (child) {
    serializeNode(child, out);
    child = child->nextSibling;
  }
  out += "</";
  out += node->name;
  out.push_back('>');
}

Layer* HTMLParserContext::convertInlineSvg(
    const std::shared_ptr<DOMNode>& element,
    const std::unordered_map<std::string, std::string>& /*props*/) {
  auto* layer = _document->makeNode<Layer>();
  layer->name = "svg";
  std::string serialized;
  serializeNode(element, serialized);
  layer->importDirective.format = "svg";
  layer->importDirective.content = serialized;
  return layer;
}

//==============================================================================
// Style application
//==============================================================================

namespace {

bool ParseLengthToPxOrPercent(const std::string& value, float* outPx, float* outPercent,
                              const std::string& propName, std::vector<std::string>* warnings) {
  float v = 0;
  LengthUnit unit;
  if (!ParseLength(value, &v, &unit)) {
    return false;
  }
  if (unit == LengthUnit::Px) {
    *outPx = v;
    return true;
  }
  if (unit == LengthUnit::Percent) {
    *outPercent = v;
    return true;
  }
  if (warnings) {
    warnings->push_back("Unsupported unit on '" + propName + ": " + value +
                        "'; only px and % are recognized.");
  }
  return false;
}

void ParsePaddingShorthand(const std::string& value, Padding* outPadding) {
  std::vector<std::string> tokens;
  std::string cur;
  for (char c : value) {
    if (std::isspace(static_cast<unsigned char>(c))) {
      if (!cur.empty()) {
        tokens.push_back(cur);
        cur.clear();
      }
    } else {
      cur.push_back(c);
    }
  }
  if (!cur.empty()) tokens.push_back(cur);
  auto px = [](const std::string& tok) {
    float v = 0;
    LengthUnit unit;
    if (ParseLength(tok, &v, &unit) && unit == LengthUnit::Px) return v;
    return 0.0f;
  };
  if (tokens.size() == 1) {
    float v = px(tokens[0]);
    outPadding->top = outPadding->right = outPadding->bottom = outPadding->left = v;
  } else if (tokens.size() == 2) {
    float v0 = px(tokens[0]);
    float h = px(tokens[1]);
    outPadding->top = outPadding->bottom = v0;
    outPadding->left = outPadding->right = h;
  } else if (tokens.size() == 3) {
    outPadding->top = px(tokens[0]);
    outPadding->left = outPadding->right = px(tokens[1]);
    outPadding->bottom = px(tokens[2]);
  } else if (tokens.size() >= 4) {
    outPadding->top = px(tokens[0]);
    outPadding->right = px(tokens[1]);
    outPadding->bottom = px(tokens[2]);
    outPadding->left = px(tokens[3]);
  }
}

Alignment AlignItemsFromCss(const std::string& value) {
  std::string v = ToLower(value);
  if (v == "center") return Alignment::Center;
  if (v == "flex-start" || v == "start") return Alignment::Start;
  if (v == "flex-end" || v == "end") return Alignment::End;
  return Alignment::Stretch;
}

Arrangement JustifyContentFromCss(const std::string& value) {
  std::string v = ToLower(value);
  if (v == "center") return Arrangement::Center;
  if (v == "flex-end" || v == "end") return Arrangement::End;
  if (v == "space-between") return Arrangement::SpaceBetween;
  if (v == "space-around") return Arrangement::SpaceAround;
  if (v == "space-evenly") return Arrangement::SpaceEvenly;
  return Arrangement::Start;
}

TextAlign TextAlignFromCss(const std::string& value) {
  std::string v = ToLower(value);
  if (v == "center") return TextAlign::Center;
  if (v == "right" || v == "end") return TextAlign::End;
  if (v == "justify") return TextAlign::Justify;
  return TextAlign::Start;
}

bool ParseFontWeight(const std::string& value, bool* outBold) {
  std::string v = ToLower(value);
  if (v == "bold" || v == "bolder") {
    *outBold = true;
    return true;
  }
  if (v == "normal" || v == "lighter") {
    *outBold = false;
    return true;
  }
  // Numeric 100..900.
  char* endp = nullptr;
  long n = std::strtol(value.c_str(), &endp, 10);
  if (endp == value.c_str()) return false;
  *outBold = (n >= 600);
  return true;
}

}  // namespace

TextStyle HTMLParserContext::applyStyle(Layer* layer, const std::string& tag,
                                        const std::unordered_map<std::string, std::string>& props,
                                        const TextStyle& parentTextStyle) {
  std::vector<std::string> localWarnings;
  TextStyle childStyle = parentTextStyle;

  // Container layout. Default for HTML block elements is Vertical (block flow); flex containers
  // default to Horizontal (CSS row).
  bool displayFlex = false;
  if (auto it = props.find("display"); it != props.end()) {
    std::string v = ToLower(it->second);
    if (v == "none") {
      layer->visible = false;
    } else if (v == "flex" || v == "inline-flex") {
      displayFlex = true;
    }
  }

  if (displayFlex) {
    layer->layout = LayoutMode::Horizontal;
  } else if (tag == "body" || tag == "html") {
    layer->layout = LayoutMode::Vertical;
  } else if (IsHtmlContainerTag(tag) && layer->layout == LayoutMode::None) {
    // Block-level container default: stack vertically.
    layer->layout = LayoutMode::Vertical;
  }

  if (auto it = props.find("flex-direction"); it != props.end()) {
    std::string v = ToLower(it->second);
    if (v == "row" || v == "row-reverse") layer->layout = LayoutMode::Horizontal;
    if (v == "column" || v == "column-reverse") layer->layout = LayoutMode::Vertical;
    if (v == "row-reverse" || v == "column-reverse") {
      localWarnings.push_back("flex-direction reverse variants are not supported; using forward.");
    }
  }

  if (auto it = props.find("align-items"); it != props.end()) {
    layer->alignment = AlignItemsFromCss(it->second);
  }
  if (auto it = props.find("justify-content"); it != props.end()) {
    layer->arrangement = JustifyContentFromCss(it->second);
  }

  // Sizing.
  auto applySize = [&](const std::string& key, float* sizeField, float* percentField) {
    auto it = props.find(key);
    if (it == props.end()) return;
    float px = NAN;
    float pct = NAN;
    if (ParseLengthToPxOrPercent(it->second, &px, &pct, key, &localWarnings)) {
      if (!std::isnan(px)) *sizeField = px;
      if (!std::isnan(pct)) *percentField = pct;
    }
  };
  applySize("width", &layer->width, &layer->percentWidth);
  applySize("height", &layer->height, &layer->percentHeight);

  // Padding.
  Padding padding = layer->padding;
  if (auto it = props.find("padding"); it != props.end()) {
    Padding parsed = {};
    ParsePaddingShorthand(it->second, &parsed);
    padding = parsed;
  }
  auto applySide = [&](const std::string& key, float* side) {
    auto it = props.find(key);
    if (it == props.end()) return;
    float v = 0;
    LengthUnit unit;
    if (ParseLength(it->second, &v, &unit) && unit == LengthUnit::Px) {
      *side = v;
    }
  };
  applySide("padding-top", &padding.top);
  applySide("padding-right", &padding.right);
  applySide("padding-bottom", &padding.bottom);
  applySide("padding-left", &padding.left);
  layer->padding = padding;

  // Gap.
  if (auto it = props.find("gap"); it != props.end()) {
    float v = 0;
    LengthUnit unit;
    if (ParseLength(it->second, &v, &unit) && unit == LengthUnit::Px) {
      layer->gap = v;
    }
  }

  // Flex grow.
  if (auto it = props.find("flex"); it != props.end()) {
    char* endp = nullptr;
    float v = static_cast<float>(std::strtod(it->second.c_str(), &endp));
    if (endp != it->second.c_str()) {
      layer->flex = v;
    }
  }
  if (auto it = props.find("flex-grow"); it != props.end()) {
    char* endp = nullptr;
    float v = static_cast<float>(std::strtod(it->second.c_str(), &endp));
    if (endp != it->second.c_str()) {
      layer->flex = v;
    }
  }

  // Opacity.
  if (auto it = props.find("opacity"); it != props.end()) {
    char* endp = nullptr;
    float v = static_cast<float>(std::strtod(it->second.c_str(), &endp));
    if (endp != it->second.c_str()) {
      layer->alpha = std::max(0.0f, std::min(1.0f, v));
    }
  }

  // Positioning.
  if (auto it = props.find("position"); it != props.end()) {
    std::string v = ToLower(it->second);
    if (v == "absolute" || v == "fixed") {
      layer->includeInLayout = false;
    }
  }
  auto applyEdge = [&](const std::string& key, float* edge) {
    auto it = props.find(key);
    if (it == props.end()) return;
    float v = 0;
    LengthUnit unit;
    if (ParseLength(it->second, &v, &unit) && unit == LengthUnit::Px) {
      *edge = v;
    }
  };
  applyEdge("left", &layer->left);
  applyEdge("right", &layer->right);
  applyEdge("top", &layer->top);
  applyEdge("bottom", &layer->bottom);

  // Overflow.
  if (auto it = props.find("overflow"); it != props.end()) {
    std::string v = ToLower(it->second);
    if (v == "hidden" || v == "clip") {
      layer->clipToBounds = true;
    }
  }

  // Background color → Rectangle{100% × 100%} + Fill.
  if (auto it = props.find("background-color"); it != props.end()) {
    Color bg = {};
    if (ParseCSSColor(it->second, &bg)) {
      auto* rect = _document->makeNode<Rectangle>();
      rect->left = 0;
      rect->top = 0;
      rect->right = 0;
      rect->bottom = 0;
      // Border-radius applies to this background rectangle.
      if (auto rIt = props.find("border-radius"); rIt != props.end()) {
        float r = 0;
        LengthUnit unit;
        if (ParseLength(rIt->second, &r, &unit) && unit == LengthUnit::Px) {
          rect->roundness = r;
        }
      }
      layer->contents.push_back(rect);
      auto* solid = _document->makeNode<SolidColor>();
      solid->color = bg;
      auto* fill = _document->makeNode<Fill>();
      fill->color = solid;
      layer->contents.push_back(fill);
    } else {
      localWarnings.push_back("Unrecognized background-color value '" + it->second + "'.");
    }
  } else if (auto rIt = props.find("border-radius"); rIt != props.end()) {
    // border-radius without background still rounds any later painted geometry attached to the
    // layer. Add a RoundCorner modifier.
    float r = 0;
    LengthUnit unit;
    if (ParseLength(rIt->second, &r, &unit) && unit == LengthUnit::Px) {
      auto* rc = _document->makeNode<RoundCorner>();
      rc->radius = r;
      layer->contents.push_back(rc);
    }
  }

  // Border (shorthand `border: Npx solid color`). We map to a Stroke painter that paints the
  // background rectangle if one was added.
  if (auto it = props.find("border"); it != props.end()) {
    std::vector<std::string> parts;
    std::string cur;
    for (char c : it->second) {
      if (std::isspace(static_cast<unsigned char>(c))) {
        if (!cur.empty()) {
          parts.push_back(cur);
          cur.clear();
        }
      } else {
        cur.push_back(c);
      }
    }
    if (!cur.empty()) parts.push_back(cur);
    float borderWidth = 0;
    Color borderColor = {0, 0, 0, 1, ColorSpace::SRGB};
    LengthUnit unit;
    bool widthSet = false;
    bool colorSet = false;
    for (auto& tok : parts) {
      if (!widthSet && ParseLength(tok, &borderWidth, &unit) && unit == LengthUnit::Px) {
        widthSet = true;
        continue;
      }
      Color c;
      if (ParseCSSColor(tok, &c)) {
        borderColor = c;
        colorSet = true;
      }
    }
    if (widthSet) {
      // If no background rectangle was added, append one so the stroke has geometry to paint.
      bool hasGeometry = false;
      for (auto* el : layer->contents) {
        if (el->nodeType() == NodeType::Rectangle) {
          hasGeometry = true;
          break;
        }
      }
      if (!hasGeometry) {
        auto* rect = _document->makeNode<Rectangle>();
        rect->left = 0;
        rect->top = 0;
        rect->right = 0;
        rect->bottom = 0;
        if (auto rIt = props.find("border-radius"); rIt != props.end()) {
          float r = 0;
          if (ParseLength(rIt->second, &r, &unit) && unit == LengthUnit::Px) {
            rect->roundness = r;
          }
        }
        layer->contents.push_back(rect);
      }
      auto* solid = _document->makeNode<SolidColor>();
      solid->color = borderColor;
      auto* stroke = _document->makeNode<Stroke>();
      stroke->color = solid;
      stroke->width = borderWidth;
      layer->contents.push_back(stroke);
      if (!colorSet) {
        localWarnings.push_back("border declaration '" + it->second +
                                "' has no color; defaulting to black.");
      }
    }
  }

  // Box-shadow → DropShadowStyle. Accept `Hpx Vpx Bpx color` (and tolerate spread by ignoring it).
  if (auto it = props.find("box-shadow"); it != props.end()) {
    std::string v = it->second;
    bool inset = false;
    std::string lower = ToLower(v);
    if (lower.find("inset") != std::string::npos) {
      inset = true;
      // Strip the `inset` keyword from the working copy (handled via fall-back search below).
    }
    // Tokenize on whitespace not inside a function call.
    std::vector<std::string> tokens;
    std::string cur;
    int parenDepth = 0;
    for (char c : v) {
      if (c == '(') parenDepth++;
      if (c == ')') parenDepth--;
      if (parenDepth == 0 && std::isspace(static_cast<unsigned char>(c))) {
        if (!cur.empty()) {
          tokens.push_back(cur);
          cur.clear();
        }
      } else {
        cur.push_back(c);
      }
    }
    if (!cur.empty()) tokens.push_back(cur);

    std::vector<float> lengths;
    Color shadowColor = {0, 0, 0, 1, ColorSpace::SRGB};
    bool colorSet = false;
    for (auto& tok : tokens) {
      if (ToLower(tok) == "inset") continue;
      float lv = 0;
      LengthUnit unit;
      if (lengths.size() < 4 && ParseLength(tok, &lv, &unit) && unit == LengthUnit::Px) {
        lengths.push_back(lv);
        continue;
      }
      Color c;
      if (ParseCSSColor(tok, &c)) {
        shadowColor = c;
        colorSet = true;
      }
    }
    (void)colorSet;
    if (lengths.size() >= 2 && !inset) {
      auto* shadow = _document->makeNode<DropShadowStyle>();
      shadow->offsetX = lengths[0];
      shadow->offsetY = lengths[1];
      shadow->blurX = lengths.size() >= 3 ? lengths[2] : 0;
      shadow->blurY = shadow->blurX;
      shadow->color = shadowColor;
      layer->styles.push_back(shadow);
    } else if (inset) {
      localWarnings.push_back("inset box-shadow not yet supported; dropped.");
    }
  }

  // Text-related cascade.
  if (auto it = props.find("color"); it != props.end()) {
    Color c;
    if (ParseCSSColor(it->second, &c)) {
      childStyle.color = c;
    }
  }
  if (auto it = props.find("font-family"); it != props.end()) {
    // Take first family (split on comma) and trim quotes.
    auto comma = it->second.find(',');
    std::string first = comma == std::string::npos ? it->second : it->second.substr(0, comma);
    first = Trim(first);
    if (!first.empty() && (first.front() == '"' || first.front() == '\'')) {
      first.erase(first.begin());
    }
    if (!first.empty() && (first.back() == '"' || first.back() == '\'')) {
      first.pop_back();
    }
    childStyle.fontFamily = first;
  }
  if (auto it = props.find("font-size"); it != props.end()) {
    float v = 0;
    LengthUnit unit;
    if (ParseLength(it->second, &v, &unit) && unit == LengthUnit::Px) {
      childStyle.fontSize = v;
    }
  }
  if (auto it = props.find("font-weight"); it != props.end()) {
    bool bold = false;
    if (ParseFontWeight(it->second, &bold)) {
      childStyle.bold = bold;
    }
  }
  if (auto it = props.find("font-style"); it != props.end()) {
    std::string v = ToLower(it->second);
    childStyle.italic = (v == "italic" || v == "oblique");
  }
  if (auto it = props.find("letter-spacing"); it != props.end()) {
    float v = 0;
    LengthUnit unit;
    if (ParseLength(it->second, &v, &unit) && unit == LengthUnit::Px) {
      childStyle.letterSpacing = v;
    }
  }
  if (auto it = props.find("line-height"); it != props.end()) {
    float v = 0;
    LengthUnit unit;
    if (ParseLength(it->second, &v, &unit) && unit == LengthUnit::Px) {
      childStyle.lineHeight = v;
    } else {
      // Unitless multiplier — apply to current font-size.
      char* endp = nullptr;
      float multiplier = static_cast<float>(std::strtod(it->second.c_str(), &endp));
      if (endp != it->second.c_str() && multiplier > 0) {
        childStyle.lineHeight = childStyle.fontSize * multiplier;
      }
    }
  }
  if (auto it = props.find("text-align"); it != props.end()) {
    childStyle.textAlign = TextAlignFromCss(it->second);
  }

  for (auto& w : localWarnings) {
    warn(w);
  }
  return childStyle;
}

}  // namespace

//==============================================================================
// Public API
//==============================================================================

std::shared_ptr<PAGXDocument> HTMLImporter::Parse(const std::string& filePath,
                                                  const Options& options) {
  HTMLParserContext ctx(options);
  return ctx.parseFile(filePath);
}

std::shared_ptr<PAGXDocument> HTMLImporter::Parse(const uint8_t* data, size_t length,
                                                  const Options& options) {
  HTMLParserContext ctx(options);
  return ctx.parse(data, length);
}

std::shared_ptr<PAGXDocument> HTMLImporter::ParseString(const std::string& htmlContent,
                                                        const Options& options) {
  return Parse(reinterpret_cast<const uint8_t*>(htmlContent.data()), htmlContent.size(), options);
}

}  // namespace pagx
