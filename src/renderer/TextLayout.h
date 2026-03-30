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

#include <memory>
#include <string>
#include <unordered_map>
#include <vector>
#include "ShapedText.h"
#include "VerticalTextUtils.h"
#include "pagx/PAGXDocument.h"
#include "tgfx/core/Font.h"
#include "tgfx/core/Typeface.h"

namespace pagx {

/**
 * Holds font location info and creates the Typeface on demand. Once created, the Typeface is
 * cached for subsequent access.
 */
class TypefaceHolder {
 public:
  explicit TypefaceHolder(std::shared_ptr<tgfx::Typeface> typeface);

  TypefaceHolder(std::string path, int ttcIndex, std::string fontFamily, std::string fontStyle);

  std::shared_ptr<tgfx::Typeface> getTypeface();

  const std::string& getFontFamily() const;

  const std::string& getFontStyle() const;

 private:
  std::string path = {};
  std::string fontFamily = {};
  std::string fontStyle = {};
  int ttcIndex = 0;
  std::shared_ptr<tgfx::Typeface> typeface = nullptr;
};

/**
 * TextLayout performs text layout on PAGXDocument, converting Text elements into positioned glyph
 * data (TextBlob). It handles font matching, fallback, text shaping, and layout (alignment, line
 * breaking, etc.).
 */
class TextLayout {
 public:
  TextLayout() = default;

  /**
   * Registers a typeface for font matching. Registered typefaces are matched first by fontFamily
   * and fontStyle. If no registered typeface matches, the system font is used.
   */
  void registerTypeface(std::shared_ptr<tgfx::Typeface> typeface);

  /**
   * Adds fallback typefaces used when a character is not found in the primary font (either
   * registered or system). Typefaces are tried in order until one containing the character is found.
   */
  void addFallbackTypefaces(std::vector<std::shared_ptr<tgfx::Typeface>> typefaces);

  /**
   * Adds a deferred fallback font that will be loaded on demand when needed during text shaping.
   * @param path The font file path (empty if using fontFamily for name-based loading).
   * @param ttcIndex The face index within a TTC font collection.
   * @param fontFamily The font family name for matching and name-based loading.
   * @param fontStyle The font style name.
   */
  void addFallbackFont(const std::string& path, int ttcIndex, const std::string& fontFamily,
                       const std::string& fontStyle);

  /**
   * Performs text layout for all Text nodes in the document.
   */
  TextLayoutResult layout(PAGXDocument* document);

 private:
  friend class TextLayoutContext;

  struct FontKey {
    std::string family = {};
    std::string style = {};

    bool operator==(const FontKey& other) const {
      return family == other.family && style == other.style;
    }
  };

  struct FontKeyHash {
    size_t operator()(const FontKey& key) const {
      return std::hash<std::string>()(key.family) ^ (std::hash<std::string>()(key.style) << 1);
    }
  };

  std::unordered_map<FontKey, std::shared_ptr<tgfx::Typeface>, FontKeyHash> registeredTypefaces =
      {};
  std::vector<TypefaceHolder> fallbackTypefaces = {};
};

class Text;
class TextBox;
class Font;

class TextLayoutContext {
 public:
  TextLayoutContext(TextLayout* textLayout, PAGXDocument* document);

  TextLayoutResult run();

 private:
  struct ShapedGlyphRun {
    tgfx::Font font = {};
    std::vector<tgfx::GlyphID> glyphIDs = {};
    std::vector<float> xPositions = {};
    float startX = 0;
    bool canUseDefaultMode = true;
  };

  struct GlyphInfo {
    tgfx::GlyphID glyphID = 0;
    tgfx::Font font = {};
    float advance = 0;
    float xPosition = 0;
    int32_t unichar = 0;
    float fontSize = 0;
    float ascent = 0;
    float descent = 0;
    float fontLineHeight = 0;
    Text* sourceText = nullptr;
    uint32_t cluster = 0;
    float xOffset = 0;
    float yOffset = 0;
    uint8_t bidiLevel = 0;
  };

  struct LineInfo {
    std::vector<GlyphInfo> glyphs = {};
    float width = 0;
    float maxAscent = 0;
    float maxDescent = 0;
    float maxLineHeight = 0;
    float metricsHeight = 0;
    float roundingRatio = 1.0f;
  };

  struct VerticalGlyphInfo {
    std::vector<GlyphInfo> glyphs = {};
    VerticalOrientation orientation = VerticalOrientation::Upright;
    float height = 0;
    float width = 0;
    float leadingSquash = 0;
    bool canBreakBefore = false;
  };

  struct ColumnInfo {
    std::vector<VerticalGlyphInfo> glyphs = {};
    float height = 0;
    float maxColumnWidth = 0;
  };

  struct ShapedInfo {
    Text* text = nullptr;
    std::vector<ShapedGlyphRun> runs = {};
    float totalWidth = 0;
    std::vector<GlyphInfo> allGlyphs = {};
    bool paragraphRTL = false;
  };

  struct TextSegment {
    size_t start = 0;
    size_t length = 0;
    bool isNewline = false;
    bool isTab = false;
#ifdef PAG_BUILD_PAGX
    uint8_t bidiLevel = 0;
#endif
  };

  using FontKey = TextLayout::FontKey;

  static void SplitTextSegments(const std::string& content, size_t rangeStart, size_t rangeLength,
                                uint8_t bidiLevel, std::vector<TextSegment>& segments);

  static GlyphInfo CreateNewlineGlyph(Text* text, const tgfx::Font& font);

  static GlyphInfo CreateTabGlyph(Text* text, const tgfx::Font& font, float tabWidth,
                                  float xPosition);

  void storeShapedText(Text* text, ShapedText&& shapedText);

  static int StylePriority(const std::string& style);

  static bool IsPreferredFontKey(const FontKey& candidate, const FontKey& current);

  void processLayer(Layer* layer);

  std::vector<Text*> processScope(const std::vector<Element*>& elements);

  bool tryUseEmbeddedGlyphRuns(const std::vector<Text*>& textElements);

  void processTextWithLayout(std::vector<Text*>& textElements, const TextBox* textBox);

  void processTextWithoutLayout(Text* text);

  void buildTextBlobWithoutLayoutSingleLine(Text* text, const ShapedInfo& info);

  void buildTextBlobWithoutLayoutMultiLine(Text* text, const ShapedInfo& info);

  ShapedText buildShapedTextFromEmbeddedGlyphRuns(const Text* text);

  std::shared_ptr<tgfx::Typeface> buildTypefaceFromFont(const Font* fontNode);

  void shapeText(Text* text, ShapedInfo& info, bool vertical = false);

  std::vector<LineInfo> layoutLines(const std::vector<GlyphInfo>& allGlyphs,
                                    const TextBox* textBox);

#ifdef PAG_BUILD_PAGX
  static void ApplyPunctuationSquashToLines(std::vector<LineInfo>& lines);
#endif

  static void FinishLine(LineInfo* line, float lineHeight, float newlineFontLineHeight);

  void buildTextBlobWithLayout(const TextBox* textBox, const std::vector<LineInfo>& lines,
                               bool paragraphRTL = false);

  static void RemoveTrailingLetterSpacing(std::vector<VerticalGlyphInfo>& glyphs);

  static void FinishColumn(ColumnInfo* column, float lineHeight, float newlineFontLineHeight);

  std::vector<ColumnInfo> layoutColumns(const std::vector<GlyphInfo>& allGlyphs,
                                        const TextBox* textBox);

#ifdef PAG_BUILD_PAGX
  static bool IsSquashable(const VerticalGlyphInfo& vg);

  static void ApplyPunctuationSquashToColumns(std::vector<ColumnInfo>& columns);
#endif

  void buildTextBlobVertical(const TextBox* textBox, const std::vector<ColumnInfo>& columns);

  std::shared_ptr<tgfx::Typeface> findTypeface(const std::string& fontFamily,
                                               const std::string& fontStyle);

  TextLayout* textLayout = nullptr;
  PAGXDocument* document = nullptr;
  ShapedTextMap result = {};
  std::vector<Text*> textOrder = {};
  std::unordered_map<const Font*, std::shared_ptr<tgfx::Typeface>> fontCache = {};
};

}  // namespace pagx
