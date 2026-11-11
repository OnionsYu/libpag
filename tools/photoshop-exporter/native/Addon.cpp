/////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Tencent. All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy at http://www.apache.org/licenses/LICENSE-2.0
/////////////////////////////////////////////////////////////////////////////////////////////////

#include <UxpAddon.h>
#include <cmath>
#include <cstring>
#include <nlohmann/json.hpp>
#include <stdexcept>
#include <unordered_map>
#include "pagx/PAGXExporter.h"
#include "pagx/PAGXImporter.h"
#include "pagx/PhotoshopImporter.h"
#include "tgfx/core/ImageCodec.h"
#include "tgfx/core/Pixmap.h"

namespace {
using JSON = nlohmann::json;
constexpr size_t MaxPixels = 64 * 1024 * 1024;
constexpr size_t MaxEncodedBytes = 128 * 1024 * 1024;
struct Session {
  photoshop::Document source;
  pagx::PhotoshopConversionPlan plan;
  bool forceComposite = false;
  size_t encodedBytes = 0;
  std::vector<photoshop::RasterAsset> assets;
};
// Access is confined to UXP's scripting thread. No Photoshop or SDK objects escape callbacks.
std::unordered_map<std::string, Session> sessions;
uint64_t nextId = 0;

void CheckStatus(addon_status status) {
  if (status != addon_ok) throw std::runtime_error("UXP addon argument or API error");
}
std::vector<addon_value> Args(addon_env env, addon_callback_info info, size_t count) {
  std::vector<addon_value> args(count + 1);
  size_t actual = args.size();
  CheckStatus(
      UxpAddonApis.uxp_addon_get_cb_info(env, info, &actual, args.data(), nullptr, nullptr));
  if (actual != count) throw std::invalid_argument("Incorrect argument count");
  args.resize(count);
  return args;
}
std::string String(addon_env env, addon_value value, size_t limit = 32 * 1024 * 1024) {
  size_t size = 0;
  CheckStatus(UxpAddonApis.uxp_addon_get_value_string_utf8(env, value, nullptr, 0, &size));
  if (size > limit) throw std::invalid_argument("String argument exceeds limit");
  std::vector<char> text(size + 1);
  CheckStatus(
      UxpAddonApis.uxp_addon_get_value_string_utf8(env, value, text.data(), text.size(), &size));
  return {text.data(), size};
}
addon_value String(addon_env env, const std::string& value) {
  addon_value result = nullptr;
  CheckStatus(UxpAddonApis.uxp_addon_create_string_utf8(env, value.data(), value.size(), &result));
  return result;
}
addon_value Undefined(addon_env env) {
  addon_value result = nullptr;
  CheckStatus(UxpAddonApis.uxp_addon_get_undefined(env, &result));
  return result;
}
int PixelInteger(const JSON& value) {
  if (!value.is_number_integer() || value.get<double>() < 0 || value.get<double>() > INT32_MAX) {
    throw std::invalid_argument("Invalid capture integer");
  }
  return value.get<int>();
}
template <typename Function>
addon_value Guard(addon_env env, Function function) noexcept {
  try {
    return function();
  } catch (const std::exception& error) {
    UxpAddonApis.uxp_addon_throw_error(env, "PAGX_EXPORT", error.what());
  } catch (...) {
    UxpAddonApis.uxp_addon_throw_error(env, "PAGX_EXPORT", "Unknown native export error");
  }
  return nullptr;
}
addon_value Info(addon_env env, addon_callback_info) {
  return Guard(env, [&] {
    return String(env, JSON{{"protocolVersion", 2},
                            {"schemaVersion", 2},
                            {"version", "0.2.0"},
                            {"maxPixels", MaxPixels},
                            {"maxEncodedBytes", MaxEncodedBytes}}
                           .dump());
  });
}
addon_value Begin(addon_env env, addon_callback_info info) {
  return Guard(env, [&] {
    auto args = Args(env, info, 2);
    if (sessions.size() >= 2) throw std::runtime_error("Too many active exports");
    Session session;
    session.source = photoshop::ParseSnapshot(String(env, args[0]));
    CheckStatus(UxpAddonApis.uxp_addon_get_value_bool(env, args[1], &session.forceComposite));
    session.plan = pagx::PhotoshopImporter::Analyze(session.source, session.forceComposite);
    const auto id = std::to_string(++nextId);
    JSON result{{"sessionId", id},
                {"composite", session.plan.composite},
                {"rasters", JSON::array()},
                {"diagnostics", JSON::array()}};
    for (const auto& request : session.plan.rasters) {
      result["rasters"].push_back({{"id", request.id},
                                   {"layerId", request.layerId},
                                   {"layerIds", request.layerIds},
                                   {"parentId", request.parentId}});
    }
    const auto& statistics = session.plan.statistics;
    result["statistics"] = {{"vectorShapes", statistics.vectorShapes},
                            {"outlinedTexts", statistics.outlinedTexts},
                            {"sourceImages", statistics.sourceImages},
                            {"fallbackScopes", statistics.fallbackScopes}};
    for (const auto& diagnostic : session.plan.diagnostics) {
      result["diagnostics"].push_back({{"code", diagnostic.code},
                                       {"layerId", diagnostic.layerId},
                                       {"message", diagnostic.message}});
    }
    auto response = String(env, result.dump());
    sessions.emplace(id, std::move(session));
    return response;
  });
}
addon_value AddRaster(addon_env env, addon_callback_info info) {
  return Guard(env, [&] {
    auto args = Args(env, info, 3);
    auto& session = sessions.at(String(env, args[0], 32));
    auto metadata = JSON::parse(String(env, args[1], 4096));
    const auto requestId = metadata.at("requestId").get<std::string>();
    if (session.assets.size() >= session.plan.rasters.size())
      throw std::invalid_argument("Too many captures");
    if (requestId != session.plan.rasters[session.assets.size()].id) {
      throw std::invalid_argument("Captures must follow plan order");
    }
    void* bytes = nullptr;
    size_t length = 0;
    CheckStatus(UxpAddonApis.uxp_addon_get_arraybuffer_info(env, args[2], &bytes, &length));
    const auto width = PixelInteger(metadata.at("width"));
    const auto height = PixelInteger(metadata.at("height"));
    const auto components = PixelInteger(metadata.at("components"));
    const auto componentSize = PixelInteger(metadata.at("componentSize"));
    const auto bytesPerComponent = componentSize / 8;
    if ((componentSize != 8 && componentSize != 16) ||
        (componentSize == 16 && metadata.value("fullRange", false) != true) ||
        metadata.at("colorProfile") != "sRGB IEC61966-2.1" || metadata.at("alpha") != "straight" ||
        (components != 3 && components != 4) || width < 0 || height < 0 ||
        (width == 0) != (height == 0) ||
        static_cast<uint64_t>(width) * static_cast<uint64_t>(height) * bytesPerComponent >
            MaxPixels ||
        length != static_cast<size_t>(width) * static_cast<size_t>(height) * components *
                      bytesPerComponent) {
      throw std::invalid_argument("Invalid capture pixel contract or memory limit exceeded");
    }
    photoshop::RasterAsset asset;
    asset.requestId = requestId;
    if (!metadata.at("left").is_number() || !metadata.at("top").is_number()) {
      throw std::invalid_argument("Invalid capture origin");
    }
    asset.bounds = {metadata.at("left").get<float>(), metadata.at("top").get<float>(),
                    static_cast<float>(width), static_cast<float>(height)};
    if (!std::isfinite(asset.bounds.left) || !std::isfinite(asset.bounds.top)) {
      throw std::invalid_argument("Invalid capture origin");
    }
    if (width > 0) {
      std::vector<uint8_t> rgba;
      if (components == 3 || componentSize == 16) {
        rgba.resize(static_cast<size_t>(width) * height * 4);
        auto input = static_cast<const uint8_t*>(bytes);
        const size_t stride = components * bytesPerComponent;
        for (size_t i = 0, j = 0; i < length; i += stride, j += 4) {
          rgba[j + 3] = 255;
          for (int channel = 0; channel < components; ++channel) {
            if (componentSize == 16) {
              // Uint16Array uses native byte order. fullRange=true means 0..65535, not 0..32768.
              uint16_t value = 0;
              std::memcpy(&value, input + i + channel * 2, sizeof(value));
              rgba[j + channel] = static_cast<uint8_t>((static_cast<uint32_t>(value) + 128) / 257);
            } else {
              rgba[j + channel] = input[i + channel];
            }
          }
        }
        bytes = rgba.data();
      }
      auto imageInfo =
          tgfx::ImageInfo::Make(width, height, tgfx::ColorType::RGBA_8888,
                                tgfx::AlphaType::Unpremultiplied, 0, tgfx::ColorSpace::SRGB());
      auto png =
          tgfx::ImageCodec::Encode(tgfx::Pixmap(imageInfo, bytes), tgfx::EncodedFormat::PNG, 100);
      if (!png) throw std::runtime_error("PNG encoding failed");
      if (png->size() > MaxEncodedBytes - session.encodedBytes) {
        throw std::runtime_error("Export images exceed the 128 MiB encoded-resource limit");
      }
      asset.png.assign(png->bytes(), png->bytes() + png->size());
      session.encodedBytes += png->size();
    }
    session.assets.push_back(std::move(asset));
    return Undefined(env);
  });
}
addon_value Finish(addon_env env, addon_callback_info info) {
  return Guard(env, [&] {
    auto args = Args(env, info, 1);
    auto node = sessions.extract(String(env, args[0], 32));
    if (node.empty()) throw std::invalid_argument("Export session does not exist");
    auto& session = node.mapped();
    auto document =
        pagx::PhotoshopImporter::Convert(session.source, session.assets, session.forceComposite);
    auto xml = pagx::PAGXExporter::ToXML(*document);
    if (xml.size() > 256 * 1024 * 1024) throw std::runtime_error("PAGX exceeds 256 MiB");
    auto validation = pagx::PAGXImporter::FromXML(xml);
    if (!validation || !validation->errors.empty() || validation->hasUnresolvedImports() ||
        !validation->getExternalFilePaths().empty()) {
      throw std::runtime_error("Generated PAGX failed structural validation");
    }
    return String(env, xml);
  });
}
addon_value Discard(addon_env env, addon_callback_info info) {
  return Guard(env, [&] {
    auto args = Args(env, info, 1);
    sessions.erase(String(env, args[0], 32));
    return Undefined(env);
  });
}

addon_value Init(addon_env env, addon_value exports, const addon_apis&) {
  const std::pair<const char*, addon_callback> methods[] = {{"info", Info},
                                                            {"begin", Begin},
                                                            {"addRaster", AddRaster},
                                                            {"finish", Finish},
                                                            {"discard", Discard}};
  for (const auto& method : methods) {
    addon_value fn = nullptr;
    CheckStatus(UxpAddonApis.uxp_addon_create_function(env, method.first, std::strlen(method.first),
                                                       method.second, nullptr, &fn));
    CheckStatus(UxpAddonApis.uxp_addon_set_named_property(env, exports, method.first, fn));
  }
  return exports;
}
void Terminate(addon_env) {
  sessions.clear();
}
}  // namespace
UXP_ADDON_INIT(Init)
UXP_ADDON_TERMINATE(Terminate)
