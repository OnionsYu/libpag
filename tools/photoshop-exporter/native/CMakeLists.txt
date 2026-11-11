set(UXP_SDK_DIR "" CACHE PATH "Root of the Adobe UXP Hybrid Plugin SDK")
if (NOT EXISTS "${UXP_SDK_DIR}/src/utilities/UxpAddon.h")
    message(FATAL_ERROR "Set UXP_SDK_DIR to the extracted official Adobe Hybrid SDK")
endif ()
if (PAG_BUILD_SHARED OR PAG_BUILD_FRAMEWORK)
    message(FATAL_ERROR "The Photoshop addon requires PAG_BUILD_SHARED=OFF and PAG_BUILD_FRAMEWORK=OFF")
endif ()
if (NOT APPLE AND NOT WIN32)
    message(FATAL_ERROR "The Photoshop addon supports macOS and Windows")
endif ()

add_library(pagx-photoshop-addon MODULE Addon.cpp)
target_include_directories(pagx-photoshop-addon PRIVATE
        ${PROJECT_SOURCE_DIR}/include
        ${TGFX_DIR}/include
        ${TGFX_DIR}/third_party/json/include)
target_include_directories(pagx-photoshop-addon SYSTEM PRIVATE "${UXP_SDK_DIR}/src/utilities")
target_link_libraries(pagx-photoshop-addon PRIVATE pag)
set_target_properties(pagx-photoshop-addon PROPERTIES
        PREFIX "" OUTPUT_NAME "pagx-photoshop" SUFFIX ".uxpaddon")
if (APPLE)
    target_link_options(pagx-photoshop-addon PRIVATE
            "LINKER:-exported_symbol,_uxp_addon_init"
            "LINKER:-exported_symbol,_uxp_addon_terminate")
endif ()
