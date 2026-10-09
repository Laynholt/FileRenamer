cmake_minimum_required(VERSION 3.21)

project(FileRenamer
    VERSION 1.1.0
    DESCRIPTION "Pattern-based file renamer for Windows"
    LANGUAGES CXX RC
)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

if(WIN32)
    add_definitions(-DUNICODE -D_UNICODE)
endif()

set(SOURCES
    src/main.cpp
    src/Application.cpp
    src/ApplicationSuggestions.cpp
    src/ExplorerPathProvider.cpp
    src/RenamerService.cpp
    src/RenameSuggestions.cpp
    src/ToolTip.cpp
    src/UpdateService.cpp
    src/UiRenderer.cpp
    src/resource.rc
)

set(HEADERS
    src/Application.h
    src/ExplorerPathProvider.h
    src/RenamerService.h
    src/RenameSuggestions.h
    src/ToolTip.h
    src/UpdateService.h
    src/UiRenderer.h
    src/resource.h
)

add_executable(${PROJECT_NAME} WIN32 ${SOURCES} ${HEADERS})

set(WCW_SOURCE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/../win32-custom-widgets" CACHE PATH "Win32 Custom Widgets source directory")
add_subdirectory("${WCW_SOURCE_DIR}" "${CMAKE_BINARY_DIR}/win32-custom-widgets")
target_link_libraries(${PROJECT_NAME} PRIVATE Win32CustomWidgets::Win32CustomWidgets)

target_include_directories(${PROJECT_NAME} PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
)

if(WIN32)
    target_link_libraries(${PROJECT_NAME} PRIVATE
        comctl32
        gdi32
        gdiplus
        ole32
        oleaut32
        shell32
        shlwapi
        user32
        winhttp
    )
endif()

if(MSVC)
    target_compile_options(${PROJECT_NAME} PRIVATE
        /W4
        /permissive-
        /utf-8
    )
elseif(MINGW)
    target_compile_options(${PROJECT_NAME} PRIVATE
        -Wall
        -Wextra
        -Wpedantic
    )
    target_link_options(${PROJECT_NAME} PRIVATE -mwindows)
endif()

set_target_properties(${PROJECT_NAME} PROPERTIES
    OUTPUT_NAME "FileRenamer"
    WIN32_EXECUTABLE TRUE
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
)

install(TARGETS ${PROJECT_NAME}
    RUNTIME DESTINATION bin
)

include(CTest)
if(BUILD_TESTING)
    add_executable(RenamerTests tests/RenamerTests.cpp src/RenamerService.cpp src/RenameSuggestions.cpp)
    target_include_directories(RenamerTests PRIVATE src)
    target_compile_definitions(RenamerTests PRIVATE UNICODE _UNICODE NOMINMAX)
    target_link_libraries(RenamerTests PRIVATE shlwapi ole32)
    if(MSVC)
        target_compile_options(RenamerTests PRIVATE /W4 /permissive- /utf-8)
    endif()
    add_test(NAME RenamerTests COMMAND RenamerTests)
    add_executable(RenamerUiTests tests/RenamerUiTests.cpp)
    target_compile_definitions(RenamerUiTests PRIVATE UNICODE _UNICODE NOMINMAX)
    target_link_libraries(RenamerUiTests PRIVATE user32 gdi32)
    if(MSVC)
        target_compile_options(RenamerUiTests PRIVATE /W4 /permissive- /utf-8)
    endif()
    add_dependencies(RenamerUiTests FileRenamer)
    add_test(NAME RenamerUiTests COMMAND RenamerUiTests $<TARGET_FILE:FileRenamer>)
    set_tests_properties(RenamerUiTests PROPERTIES TIMEOUT 30)
endif()
