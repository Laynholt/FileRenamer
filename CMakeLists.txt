cmake_minimum_required(VERSION 3.16)

project(FileRenamer
    VERSION 1.0.0
    DESCRIPTION "Pattern-based file renamer for Windows"
    LANGUAGES CXX RC
)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

if(WIN32)
    add_definitions(-DUNICODE -D_UNICODE)
endif()

set(SOURCES
    src/main.cpp
    src/Application.cpp
    src/ExplorerPathProvider.cpp
    src/RenamerService.cpp
    src/Tooltil.cpp
    src/UpdateService.cpp
    src/UiRenderer.cpp
    src/resource.rc
)

set(HEADERS
    src/Application.h
    src/ExplorerPathProvider.h
    src/RenamerService.h
    src/Tooltil.h
    src/UpdateService.h
    src/UiRenderer.h
    src/resource.h
)

add_executable(${PROJECT_NAME} WIN32 ${SOURCES} ${HEADERS})

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
