# After compilation it will destribute a static lib into a lib folder
# and destributes all header files into an include folder.

cmake_minimum_required(VERSION 2.8)
project(SimpleXlsx)

if(NOT CMAKE_DEBUG_POSTFIX)
  set(CMAKE_DEBUG_POSTFIX d)
endif()

set(UNICODE false CACHE BOOL "Use Unicode macro")

if(${UNICODE})
    add_definitions(-DUNICODE -D_UNICODE)
endif()

file(GLOB MAIN_HDRS
    "*.hpp"
    "*.h"
    )
file(GLOB XLSX_HDRS
    "Xlsx/*.h"
    )
file(GLOB XLSXCOLORS_HDRS
    "XLSXColors/*.h"
    )

file(GLOB ZIP_HDRS
    "Zip/*.h"
    )

file(GLOB SRCS
    "*.cpp"
    "Xlsx/*.cpp"
    "XLSXColors/*.cpp"
    "Zip/*.cpp"
    )

add_library(SimpleXlsx STATIC
            ${SRCS}
            ${MAIN_HDRS}
            ${XLSX_HDRS}
            ${XLSXCOLORS_HDRS}
            ${ZIP_HDRS}
            )
set_target_properties(SimpleXlsx PROPERTIES DEBUG_POSTFIX ${CMAKE_DEBUG_POSTFIX})

install(TARGETS SimpleXlsx DESTINATION lib)
install(FILES ${MAIN_HDRS} DESTINATION include)
install(FILES ${XLSX_HDRS} DESTINATION include/Xlsx)
install(FILES ${XLSXCOLORS_HDRS} DESTINATION include/XLSXColors)
install(FILES ${ZIP_HDRS} DESTINATION include/Zip)