cmake_minimum_required(VERSION 3.20)
project(officeread VERSION 0.1.0 LANGUAGES CXX C)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

include(FetchContent)
set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
set(MINIZ_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(MINIZ_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
FetchContent_Declare(miniz
  GIT_REPOSITORY https://github.com/richgel999/miniz.git
  GIT_TAG 3.1.0)
FetchContent_MakeAvailable(miniz)

add_library(officeread
  src/officeread.cpp
  src/ooxml.cpp
  src/ole.cpp
  src/text.cpp)
target_include_directories(officeread
  PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
    $<INSTALL_INTERFACE:include>
  PRIVATE src ${miniz_SOURCE_DIR})
target_link_libraries(officeread PRIVATE $<BUILD_INTERFACE:miniz>)
if(MSVC)
  target_compile_options(officeread PRIVATE /W4 /utf-8)
else()
  target_compile_options(officeread PRIVATE -Wall -Wextra -Wpedantic)
endif()

add_executable(officeread_cli cmd/officeread.cpp)
set_target_properties(officeread_cli PROPERTIES OUTPUT_NAME officeread)
target_link_libraries(officeread_cli PRIVATE officeread)

enable_testing()
add_executable(officeread_tests tests/tests.cpp)
target_link_libraries(officeread_tests PRIVATE officeread)
add_test(NAME officeread_tests COMMAND officeread_tests)

install(TARGETS officeread officeread_cli
  EXPORT officereadTargets
  ARCHIVE DESTINATION lib
  LIBRARY DESTINATION lib
  RUNTIME DESTINATION bin)
install(DIRECTORY include/ DESTINATION include)
install(EXPORT officereadTargets
  FILE officereadTargets.cmake
  NAMESPACE officeread::
  DESTINATION lib/cmake/officeread)
