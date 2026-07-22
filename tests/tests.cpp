#include "officeread/officeread.hpp"

#include <array>
#include <cassert>
#include <filesystem>
#include <iostream>
#include <stdexcept>
#include <string_view>

#ifndef OFFICEREAD_TESTDATA_DIR
#define OFFICEREAD_TESTDATA_DIR "testdata"
#endif

namespace {

void test_markdown() {
  officeread::Result r;
  r.text = "hello";
  r.images.push_back({"a b.png", "picture", ".png", {}});
  const auto md = r.markdown("images");
  assert(md.find("hello") != std::string::npos);
  assert(md.find("a%20b.png") != std::string::npos);
}

void test_corpus_smoke() {
  namespace fs = std::filesystem;
  const fs::path samples = fs::path(OFFICEREAD_TESTDATA_DIR) / "samples";
  if (!fs::is_directory(samples)) {
    throw std::runtime_error("test corpus not found: " + samples.string());
  }

  constexpr std::array<std::string_view, 6> formats{
      ".docx", ".pptx", ".xlsx", ".doc", ".ppt", ".xls"};
  for (const auto format : formats) {
    fs::path selected;
    for (const auto& entry : fs::directory_iterator(samples)) {
      if (!entry.is_regular_file() || entry.path().extension() != format ||
          entry.file_size() <= 100) {
        continue;
      }
      if (selected.empty() || entry.file_size() < fs::file_size(selected)) {
        selected = entry.path();
      }
    }
    if (selected.empty()) {
      throw std::runtime_error("missing test sample for " + std::string(format));
    }
    const auto result = officeread::extract(selected);
    if (result.text.empty() && result.images.empty()) {
      throw std::runtime_error("empty extraction for " + selected.string());
    }
  }
}

}  // namespace

int main() {
  test_markdown();
  test_corpus_smoke();
  std::cout << "ok\n";
}
