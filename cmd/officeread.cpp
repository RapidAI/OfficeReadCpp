#include "officeread/officeread.hpp"

#include <filesystem>
#include <iostream>
#include <string>

int main(int argc, char** argv) {
  officeread::Options options; bool markdown = false, text_only = false; std::filesystem::path input;
  for (int i = 1; i < argc; ++i) {
    std::string arg = argv[i];
    if (arg == "-images" && i + 1 < argc) options.image_dir = argv[++i];
    else if (arg == "-metadata") options.include_metadata = true;
    else if (arg == "-strict-office-images") options.strict_office_images = true;
    else if (arg == "-strict-office-content") options.strict_office_content = true;
    else if (arg == "-markdown") markdown = true;
    else if (arg == "-text-only") text_only = true;
    else if (arg == "-h" || arg == "--help") {
      std::cout << "usage: officeread [options] file\n"
                   "  -images DIR\n"
                   "  -metadata\n"
                   "  -markdown\n"
                   "  -text-only\n"
                   "  -strict-office-images\n"
                   "  -strict-office-content\n";
      return 0;
    }
    else if (!arg.empty() && arg.front() == '-') { std::cerr << "unknown option: " << arg << '\n'; return 2; }
    else input = std::filesystem::path(arg);
  }
  if (input.empty()) { std::cerr << "usage: officeread [options] file\n"; return 2; }
  if (markdown && options.image_dir.empty()) options.image_dir = "images";
  try {
    const auto result = officeread::extract(input, options);
    std::cout << (markdown ? result.markdown(options.image_dir.generic_string()) : result.text);
    if (!result.text.empty()) std::cout << '\n';
    if (!text_only) std::cerr << "images: " << result.images.size() << '\n';
  } catch (const std::exception& e) { std::cerr << "officeread: " << e.what() << '\n'; return 1; }
}
