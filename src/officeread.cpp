#include "internal.hpp"

#include <algorithm>
#include <fstream>
#include <stdexcept>
#include <unordered_map>

namespace fs = std::filesystem;

namespace officeread {
namespace {

detail::Bytes read_file(const fs::path& path) {
  std::ifstream in(path, std::ios::binary | std::ios::ate);
  if (!in) throw std::runtime_error("cannot open file: " + path.string());
  const auto end = in.tellg();
  if (end < 0) throw std::runtime_error("cannot determine file size: " + path.string());
  detail::Bytes out(static_cast<std::size_t>(end));
  in.seekg(0);
  in.read(reinterpret_cast<char*>(out.data()), static_cast<std::streamsize>(out.size()));
  if (!in && !out.empty()) throw std::runtime_error("cannot read file: " + path.string());
  return out;
}

bool is_zip(std::span<const std::byte> b) {
  return b.size() >= 4 && b[0] == std::byte{'P'} && b[1] == std::byte{'K'} &&
         (b[2] == std::byte{3} || b[2] == std::byte{5} || b[2] == std::byte{7}) &&
         (b[3] == std::byte{4} || b[3] == std::byte{6} || b[3] == std::byte{8});
}

void finalize_images(Result& result, const Options& options) {
  std::unordered_map<std::string, unsigned> used;
  for (std::size_t i = 0; i < result.images.size(); ++i) {
    auto& image = result.images[i];
    if (image.ext.empty()) image.ext = detail::image_extension(image.data);
    std::string stem = image.name.empty() ? "image-" + std::to_string(i + 1) : image.name;
    stem = fs::path(stem).filename().string();
    if (!image.ext.empty() && fs::path(stem).extension().empty()) stem += image.ext;
    stem = detail::sanitize_filename(stem);
    auto& count = used[stem];
    if (count++) {
      const fs::path p(stem);
      stem = p.stem().string() + "-" + std::to_string(count) + p.extension().string();
    }
    image.name = stem;
  }
  if (!options.image_dir.empty()) {
    fs::create_directories(options.image_dir);
    for (const auto& image : result.images) {
      std::ofstream out(options.image_dir / image.name, std::ios::binary);
      if (!out) throw std::runtime_error("cannot create image: " + (options.image_dir / image.name).string());
      out.write(reinterpret_cast<const char*>(image.data.data()),
                static_cast<std::streamsize>(image.data.size()));
    }
  }
}

}  // namespace

Result extract(const fs::path& filename, const Options& options) {
  const auto bytes = read_file(filename);
  Result result = is_zip(bytes) ? detail::extract_ooxml(filename, bytes, options)
                                : detail::extract_legacy(filename, bytes, options);
  finalize_images(result, options);
  return result;
}

std::string Result::markdown(std::string_view image_base) const {
  std::string out = structured_markdown.empty() ? text : structured_markdown;
  if (!images.empty()) {
    if (!out.empty() && out.back() != '\n') out += '\n';
    out += "\n## Images\n\n";
    for (std::size_t i = 0; i < images.size(); ++i) {
      const auto& image = images[i];
      std::string target;
      if (!image_base.empty()) {
        target = std::string(image_base);
        if (target.back() != '/' && target.back() != '\\') target += '/';
      }
      target += image.name;
      const std::string alt = image.alt.empty() ? "image " + std::to_string(i + 1) : image.alt;
      out += "![" + alt + "](" + detail::url_escape_path(target) + ")\n\n";
    }
  }
  return out;
}

}  // namespace officeread
