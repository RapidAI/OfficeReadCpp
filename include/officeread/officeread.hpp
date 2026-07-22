#pragma once

#include <cstddef>
#include <filesystem>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace officeread {

/// An embedded image owned by an extraction result.
struct Image {
  /// Safe, deduplicated output filename.
  std::string name;
  /// Alternative text when the source document provides one.
  std::string alt;
  /// File extension including the leading dot, for example ".png".
  std::string ext;
  /// Complete encoded image payload.
  std::vector<std::byte> data;
};

/// Controls extraction and optional image output.
struct Options {
  /// Directory receiving image files. An empty path disables disk output.
  std::filesystem::path image_dir;
  /// Includes document properties and other recoverable metadata text.
  bool include_metadata = false;
  /// Requests Office-compatible picture semantics where supported.
  bool strict_office_images = false;
  /// Excludes auxiliary OOXML content such as cached chart/drawing text.
  bool strict_office_content = false;
};

/// Text, Markdown, and images extracted from one document.
struct Result {
  /// Cleaned UTF-8 visible text.
  std::string text;
  /// Parser-produced Markdown before final image references are appended.
  std::string structured_markdown;
  /// Embedded images in document order.
  std::vector<Image> images;

  /// Renders complete Markdown, optionally prefixing image links with image_base.
  [[nodiscard]] std::string markdown(std::string_view image_base = {}) const;
};

/// Reads and extracts a supported Office document.
///
/// Throws std::exception-derived errors for I/O, unsupported containers, parse
/// failures, or image output failures.
[[nodiscard]] Result extract(const std::filesystem::path& filename,
                             const Options& options = {});

}  // namespace officeread
