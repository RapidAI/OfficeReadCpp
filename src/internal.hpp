#pragma once

#include "officeread/officeread.hpp"

#include <cstdint>
#include <map>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace officeread::detail {

using Bytes = std::vector<std::byte>;

std::string clean_text(std::string_view input);
std::string join_text(const std::vector<std::string>& parts);
std::string xml_unescape(std::string_view input);
std::string normalize_part_path(std::string_view source, std::string_view target);
std::string image_extension(std::span<const std::byte> data);
std::string sanitize_filename(std::string_view name);
std::string url_escape_path(std::string_view value);

Result extract_ooxml(const std::filesystem::path& filename,
                     std::span<const std::byte> data, const Options& options,
                     unsigned depth = 0);
Result extract_legacy(const std::filesystem::path& filename,
                      std::span<const std::byte> data, const Options& options,
                      unsigned depth = 0);

}  // namespace officeread::detail
