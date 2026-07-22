#include "internal.hpp"

#include <algorithm>
#include <cctype>
#include <iomanip>
#include <sstream>
#include <unordered_set>

namespace officeread::detail {
namespace {

void replace_all(std::string& value, std::string_view from, std::string_view to) {
  std::size_t pos = 0;
  while ((pos = value.find(from, pos)) != std::string::npos) {
    value.replace(pos, from.size(), to);
    pos += to.size();
  }
}

std::string trim(std::string_view s) {
  std::size_t first = 0, last = s.size();
  while (first < last && std::isspace(static_cast<unsigned char>(s[first]))) ++first;
  while (last > first && std::isspace(static_cast<unsigned char>(s[last - 1]))) --last;
  return std::string(s.substr(first, last - first));
}

}  // namespace

std::string xml_unescape(std::string_view input) {
  std::string out(input);
  replace_all(out, "&lt;", "<");
  replace_all(out, "&gt;", ">");
  replace_all(out, "&quot;", "\"");
  replace_all(out, "&apos;", "'");
  replace_all(out, "&amp;", "&");
  return out;
}

std::string clean_text(std::string_view input) {
  std::string out;
  out.reserve(input.size());
  bool space = false;
  for (unsigned char c : input) {
    if (c == '\r') continue;
    if (c == '\n') {
      while (!out.empty() && out.back() == ' ') out.pop_back();
      if (!out.empty() && out.back() != '\n') out += '\n';
      space = false;
    } else if (c == '\t' || c == ' ') {
      space = !out.empty() && out.back() != '\n';
    } else if (c >= 0x20 || c >= 0x80) {
      if (space) out += ' ';
      out += static_cast<char>(c);
      space = false;
    }
  }
  return trim(out);
}

std::string join_text(const std::vector<std::string>& parts) {
  std::string out;
  std::unordered_set<std::string> seen;
  for (const auto& raw : parts) {
    auto part = clean_text(raw);
    if (part.empty() || !seen.insert(part).second) continue;
    if (!out.empty()) out += '\n';
    out += part;
  }
  return out;
}

std::string normalize_part_path(std::string_view source, std::string_view target) {
  std::string combined;
  if (!target.empty() && target.front() == '/') combined = std::string(target.substr(1));
  else {
    const auto slash = source.rfind('/');
    combined = slash == std::string_view::npos ? std::string(target)
                                                : std::string(source.substr(0, slash + 1)) + std::string(target);
  }
  replace_all(combined, "\\", "/");
  std::vector<std::string> pieces;
  std::stringstream stream(combined);
  std::string piece;
  while (std::getline(stream, piece, '/')) {
    if (piece.empty() || piece == ".") continue;
    if (piece == "..") { if (!pieces.empty()) pieces.pop_back(); }
    else pieces.push_back(piece);
  }
  std::string out;
  for (const auto& p : pieces) { if (!out.empty()) out += '/'; out += p; }
  return out;
}

std::string image_extension(std::span<const std::byte> b) {
  auto u = [&](std::size_t i) { return std::to_integer<unsigned char>(b[i]); };
  if (b.size() >= 8 && u(0) == 0x89 && u(1) == 'P' && u(2) == 'N' && u(3) == 'G') return ".png";
  if (b.size() >= 3 && u(0) == 0xff && u(1) == 0xd8 && u(2) == 0xff) return ".jpg";
  if (b.size() >= 6 && std::string(reinterpret_cast<const char*>(b.data()), 3) == "GIF") return ".gif";
  if (b.size() >= 2 && u(0) == 'B' && u(1) == 'M') return ".bmp";
  if (b.size() >= 4 && ((u(0) == 'I' && u(1) == 'I') || (u(0) == 'M' && u(1) == 'M'))) return ".tif";
  if (b.size() >= 12 && std::string(reinterpret_cast<const char*>(b.data()), 4) == "RIFF" &&
      std::string(reinterpret_cast<const char*>(b.data() + 8), 4) == "WEBP") return ".webp";
  const std::string head(reinterpret_cast<const char*>(b.data()), std::min<std::size_t>(b.size(), 512));
  if (head.find("<svg") != std::string::npos) return ".svg";
  return {};
}

std::string sanitize_filename(std::string_view name) {
  std::string out;
  for (unsigned char c : name) {
    if (c < 32 || std::string_view("<>:\"/\\|?*").find(static_cast<char>(c)) != std::string_view::npos) out += '_';
    else out += static_cast<char>(c);
  }
  while (!out.empty() && (out.back() == ' ' || out.back() == '.')) out.pop_back();
  if (out.empty()) out = "image";
  return out;
}

std::string url_escape_path(std::string_view value) {
  std::ostringstream out;
  for (unsigned char c : value) {
    if (std::isalnum(c) || std::string_view("-._~/:").find(static_cast<char>(c)) != std::string_view::npos)
      out << static_cast<char>(c);
    else out << '%' << std::uppercase << std::hex << std::setw(2) << std::setfill('0') << static_cast<int>(c)
             << std::nouppercase << std::dec;
  }
  return out.str();
}

}  // namespace officeread::detail
