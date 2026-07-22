#include "internal.hpp"

#include <algorithm>
#include <array>
#include <codecvt>
#include <cstdint>
#include <cstring>
#include <locale>
#include <stdexcept>
#include <unordered_set>

namespace officeread::detail {
namespace {

constexpr std::uint32_t free_sector = 0xffffffffu;
constexpr std::uint32_t end_chain = 0xfffffffeu;

std::uint16_t u16(std::span<const std::byte> b, std::size_t p) {
  if (p + 2 > b.size()) return 0;
  return std::to_integer<std::uint8_t>(b[p]) | (std::to_integer<std::uint8_t>(b[p + 1]) << 8);
}
std::uint32_t u32(std::span<const std::byte> b, std::size_t p) {
  return std::uint32_t(u16(b, p)) | (std::uint32_t(u16(b, p + 2)) << 16);
}
std::uint64_t u64(std::span<const std::byte> b, std::size_t p) {
  return std::uint64_t(u32(b, p)) | (std::uint64_t(u32(b, p + 4)) << 32);
}

struct Entry { std::string name; std::uint8_t type{}; std::uint32_t start{}; std::uint64_t size{}; };

std::string utf16_name(std::span<const std::byte> bytes, std::size_t chars) {
  std::u16string value;
  for (std::size_t i = 0; i < chars && i * 2 + 1 < bytes.size(); ++i) value.push_back(static_cast<char16_t>(u16(bytes, i * 2)));
  try { return std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t>{}.to_bytes(value); } catch (...) { return {}; }
}

class Ole {
 public:
  explicit Ole(std::span<const std::byte> data) : data_(data) {
    static constexpr std::array<unsigned char, 8> sig{0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1};
    if (data.size() < 512 || !std::equal(sig.begin(), sig.end(), reinterpret_cast<const unsigned char*>(data.data()))) throw std::runtime_error("not an OLE compound file");
    sector_size_ = std::size_t{1} << u16(data, 30); mini_sector_size_ = std::size_t{1} << u16(data, 32);
    mini_cutoff_ = u32(data, 56); dir_start_ = u32(data, 48);
    const auto fat_count = u32(data, 44);
    std::vector<std::uint32_t> difat;
    for (std::size_t p = 76; p + 4 <= 512 && difat.size() < fat_count; p += 4) { auto s = u32(data, p); if (s != free_sector) difat.push_back(s); }
    auto next_difat = u32(data, 68); auto difat_count = u32(data, 72);
    for (std::uint32_t n = 0; n < difat_count && next_difat < 0xfffffffa; ++n) {
      auto sec = sector(next_difat); if (sec.empty()) break;
      for (std::size_t p = 0; p + 4 < sec.size() && difat.size() < fat_count; p += 4) { auto s = u32(sec, p); if (s != free_sector) difat.push_back(s); }
      next_difat = u32(sec, sec.size() - 4);
    }
    for (auto s : difat) { auto sec = sector(s); for (std::size_t p = 0; p + 4 <= sec.size(); p += 4) fat_.push_back(u32(sec, p)); }
    parse_directory(); parse_minifat();
  }

  const std::vector<Entry>& entries() const { return entries_; }
  Bytes stream(const Entry& e) const {
    if (e.size < mini_cutoff_ && e.type == 2 && !mini_stream_.empty()) return mini_chain(e.start, static_cast<std::size_t>(e.size));
    return chain(e.start, static_cast<std::size_t>(e.size));
  }

 private:
  std::span<const std::byte> sector(std::uint32_t id) const {
    const std::size_t pos = 512 + static_cast<std::size_t>(id) * sector_size_;
    if (pos >= data_.size()) return {};
    return data_.subspan(pos, std::min(sector_size_, data_.size() - pos));
  }
  Bytes chain(std::uint32_t start, std::size_t wanted = SIZE_MAX) const {
    Bytes out; std::unordered_set<std::uint32_t> seen;
    for (auto id = start; id < 0xfffffffa && id < fat_.size() && seen.insert(id).second;) {
      auto sec = sector(id); out.insert(out.end(), sec.begin(), sec.end()); id = fat_[id];
      if (out.size() >= wanted) break;
    }
    if (out.size() > wanted) out.resize(wanted);
    return out;
  }
  Bytes mini_chain(std::uint32_t start, std::size_t wanted) const {
    Bytes out; std::unordered_set<std::uint32_t> seen;
    for (auto id = start; id < 0xfffffffa && id < minifat_.size() && seen.insert(id).second;) {
      const auto pos = static_cast<std::size_t>(id) * mini_sector_size_;
      if (pos >= mini_stream_.size()) break;
      const auto n = std::min(mini_sector_size_, mini_stream_.size() - pos);
      out.insert(out.end(), mini_stream_.begin() + static_cast<std::ptrdiff_t>(pos), mini_stream_.begin() + static_cast<std::ptrdiff_t>(pos + n));
      id = minifat_[id]; if (out.size() >= wanted) break;
    }
    if (out.size() > wanted) out.resize(wanted);
    return out;
  }
  void parse_directory() {
    auto bytes = chain(dir_start_);
    for (std::size_t p = 0; p + 128 <= bytes.size(); p += 128) {
      const auto len = u16(bytes, p + 64); const auto type = std::to_integer<std::uint8_t>(bytes[p + 66]);
      if (!type || len < 2 || len > 64) continue;
      entries_.push_back({utf16_name(std::span(bytes).subspan(p, 64), len / 2 - 1), type, u32(bytes, p + 116), u64(bytes, p + 120)});
    }
    if (!entries_.empty()) mini_stream_ = chain(entries_.front().start, static_cast<std::size_t>(entries_.front().size));
  }
  void parse_minifat() {
    const auto start = u32(data_, 60), count = u32(data_, 64); auto bytes = chain(start, static_cast<std::size_t>(count) * sector_size_);
    for (std::size_t p = 0; p + 4 <= bytes.size(); p += 4) minifat_.push_back(u32(bytes, p));
  }
  std::span<const std::byte> data_; std::size_t sector_size_{512}, mini_sector_size_{64}; std::uint32_t mini_cutoff_{4096}, dir_start_{};
  std::vector<std::uint32_t> fat_, minifat_; std::vector<Entry> entries_; Bytes mini_stream_;
};

std::vector<std::string> strings(std::span<const std::byte> data) {
  std::vector<std::string> out;
  std::string ascii;
  auto flush = [&] { if (ascii.size() >= 4) out.push_back(ascii); ascii.clear(); };
  for (auto b : data) { auto c = std::to_integer<unsigned char>(b); if ((c >= 32 && c < 127) || c >= 160) ascii += static_cast<char>(c); else flush(); } flush();
  std::u16string wide;
  auto flush_wide = [&] {
    if (wide.size() >= 2) try { out.push_back(std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t>{}.to_bytes(wide)); } catch (...) {}
    wide.clear();
  };
  for (std::size_t p = 0; p + 1 < data.size(); p += 2) { auto c = u16(data, p); if (c == 9 || c == 10 || c == 13 || c >= 32) wide.push_back(static_cast<char16_t>(c)); else flush_wide(); } flush_wide();
  return out;
}

bool content_stream(std::string name) {
  std::transform(name.begin(), name.end(), name.begin(), ::tolower);
  return name == "worddocument" || name == "0table" || name == "1table" || name == "powerpoint document" ||
         name == "current user" || name == "workbook" || name == "book" || name.find("contents") != std::string::npos;
}

}  // namespace

Result extract_legacy(const std::filesystem::path& filename, std::span<const std::byte> data,
                      const Options& options, unsigned) {
  Ole ole(data); Result result; std::vector<std::string> parts; std::size_t image_no = 0;
  for (const auto& entry : ole.entries()) {
    if (entry.type != 2 || entry.size == 0 || entry.size > (256ull << 20)) continue;
    auto bytes = ole.stream(entry);
    const auto ext = image_extension(bytes);
    if (!ext.empty()) { Image image; image.name = entry.name.empty() ? "image-" + std::to_string(++image_no) : entry.name; image.ext = ext; image.data = std::move(bytes); result.images.push_back(std::move(image)); continue; }
    if (content_stream(entry.name) || options.include_metadata) { auto found = strings(bytes); parts.insert(parts.end(), found.begin(), found.end()); }
  }
  if (parts.empty()) parts = strings(data);
  result.text = join_text(parts); result.structured_markdown = result.text;
  if (filename.extension() == ".xls" && !result.text.empty()) result.structured_markdown = "## Workbook\n\n" + result.text;
  return result;
}

}  // namespace officeread::detail
