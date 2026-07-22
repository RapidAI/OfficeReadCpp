#include "internal.hpp"

#include <miniz.h>

#include <algorithm>
#include <cctype>
#include <filesystem>
#include <regex>
#include <stdexcept>
#include <unordered_map>

namespace officeread::detail {
namespace {

struct Package {
  std::unordered_map<std::string, Bytes> files;
};

Package open_package(std::span<const std::byte> data) {
  mz_zip_archive zip{};
  if (!mz_zip_reader_init_mem(&zip, data.data(), data.size(), 0)) throw std::runtime_error("invalid OOXML ZIP package");
  Package package;
  const auto count = mz_zip_reader_get_num_files(&zip);
  for (mz_uint i = 0; i < count; ++i) {
    mz_zip_archive_file_stat stat{};
    if (!mz_zip_reader_file_stat(&zip, i, &stat) || stat.m_is_directory || stat.m_uncomp_size > (256ull << 20)) continue;
    Bytes bytes(static_cast<std::size_t>(stat.m_uncomp_size));
    if (mz_zip_reader_extract_to_mem(&zip, i, bytes.data(), bytes.size(), 0)) {
      std::string name(stat.m_filename);
      std::replace(name.begin(), name.end(), '\\', '/');
      package.files.emplace(std::move(name), std::move(bytes));
    }
  }
  mz_zip_reader_end(&zip);
  return package;
}

std::string as_string(const Bytes& bytes) {
  return std::string(reinterpret_cast<const char*>(bytes.data()), bytes.size());
}

std::string attr(std::string_view tag, std::string_view local_name) {
  const std::regex re("(?:[A-Za-z0-9_]+:)?" + std::string(local_name) + R"(\s*=\s*["']([^"']*)["'])",
                      std::regex::icase);
  std::cmatch m;
  if (std::regex_search(tag.begin(), tag.end(), m, re)) return xml_unescape(m[1].str());
  return {};
}

std::vector<std::string> xml_visible_text(std::string_view xml) {
  std::vector<std::string> parts;
  std::string current;
  bool hidden = false, deleted = false, instruction = false;
  for (std::size_t pos = 0; pos < xml.size();) {
    const auto open = xml.find('<', pos);
    if (open == std::string_view::npos) break;
    if (open > pos && !hidden && !deleted && !instruction) current += xml_unescape(xml.substr(pos, open - pos));
    const auto close = xml.find('>', open + 1);
    if (close == std::string_view::npos) break;
    auto tag = xml.substr(open, close - open + 1);
    std::string lower(tag);
    std::transform(lower.begin(), lower.end(), lower.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    if (lower.find("<w:del") == 0) deleted = true;
    else if (lower.find("</w:del") == 0) deleted = false;
    else if (lower.find("<w:instrtext") == 0 || lower.find("<instrtext") == 0) instruction = true;
    else if (lower.find("</w:instrtext") == 0 || lower.find("</instrtext") == 0) instruction = false;
    else if (lower.find("<w:vanish") == 0 || lower.find("<a:hidden") == 0) hidden = true;
    else if (lower.find("</w:r") == 0 || lower.find("</a:r") == 0) hidden = false;
    if (lower.find("</w:p") == 0 || lower.find("</a:p") == 0 || lower.find("</row") == 0 ||
        lower.find("</c:tx>") == 0 || lower.find("<w:br") == 0) {
      auto cleaned = clean_text(current);
      if (!cleaned.empty()) parts.push_back(cleaned);
      current.clear();
    } else if (lower.find("<w:tab") == 0 || lower.find("<tab") == 0) current += '\t';
    pos = close + 1;
  }
  auto cleaned = clean_text(current);
  if (!cleaned.empty()) parts.push_back(cleaned);
  return parts;
}

std::vector<std::string> xml_text_elements(std::string_view xml) {
  std::vector<std::string> out;
  const std::string source(xml);
  const std::regex text_re(R"(<(?:[A-Za-z0-9_]+:)?(?:t|text|title|name|description)\b[^>]*>([^<]*)</(?:[A-Za-z0-9_]+:)?(?:t|text|title|name|description)>)",
                           std::regex::icase);
  for (auto i = std::sregex_iterator(source.begin(), source.end(), text_re); i != std::sregex_iterator(); ++i) {
    auto value = clean_text(xml_unescape((*i)[1].str()));
    if (!value.empty()) out.push_back(std::move(value));
  }
  return out;
}

std::vector<std::string> html_visible_text(std::string_view html) {
  std::string source(html);
  source = std::regex_replace(source, std::regex(R"(<(?:script|style)\b[^>]*>[\s\S]*?</(?:script|style)>)", std::regex::icase), " ");
  source = std::regex_replace(source, std::regex(R"(</?(?:p|div|h[1-6]|li|tr|br)\b[^>]*>)", std::regex::icase), "\n");
  source = std::regex_replace(source, std::regex(R"(<[^>]+>)"), " ");
  auto value = clean_text(xml_unescape(source));
  return value.empty() ? std::vector<std::string>{} : std::vector<std::string>{std::move(value)};
}

std::vector<std::string> attribute_values(std::string_view xml, std::string_view name) {
  std::vector<std::string> out; const std::string source(xml);
  const std::regex re("\\b" + std::string(name) + R"(\s*=\s*["']([^"']+)["'])", std::regex::icase);
  for (auto i = std::sregex_iterator(source.begin(), source.end(), re); i != std::sregex_iterator(); ++i) {
    auto value = clean_text(xml_unescape((*i)[1].str())); if (!value.empty()) out.push_back(std::move(value));
  }
  return out;
}

std::vector<std::string> sorted_parts(const Package& package, std::string_view prefix,
                                      std::string_view suffix = ".xml") {
  std::vector<std::string> names;
  for (const auto& [name, data] : package.files)
    if (name.starts_with(prefix) && name.ends_with(suffix)) names.push_back(name);
  std::sort(names.begin(), names.end());
  return names;
}

std::string kind(const Package& package, const std::filesystem::path& filename) {
  if (package.files.contains("word/document.xml")) return "word";
  if (package.files.contains("ppt/presentation.xml")) return "ppt";
  if (package.files.contains("xl/workbook.xml")) return "xl";
  auto ext = filename.extension().string();
  std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
  return ext == ".docx" ? "word" : ext == ".pptx" ? "ppt" : ext == ".xlsx" ? "xl" : "";
}

std::vector<std::string> select_text_parts(const Package& p, const std::string& k, const Options& options) {
  std::vector<std::string> names;
  if (k == "word") {
    for (const auto& [name, _] : p.files) {
      const bool primary = name == "word/document.xml" || name.starts_with("word/header") || name.starts_with("word/footer") ||
                           name == "word/footnotes.xml" || name == "word/endnotes.xml" || name == "word/comments.xml";
      if (primary && name.ends_with(".xml")) names.push_back(name);
    }
  } else if (k == "ppt") {
    auto add = [&](std::string_view prefix) { auto v = sorted_parts(p, prefix); names.insert(names.end(), v.begin(), v.end()); };
    add("ppt/slides/slide"); add("ppt/notesSlides/notesSlide");
    if (!options.strict_office_content) { add("ppt/slideLayouts/slideLayout"); add("ppt/slideMasters/slideMaster"); add("ppt/charts/"); }
  }
  std::sort(names.begin(), names.end());
  return names;
}

std::vector<std::string> shared_strings(const Package& p) {
  const auto it = p.files.find("xl/sharedStrings.xml");
  if (it == p.files.end()) return {};
  std::vector<std::string> out;
  const auto xml = as_string(it->second);
  const std::regex si_re(R"(<(?:\w+:)?si\b[^>]*>([\s\S]*?)</(?:\w+:)?si>)", std::regex::icase);
  for (auto i = std::sregex_iterator(xml.begin(), xml.end(), si_re); i != std::sregex_iterator(); ++i)
    out.push_back(join_text(xml_visible_text((*i)[1].str())));
  return out;
}

std::vector<std::string> xlsx_text(const Package& p, std::string& markdown, const Options& options) {
  std::vector<std::string> parts;
  if (const auto workbook = p.files.find("xl/workbook.xml"); workbook != p.files.end()) {
    auto names = attribute_values(as_string(workbook->second), "name");
    parts.insert(parts.end(), names.begin(), names.end());
  }
  const auto shared = shared_strings(p);
  for (const auto& name : sorted_parts(p, "xl/worksheets/sheet")) {
    const auto xml = as_string(p.files.at(name));
    const std::regex row_re(R"(<(?:\w+:)?row\b[^>]*>([\s\S]*?)</(?:\w+:)?row>)", std::regex::icase);
    bool heading = false;
    for (auto ri = std::sregex_iterator(xml.begin(), xml.end(), row_re); ri != std::sregex_iterator(); ++ri) {
      std::vector<std::string> cells;
      const std::string row = (*ri)[1].str();
      const std::regex cell_re(R"(<(?:\w+:)?c\b([^>]*)>([\s\S]*?)</(?:\w+:)?c>)", std::regex::icase);
      for (auto ci = std::sregex_iterator(row.begin(), row.end(), cell_re); ci != std::sregex_iterator(); ++ci) {
        const auto attrs = (*ci)[1].str(); const auto body = (*ci)[2].str();
        std::string value;
        std::smatch vm;
        const std::regex v_re(R"(<(?:\w+:)?v>([\s\S]*?)</(?:\w+:)?v>)", std::regex::icase);
        if (std::regex_search(body, vm, v_re)) value = xml_unescape(vm[1].str());
        if (attr(attrs, "t") == "s") {
          try { auto n = static_cast<std::size_t>(std::stoull(value)); if (n < shared.size()) value = shared[n]; } catch (...) {}
        } else if (attr(attrs, "t") == "inlineStr" || value.empty()) value = join_text(xml_visible_text(body));
        value = clean_text(value); cells.push_back(value); if (!value.empty()) parts.push_back(value);
      }
      if (!cells.empty()) {
        markdown += '|'; for (const auto& c : cells) markdown += " " + c + " |"; markdown += '\n';
        if (!heading) { markdown += '|'; for (std::size_t i = 0; i < cells.size(); ++i) markdown += " --- |"; markdown += '\n'; heading = true; }
      }
    }
  }
  if (!options.strict_office_content) {
    for (const auto& [name, data] : p.files)
      if ((name.starts_with("xl/comments") || name.starts_with("xl/charts/")) && name.ends_with(".xml")) {
        auto v = xml_visible_text(as_string(data)); parts.insert(parts.end(), v.begin(), v.end());
      }
  }
  if (parts.empty()) {
    for (const auto& [name, data] : p.files)
      if (name.starts_with("xl/") && name.ends_with(".xml")) {
        auto values = xml_text_elements(as_string(data));
        parts.insert(parts.end(), values.begin(), values.end());
      }
  }
  return parts;
}

std::vector<Image> media(const Package& p, const std::string& k) {
  std::vector<std::string> names;
  const std::string prefix = k == "word" ? "word/media/" : k == "ppt" ? "ppt/media/" : "xl/media/";
  for (const auto& [name, _] : p.files) if (name.starts_with(prefix)) names.push_back(name);
  std::sort(names.begin(), names.end());
  std::vector<Image> images;
  for (const auto& name : names) {
    Image image; image.name = std::filesystem::path(name).filename().string(); image.ext = std::filesystem::path(name).extension().string(); image.data = p.files.at(name);
    if (!image_extension(image.data).empty()) images.push_back(std::move(image));
  }
  return images;
}

}  // namespace

Result extract_ooxml(const std::filesystem::path& filename, std::span<const std::byte> data,
                     const Options& options, unsigned) {
  const auto package = open_package(data);
  const auto k = kind(package, filename);
  if (k.empty()) throw std::runtime_error("unsupported ZIP package: " + filename.string());
  Result result;
  std::vector<std::string> parts;
  if (k == "xl") parts = xlsx_text(package, result.structured_markdown, options);
  else for (const auto& name : select_text_parts(package, k, options)) {
    auto values = xml_visible_text(as_string(package.files.at(name)));
    parts.insert(parts.end(), values.begin(), values.end());
    if (k == "ppt" && name.starts_with("ppt/slides/slide") && !values.empty()) {
      result.structured_markdown += "## Slide\n\n" + join_text(values) + "\n\n";
    }
  }
  if (parts.empty()) {
    for (const auto& [name, bytes] : package.files) {
      if (!name.ends_with(".xml") && !name.ends_with(".html") && !name.ends_with(".htm")) continue;
      if (!options.include_metadata && (name.starts_with("docProps/") || name.find("_rels/") != std::string::npos)) continue;
      auto values = xml_text_elements(as_string(bytes));
      parts.insert(parts.end(), values.begin(), values.end());
    }
  }
  for (const auto& [name, bytes] : package.files) {
    if (name.ends_with(".html") || name.ends_with(".htm")) {
      auto values = html_visible_text(as_string(bytes)); parts.insert(parts.end(), values.begin(), values.end());
    }
  }
  if (options.include_metadata) {
    for (const auto& [name, bytes] : package.files)
      if (name.starts_with("docProps/") && name.ends_with(".xml")) { auto values = xml_visible_text(as_string(bytes)); parts.insert(parts.end(), values.begin(), values.end()); }
  }
  result.text = join_text(parts);
  if (result.structured_markdown.empty()) result.structured_markdown = result.text;
  result.images = media(package, k);
  return result;
}

}  // namespace officeread::detail
