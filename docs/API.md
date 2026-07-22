# OfficeReadCpp API 文档

公开接口只有一个头文件：

```cpp
#include <officeread/officeread.hpp>
```

所有公开符号均位于 `officeread` 命名空间。库要求 C++20。

## `officeread::extract`

```cpp
[[nodiscard]] Result extract(
    const std::filesystem::path& filename,
    const Options& options = {});
```

读取并解析一个 Office 文档。函数根据文件容器自动分派：ZIP 容器进入 OOXML 解析器，其他输入进入 legacy OLE 解析器。

参数：

- `filename`：输入文件路径。
- `options`：提取行为；省略时只提取正文、Markdown 和内存中的图片。

返回 `Result`。无法打开文件、输入容器无效、格式不支持或图片写入失败时抛出 `std::runtime_error` 或其他 `std::exception` 派生异常。库不返回部分结果。

由于返回结果拥有全部文本和图片数据，调用结束后输入文件可以被移动或删除。

## `officeread::Options`

```cpp
struct Options {
  std::filesystem::path image_dir;
  bool include_metadata = false;
  bool strict_office_images = false;
  bool strict_office_content = false;
};
```

### `image_dir`

非空时，`extract` 会创建此目录，并将结果中的有效图片写入该目录。写入后的文件名与 `Result::images[n].name` 一致。

空路径表示只在内存中返回图片，不写磁盘。

### `include_metadata`

为 `true` 时，将 OOXML 的 `docProps` 等元数据文本，以及 legacy OLE 中额外的可恢复流文本纳入结果。默认关闭，使输出更接近用户看到的正文。

### `strict_office_images`

保留与 Go API 对应的严格图片语义开关。当前 C++ 基线会保存该选项，但并非所有解析路径都能完全复刻 Microsoft Office Picture Shape 行为。调用方不应将它当作精确的 Office 图片计数保证。

### `strict_office_content`

为 `true` 时，OOXML 解析跳过部分辅助内容，例如 PPTX 母版、布局、图表缓存以及 XLSX 图表/批注扩展，使文本更接近主要文档内容。

## `officeread::Result`

```cpp
struct Result {
  std::string text;
  std::string structured_markdown;
  std::vector<Image> images;

  [[nodiscard]] std::string markdown(
      std::string_view image_base = {}) const;
};
```

### `text`

UTF-8 编码的清洗文本。内容可能包含换行，但不会保留完整排版和样式。

### `structured_markdown`

解析器生成的结构化 Markdown。XLSX 工作表会尽可能输出 Markdown 表格，PPTX 会按幻灯片生成章节。格式无法结构化时通常退化为 `text`。

此字段不保证包含最终图片引用；如需完整 Markdown，应调用 `markdown()`。

### `images`

文档中的有效嵌入图片。图片按解析顺序存放，名称在返回前完成安全化和去重。

### `markdown(image_base)`

返回完整 Markdown。以 `structured_markdown` 为主体；该字段为空时使用 `text`。若存在图片，则追加 `Images` 章节和 Markdown 图片引用。

`image_base` 是图片引用前缀：

```cpp
result.markdown();
// ![image 1](image-1.png)

result.markdown("images");
// ![image 1](images/image-1.png)

result.markdown("https://cdn.example.com/assets");
// ![image 1](https://cdn.example.com/assets/image-1.png)
```

路径中的空格和特殊字节会进行 URL 百分号编码。此函数不写文件；磁盘图片输出由 `Options::image_dir` 控制。

## `officeread::Image`

```cpp
struct Image {
  std::string name;
  std::string alt;
  std::string ext;
  std::vector<std::byte> data;
};
```

### `name`

安全化并去重后的输出文件名，例如 `image-1.png`。解析器会移除路径目录，替换 Windows 非法文件名字符，并在重名时增加数字后缀。

### `alt`

可用的图片替代文本。为空时，`Result::markdown()` 使用 `image N` 作为回退文本。

### `ext`

带点号的小写或源扩展名，例如 `.png`、`.jpg`。当包内名称缺少扩展名时，解析器会依据图片魔数推断。

### `data`

图片的完整二进制数据。类型为 `std::vector<std::byte>`，由 `Result` 独占。

## 完整示例

```cpp
#include <officeread/officeread.hpp>

#include <fstream>
#include <iostream>

int main(int argc, char** argv) {
  if (argc != 2) {
    std::cerr << "usage: example FILE\n";
    return 2;
  }

  officeread::Options options;
  options.image_dir = "output/images";
  options.strict_office_content = true;

  try {
    const officeread::Result result = officeread::extract(argv[1], options);

    std::ofstream markdown("output/document.md");
    markdown << result.markdown("images");

    std::cout << "text bytes: " << result.text.size() << '\n';
    std::cout << "images: " << result.images.size() << '\n';
  } catch (const std::exception& error) {
    std::cerr << "extract failed: " << error.what() << '\n';
    return 1;
  }
}
```

对应的 CMake：

```cmake
cmake_minimum_required(VERSION 3.20)
project(example LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
add_subdirectory(path/to/OfficeReadCpp OfficeReadCpp-build)

add_executable(example main.cpp)
target_link_libraries(example PRIVATE officeread)
```

## 线程与生命周期

- `extract` 不依赖全局可变解析状态，使用不同文件和不同 `Options` 的调用可独立执行。
- 同时写入相同 `image_dir` 可能发生文件名竞争；并发调用应使用不同目录，或由调用方进行同步。
- `Result`、`Image` 和返回字符串均为值类型，可移动、复制并在任意线程读取。

## 编码约定

- 文本字段使用 UTF-8。
- `std::filesystem::path` 遵循平台原生路径编码；Windows 下建议直接传入 `std::filesystem::path`，避免手工把本地宽字符路径转为窄字符串。
- 图片 Markdown 链接使用 `/` 分隔并进行 URL 转义。

## 稳定性

当前版本号为 `0.1.0`。公开接口较小，但在达到 `1.0.0` 前仍可能根据兼容性测试调整字段或行为。
