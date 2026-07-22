# OfficeReadCpp

OfficeReadCpp 是一个 C++20 Office 文档内容提取库，同时提供命令行工具。项目直接读取 OOXML ZIP 包与 OLE Compound File，不依赖 Microsoft Office、LibreOffice、COM 自动化或外部文档转换器。

它可从 Word、PowerPoint 和 Excel 文档中提取文本、Markdown 与嵌入图片，适合文档预览、归档、搜索索引以及 AI/RAG 预处理。

> 本项目是 `OfficeRead_new` Go 项目的 C++ 移植版。目前提供可构建、可运行的功能性移植基线；原项目中针对罕见损坏文件和特殊 Office 行为的全部边界用例，仍需持续做差分兼容。

## 功能

- 支持 `.docx`、`.pptx`、`.xlsx`、`.doc`、`.ppt`、`.xls`。
- 提取清洗后的用户可见文本。
- 为 OOXML 和 legacy Excel 内容生成 Markdown。
- 提取包内媒体图片，并生成稳定、安全的输出文件名。
- 可选包含文档元数据。
- 提供严格 Office 内容与图片语义开关。
- 同时提供静态 C++ 库和 `officeread` 命令行程序。

## 格式支持

| 格式 | 容器/解析方式 | 当前提取内容 |
| --- | --- | --- |
| DOCX | OOXML ZIP/XML | 正文、页眉页脚、脚注、尾注、批注、AltChunk HTML、媒体图片 |
| PPTX | OOXML ZIP/XML | 幻灯片、备注、母版/布局文本、图表文本、媒体图片 |
| XLSX | OOXML ZIP/XML | 工作表名称、shared/inline strings、单元格、Markdown 表格、批注/图表、媒体图片 |
| DOC | OLE CFB | WordDocument/Table 流的文本恢复、可识别图片 |
| PPT | OLE CFB | PowerPoint Document 流的文本恢复、可识别图片 |
| XLS | OLE CFB | Workbook/Book 流的文本恢复、简化 Markdown、可识别图片 |

## 环境要求

- 支持 C++20 的编译器：MSVC、GCC 或 Clang。
- CMake 3.20 或更高版本。
- Ninja、Visual Studio 或其他 CMake 生成器。
- 首次配置时需要网络连接；CMake 会通过 `FetchContent` 获取 [miniz](https://github.com/richgel999/miniz)。

## 构建与测试

使用 Ninja：

```powershell
cmake -S . -B build -G Ninja -DCMAKE_BUILD_TYPE=Release
cmake --build build --parallel
ctest --test-dir build --output-on-failure
```

使用 Visual Studio：

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Release
ctest --test-dir build -C Release --output-on-failure
```

安装到指定目录：

```powershell
cmake --install build --prefix install
```

## 命令行

基本用法：

```text
officeread [options] file
```

| 参数 | 说明 |
| --- | --- |
| `-images DIR` | 将提取的图片写入 `DIR` |
| `-metadata` | 包含文档属性等元数据 |
| `-markdown` | 输出 Markdown；未指定图片目录时默认使用 `images` |
| `-text-only` | 不在标准错误中输出图片数量 |
| `-strict-office-images` | 启用严格 Office 图片语义选项 |
| `-strict-office-content` | 限制 OOXML 文本为主要文档内容 |
| `-h`, `--help` | 显示帮助 |

示例：

```powershell
build\officeread.exe sample.docx
build\officeread.exe -markdown -images images sample.xlsx > sample.md
build\officeread.exe -metadata sample.pptx
build\officeread.exe -strict-office-content sample.docx
```

文本写入标准输出；诊断信息和图片数量写入标准错误。成功时退出码为 `0`，参数错误为 `2`，读取或解析失败为 `1`。

## C++ API 快速入门

```cpp
#include <officeread/officeread.hpp>

#include <iostream>

int main() {
  officeread::Options options;
  options.image_dir = "images";
  options.include_metadata = false;

  try {
    const auto result = officeread::extract("sample.docx", options);
    std::cout << result.text << '\n';
    std::cout << result.markdown("images") << '\n';
    std::cout << "images: " << result.images.size() << '\n';
  } catch (const std::exception& error) {
    std::cerr << error.what() << '\n';
    return 1;
  }
}
```

完整的类型、字段、异常和使用约定见 [API 文档](docs/API.md)。

## 作为 CMake 子项目使用

```cmake
add_subdirectory(path/to/OfficeReadCpp)
target_link_libraries(your_target PRIVATE officeread)
```

然后在代码中包含：

```cpp
#include <officeread/officeread.hpp>
```

## 差分验证

仓库提供抽样脚本，可将 C++ 结果与本机的 Go 原项目比较：

```powershell
powershell -ExecutionPolicy Bypass -File tools/differential.ps1 `
  -Source D:\workprj\OfficeRead_new `
  -PerFormat 10
```

脚本覆盖六种格式，并在 Go 输出非空而 C++ 输出为空时返回失败。

## 项目结构

```text
include/officeread/   公开 C++ 头文件
src/                  OOXML、OLE、文本和图片实现
cmd/                  officeread 命令行程序
tests/                CTest 测试
tools/                差分验证脚本
docs/                 API 文档
```

## 限制与安全说明

- 不解密密码保护的 Office 文档。
- 不还原完整 Office 对象模型、布局、样式、公式计算或宏行为。
- Legacy DOC/PPT/XLS 采用 OLE 流和安全字符串恢复；复杂二进制记录的语义覆盖低于 OOXML。
- `strict_office_images` 已保留在兼容 API 中；当前实现并非对所有格式都能完整复刻 Microsoft Office Picture Shape 计数。
- XML、ZIP、OLE 流设置了大小与链遍历约束，但仍应把不可信文件放在资源受限的进程中处理。
- 图片文件会写入 `Options::image_dir`；调用者应选择有权限且不会覆盖重要文件的目录。

## 许可证

仓库当前未包含许可证文件。在明确许可证前，请勿假设拥有分发或再授权权利。
