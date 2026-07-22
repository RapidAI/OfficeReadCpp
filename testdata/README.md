# Test data

本目录从 `OfficeRead_new/testdata` 复制，用于后续将 Go 回归测试逐项迁移到 C++。

- `samples/`：六种 Office 格式的核心正常与回归样本。
- `negative/`：损坏、截断或异常 Office 输入。
- `web-samples-smoke/`：小规模外部样本冒烟集合及清单。
- `reference/`：原 Go 测试源码和 Office baseline 报告，仅供迁移断言参考。

未复制 `web-samples/`：该完整外部语料约 5.8 GiB、7,348 个文件，不适合直接纳入普通 Git 仓库。需要全量兼容验证时，继续使用原目录：

```powershell
tools\differential.ps1 -Source D:\workprj\OfficeRead_new -PerFormat 10
```

测试样本可能来自不同上游项目，仅应用于兼容性测试。对外重新分发前，应逐项复核其来源和许可。
