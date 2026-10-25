# Visual Studio 2026 打开 OfficeAgent.vdproj 失败处理

> 构建链路总览（code vs installer）见 [README.md](../README.md) 的「从源码构建」一节。
> 打 MSI 前请先跑 `.\build-installer-prep.bat`（Release 代码 + SourcePath 审计）。

## 现象

在 Visual Studio 2026 打开 `AiHelper.sln` 时，`OfficeAgent/OfficeAgent.vdproj` 可能提示：

```text
找不到此项目类型所基于的应用程序
54435603-dbb4-11d2-8724-00a0c9a8b90c
```

这个 GUID 是 Visual Studio Installer Projects 的 `.vdproj` 项目类型。它不是 `WordAi`、`ExcelAi`、`PowerPointAi`、`ShareRibbon` 的 VB.NET/VSTO 版本过低，也不是 `.NET Framework 4.7.2` 本身不兼容。

## 根因

`.vdproj` 不是 Visual Studio 标准内置项目类型，需要当前 Visual Studio 实例安装并启用 `Microsoft Visual Studio Installer Projects` 扩展。

如果 VS 2026 没有安装该扩展、扩展版本过旧、扩展被禁用，或者安装到了另一个 VS 实例，`OfficeAgent.vdproj` 就会显示不兼容。

当前解决方案中的普通代码项目仍然可以正常构建：

- `ShareRibbon/ShareRibbon.vbproj`
- `WordAi/WordAi.vbproj`
- `ExcelAi/ExcelAi.vbproj`
- `PowerPointAi/PowerPointAi.vbproj`

## 处理步骤

1. 在 Visual Studio 2026 打开 `Extensions -> Manage Extensions`。
2. 搜索并安装 `Microsoft Visual Studio Installer Projects`。
3. 确认扩展安装到 Visual Studio 2026 这个实例，而不是旧版 Visual Studio。
4. 关闭所有 Visual Studio 实例。
5. 重新启动 Visual Studio 2026。
6. 重新打开 `AiHelper.sln`。

如果仍然失败：

1. 在 `Extensions -> Manage Extensions -> Installed` 确认扩展已启用。
2. 更新扩展到 3.x 或更高版本。
3. 使用 `devenv.com .\AiHelper.sln /Rebuild Debug` 验证代码项目是否仍可构建。
4. 如果扩展短期不可用，先用 VS 2026 开发和调试代码项目，再用支持该扩展的 Visual Studio 实例构建 MSI。

## 不建议的修复方式

- 不要把 `OfficeAgent.vdproj` 的项目 GUID 改成普通 VB.NET 项目 GUID。
- 不要手工大范围重排 `.vdproj` 内容。
- 不要通过升级 `TargetFrameworkVersion` 来解决这个错误；它和 `.vdproj` 项目类型加载无关。
- 不要为了绕过加载错误删除 `OfficeAgent` 项目，除非明确决定迁移安装包方案。

## 长期建议

`.vdproj` 适合快速生成 MSI，但它依赖 Visual Studio 扩展，自动化能力和版本兼容性都比较弱。后续如果要继续做三合一安装包、减少重复 DLL、支持 CI 构建，建议规划迁移到 WiX Toolset 或 MSIX。

短期内保持 `OfficeAgent.vdproj` 最小改动；中长期把安装包瘦身、共享依赖去重、注册表写入和升级卸载逻辑迁移到可脚本化、可 CI 构建的安装链路。
