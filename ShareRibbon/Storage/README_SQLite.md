# SQLite 包安装说明

记忆功能依赖 System.Data.SQLite 2.0.2。2.0 架构使用 P/Invoke，需要原生库 `e_sqlite3.dll`。

## 路径说明（与 WebView2 一致）

`e_sqlite3.dll` 放在输出目录的 `runtimes\win-x64\native\` 和 `runtimes\win-x86\native\`，与 WebView2Loader.dll 同级。

## 本地开发

1. **首次或缺失时**，在解决方案根目录执行：
   ```powershell
   nuget install SourceGear.sqlite3 -Version 3.50.4.5 -OutputDirectory packages
   ```
2. 构建 WordAi/ExcelAi/PowerPointAi 时，`build\CopySqliteNative.targets` 将 e_sqlite3.dll 作为 **Content** 加入项目（与 WebView2Loader 相同），自动复制到 `bin\Debug\runtimes\`，并被 OfficeAgent 的 **ContentFiles** 输出组收集后打入安装包。
3. 若复制失败（packages 中无 SourceGear），运行时 `SqliteNativeLoader` 会尝试从 packages 拷贝；安装包部署后无 packages，需确保构建时已复制。

## OfficeAgent 打包（用户安装）

OfficeAgent 通过 **ContentFiles** 输出组从 WordAi/ExcelAi/PowerPointAi 收集文件。e_sqlite3.dll 已作为 Content 加入各宿主项目，应会随 ContentFiles 自动包含。打包前请先构建各宿主项目，确保 `bin\Debug\runtimes\` 下已有 e_sqlite3.dll。
