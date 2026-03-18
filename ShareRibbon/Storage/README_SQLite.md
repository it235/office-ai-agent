# SQLite 包安装说明

记忆功能依赖 System.Data.SQLite 2.0.2。2.0 架构使用 P/Invoke，需要原生库 `e_sqlite3.dll`。

## 路径说明（与 WebView2 一致）

`e_sqlite3.dll` 放在输出目录的 `runtimes\win-x64\native\` 和 `runtimes\win-x86\native\`，与 WebView2Loader.dll 同级。

## 本地开发

1. **NuGet 包**：项目引用 `SQLitePCLRaw.lib.e_sqlite3.2.1.11` 包，包含原生 DLL。
2. **构建时复制**：
   - ShareRibbon.vbproj 的 `CopySqliteNativeDll` Target 会复制到 `bin\Debug\runtimes\`
   - ExcelAi/WordAi/PowerPointAi 的 Content 项也会复制到各自输出目录
3. 各宿主项目的 Content 配置（无 Condition，确保始终包含）：
   ```xml
   <Content Include="..\packages\SQLitePCLRaw.lib.e_sqlite3.2.1.11\runtimes\win-x64\native\e_sqlite3.dll">
     <Link>runtimes\win-x64\native\e_sqlite3.dll</Link>
     <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
   </Content>
   ```

## OfficeAgent 打包（用户安装）

OfficeAgent 通过 **ContentFiles** 输出组从 WordAi/ExcelAi/PowerPointAi 收集文件。e_sqlite3.dll 已作为 Content 加入各宿主项目，会随 ContentFiles 自动包含。打包前请先构建各宿主项目，确保 `bin\Debug\runtimes\` 下已有 e_sqlite3.dll。
