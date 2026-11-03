# CSV 转换（中文版）

基于 [furyutei/ConvertCsvToTable](https://github.com/furyutei/ConvertCsvToTable)，保留原 MIT 许可证。

中文版将功能区菜单、运行状态、用户提示以及安装/卸载提示改成中文。标签采用 Unicode 码位生成，不依赖 VBA 工程的日文或中文代码页。

读取文件时优先严格识别 UTF-8（带 BOM 或不带 BOM），支持 BOM 标记的 UTF-16，其他常见中文 CSV 按 GBK 读取。原有表格转换逻辑保持不变。编码已在上游损坏的内容无法仅靠切换编码恢复。

## 使用

安装后，Excel 功能区会出现“CSV 转换”。本版本默认开启自动转换，以后双击 CSV 时会重新读取并创建转换后的表格工作表。可通过“停止自动转换”和“开启自动转换”切换。

原 CSV 文件不会自动覆写。要保留新增工作表和格式，请另存为 `.xlsx`。

## 安装和卸载

关闭 Excel 后，运行 `addin/Install.vbs`。插件文件为 `ConvertCsvToTableZh.xlam`，安装到当前用户的 `%APPDATA%\Microsoft\Addins`，不会覆盖原版文件名。

若系统未启用 VBScript，可在 Excel“文件 → 选项 → 加载项 → 管理：Excel 加载项 → 转到 → 浏览”中加载该文件。

关闭 Excel 后运行 `addin/Uninstall.vbs` 可卸载中文版。也可以在 Excel 加载项列表中取消勾选。

## 源码

中文版发布在本仓库的 `main` 分支。完整克隆后，运行 `python localize.py` 可从固定的上游提交 `ea21096348933bca24f8162f71bac3f54ac9fafb` 重新生成汉化源码。修改过的源码文件使用 UTF-8；`build/*.code.txt` 为供 VBA 编辑器使用的 ASCII 代码，避免工程代码页造成乱码。

可公开的验证记录见 [`docs/verification.json`](docs/verification.json)。本机安装记录保存在被 Git 忽略的 `build` 目录。

2026-11-03 已在本机 64 位 Microsoft Excel 16.0（Build 20527）验证：UTF-8 无 BOM、UTF-8 BOM、GBK、UTF-16 BOM 四组文件通过中文、逗号、引号、单元格换行、金额及原文件不变检查；安装后的系统默认打开流程亦通过。没有修改全局宏安全设置或 VBA 工程访问设置。

汉化后的开发工作簿是 `src/ConvertCsvToTableZh.xlsm`，可在 VBA 编辑器中查看、编译。最终插件是 `addin/ConvertCsvToTableZh.xlam`。重新编译时修复了原计时辅助类用 `Set` 返回布尔值的编译错误。
