# Excel CSV 转换（中文版）

在 Windows Excel 中双击打开 CSV 时自动重新读取编码，显示中文，并生成带筛选的表格。默认开启自动转换，安装后可在“CSV 转换”选项卡中暂停或恢复。

本项目基于 [furyutei/ConvertCsvToTable](https://github.com/furyutei/ConvertCsvToTable) 汉化和修改，保留上游代码历史、作者署名与 MIT 许可证。[原版日文说明](README.upstream.ja.md)。

## 下载

前往 [最新发布页](https://github.com/zhangzhangluzi/ConvertCsvToTable-zh/releases/latest)，下载 `ConvertCsvToTable-zh-v1.0.0-zh.1.zip`，其中包含加载项、中文安装/卸载脚本及说明。仅需加载项时可单独下载 `ConvertCsvToTableZh.xlam`。

发布页同时提供 `SHA256SUMS.txt`，可核对文件完整性。

## 安装

1. 解压安装包，保存工作并关闭全部 Excel 窗口。
2. 打开 `addin` 文件夹，双击 `Install.vbs`，按中文提示完成安装。
3. 打开 Excel，确认顶部出现“CSV 转换”选项卡。
4. 直接双击原始 CSV 文件。默认已开启自动转换；如果“开启自动转换”按钮呈灰色、“停止自动转换”可点击，表示正在启用。

插件安装到当前用户的 `%APPDATA%\Microsoft\Addins\ConvertCsvToTableZh.xlam`，使用独立文件名，不覆盖原版加载项。

如果系统无法运行 `.vbs`，可手动加载：Excel“文件 → 选项 → 加载项 → 管理：Excel 加载项 → 转到 → 浏览”，选择 `ConvertCsvToTableZh.xlam` 并勾选。手动加载后请保留文件的位置。

从互联网下载的 Excel 加载项可能被 Office 阻止。确认文件来自本仓库且信任其来源后，可右键 `.xlam` 文件，进入“属性”，如有“解除锁定”选项则勾选，再重新安装或加载。无需关闭全局宏保护；公司策略限制请联系管理员。[微软说明](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)。

## 功能和范围

- 中文功能区、状态提示、用户提示和安装/卸载提示。
- 优先严格识别 UTF-8，包含带 BOM 和不带 BOM 的文件。
- 识别有 BOM 的 UTF-16；其他文件按常见中文 GBK 编码读取。
- 保留上游的表格转换、筛选、文本转数字和文本转超链接功能。

自动转换会新增工作表，磁盘上的原 CSV 不会自动覆写。要保留工作表和格式，请另存为 `.xlsx`。

编码自动判断无法覆盖所有字符集，也无法恢复上游已经丢失的字符。仍保留上游的数值推断行为，长编号和前导零请核对。

## 验证环境

2026-11-03 在 Windows 的 64 位 Microsoft Excel 16.0（Build 20527）完成验证：UTF-8 无 BOM、UTF-8 BOM、GBK、UTF-16 BOM 四组样例中的中文、引号、逗号、单元格换行和金额正确，原文件保持不变。安装后的系统默认打开流程通过验证。

依赖 Windows 桌面版 Excel 的 VBA、ADODB 和 Power Query。其他 Excel 版本未在本次验证中测试。

验证记录见 [`docs/verification.json`](docs/verification.json)。

## 卸载

保存工作并关闭 Excel，运行 `addin/Uninstall.vbs`。也可在 Excel 的加载项列表中取消勾选中文版。

## 源码

详见 [中文版开发说明](README.zh-CN.md)。汉化后的开发工作簿为 `src/ConvertCsvToTableZh.xlsm`，最终加载项为 `addin/ConvertCsvToTableZh.xlam`。

完整克隆后，`python localize.py` 可从固定上游提交重新生成汉化源码。VBA 工程通过 Excel 编译；`python prepare-fixtures.py` 和 `verify-addon.ps1` 可执行验证，`python prepare-release.py` 可准备安装包。

## 许可证

[MIT](LICENSE)。原作者：風柳（furyu / furyutei）。
