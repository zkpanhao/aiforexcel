# AI for Excel：Office.js 加载项

这是 AI for Excel 的 Office.js 实现。加载项在 Excel 任务窗格中运行，通过用户配置的 OpenAI 兼容接口完成分析和操作规划，再由本地执行器调用 Excel API 修改工作簿。

## 功能

支持的能力包括：

- 读取选区并进行只读分析
- 写入文字、数字和公式，清除内容或格式
- 设置字体、填充、边框、对齐方式、数字格式和行列尺寸
- 排序、筛选、合并与拆分单元格
- 条件格式、图表和透视汇总
- 跨工作表复制值、公式、格式和可见单元格
- 按映射表批量替换名称
- 大范围数据分批处理、取消和失败恢复
- 需要依赖顺序的任务分步执行、暂停、继续和撤销

详细的操作边界见 [`docs/功能清单.md`](../docs/功能清单.md) 和 [`docs/operation-inventory.md`](../docs/operation-inventory.md)。

## 运行环境

- Windows 或 macOS 上的 Microsoft 365 Excel
- ExcelApi 1.13 或更高版本
- Node.js 22.18+ 或 Node.js 24
- 一个可用的 OpenAI 兼容模型接口

## 安装依赖

在本目录执行：

```powershell
npm ci
```

## 本地运行

```powershell
npm start
```

启动脚本会启动本地 HTTPS 服务并尝试将 `manifest.xml` 侧载到桌面 Excel。首次运行可能需要安装并信任 Office 开发证书。也可以使用 `npm run dev` 只启动开发服务器。

运行类型检查、测试、构建和清单校验：

```powershell
npm run typecheck
npm test
npm run build
npm run validate
```

## 直接使用安装版

如果你只想使用加载项，请下载项目压缩包，不需要修改源代码。使用流程如下：

1. 解压项目压缩包到本机目录。
2. 双击仓库根目录的 `install.cmd`，等待安装完成。
3. 重新打开 Excel，在“开始”选项卡打开 **AI for Excel**。
4. 打开任务窗格的“连接设置”，填写自己的 Base URL、模型和 API Key。
5. 在新建测试工作簿中选择少量单元格，发送一条只读请求验证连接。

`install.cmd` 会自动准备 Node.js 24、安装依赖并执行构建和注册，不需要用户直接运行 Node.js 命令。

开发者也可以在仓库根目录运行：

```powershell
.\manage-local.ps1 -Action install -OutsideApp
```

该命令会检查源代码、安装本地运行环境、构建加载项并注册本机清单。它面向 Windows 开发环境，首次运行需要 Node.js 24 和相应的 Excel/证书权限。

## 连接设置

项目没有公开的默认中转地址。打开任务窗格中的“连接设置”，选择内置供应商，或填写自己的 Base URL 和模型名称，然后输入 API Key。

连接设置保存到本机安装目录的 `user-data/connection.json`。只有勾选“记住密钥（仅此设备）”时才会保存 API Key；未勾选时，密钥仅保留在当前窗格内存中。`user-data/` 已加入 Git 忽略规则，不会随项目同步。

支持的内置供应商包括 OpenAI、DeepSeek、Qwen、GLM 和 MiniMAX。选择“自定义 / 中转”即可使用其他 OpenAI 兼容接口。

## 数据与隐私

加载项只在用户发送请求后读取所需的工作簿选区。首轮请求通常只包含工作簿、工作表、地址和范围大小；模型明确需要内容时，才会读取并发送相应数据。规则化的文本清理和部分跨表映射操作在本机完成。

对话摘要只保存在当前任务窗格内存中。连接设置、密钥、工作簿、日志和构建产物均属于本地数据，不应提交到 Git 仓库。

## 目录结构

| 目录 | 说明 |
| --- | --- |
| `src/taskpane` | 任务窗格界面和交互 |
| `src/llm` | 模型接口和响应处理 |
| `src/excel` | 选区读取、Excel 操作和恢复 |
| `src/state` | 连接、对话和任务状态 |
| `src/tools` | 操作协议和参数校验 |
| `tests` | 自动化测试 |

## 开发说明

加载项使用白名单操作协议执行模型返回的计划，不执行模型生成的 VBA、JavaScript 或其他脚本。涉及工作簿修改的功能应在专用测试工作簿中验证，并确认恢复操作可用。

项目当前没有发布包或云端部署配置。贡献代码前请先运行完整的类型检查、自动化测试、生产构建和 manifest 校验。
