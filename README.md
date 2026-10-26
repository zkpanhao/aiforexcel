# AI for Excel

在 Excel 任务窗格中使用自然语言处理表格、分析数据，或拆解复杂任务并分步执行、检查与恢复。

## 项目资料

- [使用与开发说明](officejs/README.md)
- [本地安装、更新及回退](update.md)
- [GitHub 同步与双机协作](docs/collaboration.md)
- [功能清单](docs/功能清单.md)
- [验证记录](docs/testing.md)
- [协作约定](docs/AGENTS.md)

## 工作副本与安装

源代码直接在当前项目目录通过 GitHub 私有仓库同步。本机使用 `D:\OneDrive\learn\coding\AI For Excel`，不另行创建或切换工作副本。

GitHub 只保存源代码、测试、公开清单和项目文档。依赖、构建输出、日志、Excel 工作簿、密钥、证书及本机安装状态不进入仓库。每台电脑分别安装 Node.js 24 和依赖，并按安装说明更新本地插件。

日常可直接让助手“拉取 GitHub 最新版本并更新本机插件”，或“将本次修改提交并推送到 GitHub”。切换电脑前先完成提交和推送；在另一台电脑开始修改前先拉取。推送代码不会自动更新另一台电脑的已安装插件。
