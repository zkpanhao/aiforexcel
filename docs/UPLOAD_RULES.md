# GitHub 上传规则

- 源代码改用 GitHub 私有仓库同步；根目录 `.gitignore` 为统一排除规则。
- 上传源代码、测试、图标、公开 manifest、项目说明和安装脚本。
- 不上传 `node_modules/`、`dist/`、`build/`、`work/`、运行日志、API Key、证书、业务工作簿和本机安装状态。
- `user-data/` 整个目录已加入 `.gitignore`；API 连接配置、记住的密钥、常用命令 JSON 及其备份只留本机，不上传或纳入版本发布包。
- 保留 `package.json` 和 `package-lock.json`，需要时在目标目录运行 `npm ci` 重新生成依赖。
- 项目 `.npmrc` 只保存依赖安装选项；不得写入访问令牌或私有仓库密码。
- 按用户要求，直接使用当前 `learn/coding/AI For Excel` 项目目录同步 GitHub，不另建工作副本。
