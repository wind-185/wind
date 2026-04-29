# GitHub macOS 打包说明

这个压缩包用于上传到 GitHub，让 GitHub Actions 在 macOS 环境里打包生成真正可在 Mac 上运行的应用。

## 使用步骤

1. 在 GitHub 新建一个仓库。
2. 把本压缩包解压后的所有文件上传到仓库根目录。
3. 打开仓库的 `Actions` 页面。
4. 选择 `Build macOS App`。
5. 点 `Run workflow` 手动运行，或者推送到 `main/master` 后自动运行。
6. 构建完成后，在 workflow 结果页面下载 artifact：`魔方原声处理-macOS`。

## Mac 打开提示

如果 Mac 提示“无法验证开发者”，可以在 Finder 里右键应用，选择“打开”，再确认打开。
