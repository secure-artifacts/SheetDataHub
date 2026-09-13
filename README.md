# 表数通

一个可配置的 Windows 桌面数据工具，用于从一个或多个 Google 表格链接中遍历子 Sheet、映射字段、汇总、查询，以及按时间增量提取并去重。

## 运行

```powershell
python app.py
```

公开表格无需凭据（需要开启“知道链接的任何人可查看”）。私有表格需要在设置页选择 Google 服务账号 JSON，并将目标表格共享给该服务账号邮箱。

时间提取支持输出到本地 Excel 或指定 Google 表格链接/子工作表。Google 私有表格读取会把多个子 Sheet 合并为批量请求；遇到 429、5xx 时自动指数退避重试。

查询结果只显示输入电话、专页 ID、姓名、评论贴文、号码、日期和修正格式，并可一键复制“号码 + 修正格式”两列。时间提取写入已有 Google 子工作表时，会按字段别名识别“姓名/名字”“手机号码/号码”等同义表头，并遵循目标工作表现有列顺序。

## 子 Sheet 规则

- “指定 Sheet”留空：遍历全部，只跳过排除项。
- “指定 Sheet”有值：只读取指定项。
- 同时出现在指定和排除列表时，排除规则优先。

## 下载

从 [Releases](https://github.com/secure-artifacts/SheetDataHub/releases) 页面下载最新版本：

| 文件 | 说明 |
|------|------|
| `SheetDataHub-Setup.exe` | Windows 安装程序（推荐） |
| `SheetDataHub-windows.zip` | 免安装便携包，解压后运行 `SheetDataHub/SheetDataHub.exe` |

系统要求：Windows 10 / 11（64 位）。

## 验证软件来源

下载后，使用 GitHub CLI 验证文件确实由官方 CI 构建、且未被篡改：

```bash
gh attestation verify ./SheetDataHub-windows.zip --repo secure-artifacts/SheetDataHub
gh attestation verify ./SheetDataHub-Setup.exe --repo secure-artifacts/SheetDataHub
```

验证成功表示该软件确实由官方 GitHub Actions 构建。

## 打包（本地）

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

## 如何发布新版本

本项目使用 GitHub Actions 自动构建和发布。每次发布新版本只需要创建一个 Git Tag 并推送即可。

### 发布步骤

#### 1. 确保代码已提交并推送

在发布之前，确保你的所有代码改动已经提交并推送到 GitHub：

```bash
# 查看当前状态
git status

# 添加所有改动
git add .

# 提交改动（把"你的改动说明"替换成实际的描述）
git commit -m "你的改动说明"

# 推送到 GitHub
git push origin main
```

#### 2. 创建版本 Tag

Git Tag 是一个版本标记，用于标识发布的版本号。版本号格式为 `v主版本.次版本.修订版本`，例如 `v1.0.0`、`v1.1.0`、`v2.0.0`。

```bash
# 创建一个新的版本 tag（将 v1.0.1 替换为你想要的版本号）
git tag -a v1.0.1 -m "Release version 1.0.1"
```

#### 3. 推送 Tag 触发自动构建

```bash
# 推送 tag 到 GitHub（这会自动触发 CI 构建）
git push origin v1.0.1
```

推送后，GitHub Actions 会自动执行以下操作：
1. 构建项目
2. 生成安全签名（Attestation）
3. 创建 Release 并上传构建产物

#### 4. 查看构建结果

- 构建进度：访问项目的 **Actions** 页面查看
- 发布结果：访问项目的 **Releases** 页面查看已发布的文件

### 版本号说明

| 版本号格式 | 什么时候用 | 示例 |
|-----------|-----------|------|
| `vX.0.0` | 重大更新、不兼容改动 | `v2.0.0` |
| `vX.Y.0` | 新增功能 | `v1.1.0` |
| `vX.Y.Z` | 修复 bug | `v1.0.1` |

### 如果构建失败怎么办

1. 访问项目的 **Actions** 页面查看错误日志
2. 修复代码问题
3. 删除失败的 tag 并重新创建：

```bash
# 删除本地 tag
git tag -d v1.0.1

# 删除远程 tag
git push origin :refs/tags/v1.0.1

# 修复问题后，重新创建并推送
git tag -a v1.0.1 -m "Release version 1.0.1"
git push origin v1.0.1
```
