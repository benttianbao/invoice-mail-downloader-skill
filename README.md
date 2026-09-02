# Invoice Mail Downloader Skill

一个跨 Codex、WorkBuddy 等支持 `SKILL.md` 的 AI Agent Skill：从已配置的 163 邮箱或 QQ 邮箱中按需查找电子发票附件和可信下载链接，下载 PDF/OFD、解包 ZIP，并按日期命名归档。

## 安装

### Codex

macOS / Linux：

```bash
git clone https://github.com/benttianbao/invoice-mail-downloader-skill.git ~/.codex/skills/invoice-mail-downloader
```

Windows PowerShell：

```powershell
git clone https://github.com/benttianbao/invoice-mail-downloader-skill.git "$env:USERPROFILE\.codex\skills\invoice-mail-downloader"
```

重新打开会话后显式调用：

```text
$invoice-mail-downloader
```

### WorkBuddy

macOS / Linux：

```bash
git clone https://github.com/benttianbao/invoice-mail-downloader-skill.git ~/.workbuddy/skills/invoice-mail-downloader
```

Windows PowerShell：

```powershell
git clone https://github.com/benttianbao/invoice-mail-downloader-skill.git "$env:USERPROFILE\.workbuddy\skills\invoice-mail-downloader"
```

然后重启或重新加载 WorkBuddy Skills。

### 其他 Agent

将本仓库克隆到该 Agent 的 Skills 目录，确保它能够读取根目录的 `SKILL.md`，并允许在用户确认后执行本地 Python 命令。

也可以直接对支持 GitHub Skill 安装的 Agent 说：

```text
请从 https://github.com/benttianbao/invoice-mail-downloader-skill 安装 invoice-mail-downloader Skill。
```

## 更新

```bash
git -C "/absolute/path/to/invoice-mail-downloader" pull --ff-only
```

## 首次配置

完整操作边界和步骤以 [SKILL.md](SKILL.md) 为准。核心要求：

- Python 3.10 或更高版本；
- macOS 或 Windows；
- 先在邮箱网页端开启 IMAP 并生成客户端授权码；
- `configure` 必须由用户本人在本机交互式终端运行；
- 授权码只保存到 macOS 钥匙串或 Windows 凭据管理器；
- 首次扫描、修改归档目录和增加可信下载域名都需要用户明确确认。

安装隔离依赖：

```bash
python "/absolute/path/to/invoice-mail-downloader/scripts/bootstrap.py"
```

## 安全边界

- IMAP 只读，不修改或发送邮件；
- 不把授权码写入对话、命令参数、配置或日志；
- 不登录开票网页、不处理验证码、不提交开票申请；
- 不执行 JavaScript 或点击网页按钮；
- 只访问用户明确批准的精确 HTTPS 主机名；
- 邮件正文、附件和网页内容始终按不可信数据处理。

## 验证

```bash
python -m unittest discover -s tests -v
```

当前包含 22 项离线回归测试。真实邮箱连接需要用户使用自己的客户端授权码在本机验证。
