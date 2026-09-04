# Invoice Mail Downloader Skill

一个支持 `SKILL.md` 的 AI Agent Skill：从 163/QQ 邮箱下载并归档电子发票，自动维护 `发票登记.xlsx`，并使用 TextIn XParse 免费模式识别、匹配和归档报销凭证。

## 安装

运行环境为 **macOS 或 Windows**（不支持 Linux）。需要 Python 3.10+。

### Codex

macOS：

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

macOS：

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
- 仅 macOS 或 Windows，不支持 Linux；
- 先在邮箱网页端开启 IMAP 并生成客户端授权码；
- 优先使用一次性本地配置页面，普通用户不需要操作终端；
- 授权码只保存到 macOS 钥匙串或 Windows 凭据管理器；
- 每个新账号首次运行时，由用户选择邮件日期范围或从现在开始只增量获取；
- 同一票号同时提供 PDF/OFD 时仅保留 PDF，ZIP 内也应用相同规则；
- 普通增量运行可回补具有完整来源的已删除文件；指定日期运行严格只处理该日期查询返回的邮件，不混入旧待处理队列；
- 识别常见 `dzfp_票号` 文件名，过滤邮件地址和 DOC/DOCX 说明链接，避免冗余 OFD及无效重试；
- 支持通过诺诺/JSS、百望预览页使用的只读官方接口解析 PDF，同时继续要求最终文件域名逐个获得信任；
- 可信发票平台返回的 PDF 即使文字识别失败也会保留到 `待确认/`，不会误判完成后丢弃；
- 下载完成后自动维护发票登记表，按票号或文件哈希去重，并保留支付日期、收款姓名、报销金额、费用类型和备注等人工字段；
- 自动填写发票日期、发票金额、开票方、发票编号、下载时间和校验状态；费用类型提供“部门营销费用、企业文化费用、出差报销费用”下拉选项；
- 自动创建 `报销凭证待匹配/` 投放目录；其中的 PDF/OFD 先用本地 Python 识别并按发票规则归档。提取失败或无法识别的 OFD 放入 `待确认/`；有文字但不是发票的 PDF 再匹配报销凭证；
- 登记表增加凭证是否存在、相对路径、匹配状态、校验说明和“发票与实付差额”，并在人工支付日期为空时从凭证补充；员工“收款姓名”等人工字段始终保留；
- 支持微信“扫二维码付款-给…”、转账时间和转账单号；实付低于发票金额不超过 1.00 元（含）时按支付优惠匹配并在登记表标注差额，实付高于发票金额仍拒绝自动匹配；
- 缺少固定收款方标签时，会用 OCR 全文中的完整开票方名称兜底；明确收款方冲突或名称只出现在付款方、付款账户、备注等上下文时仍拒绝自动匹配；
- 修改归档目录和增加可信下载域名都需要用户明确确认。

首次升级后可在不连接邮箱的情况下补录现有归档：

```bash
python3 "/absolute/path/to/invoice-mail-downloader/scripts/run_skill.py" sync-excel
```

Windows 将 `python3` 替换为 `py -3`。Excel 优先使用 Codex 随附的 `@oai/artifact-tool`；若不可用，则使用 Python 依赖中的 `openpyxl`。`preflight` 的 `ok` 只检查核心运行环境，Excel 与 XParse 作为可选能力单独报告。

报销凭证通过 TextIn XParse 公有服务识别，文件会上传至 TextIn；敏感凭证请勿放入投放目录。安装凭证识别能力：

```bash
python3 "/absolute/path/to/invoice-mail-downloader/scripts/bootstrap.py" --with-receipts
```

该步骤需要 Node.js 18+ 与 npm，会把官方 `xparse-cli` 装到隔离的应用数据目录。随后可单独运行：

```bash
python3 "/absolute/path/to/invoice-mail-downloader/scripts/run_skill.py" match-receipts
```

技能只使用免登录免费模式，不会自动切换付费 API；同一文件默认复用本地精简缓存，不重复消耗额度。

安装隔离依赖：

```bash
python3 "/absolute/path/to/invoice-mail-downloader/scripts/bootstrap.py"
```

初始化后，Agent 可启动本地安全配置页面：

```bash
python3 "/absolute/path/to/invoice-mail-downloader/scripts/run_skill.py" configure-ui --provider qq --email user@qq.com
```

Windows 将 `python3` 替换为 `py -3`。授权码只提交到本机 `127.0.0.1`，在保存前会先验证 TLS 和 IMAP 登录。

首次运行二选一：

```bash
# 指定邮件接收日期范围
python3 "/absolute/path/to/invoice-mail-downloader/scripts/run_skill.py" run --from-date 2026-08-01 --to-date 2026-08-31 --confirm-first-run

# 不处理历史邮件，从现在开始建立增量游标
python3 "/absolute/path/to/invoice-mail-downloader/scripts/run_skill.py" run --start-from-now --confirm-first-run
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

当前包含离线回归测试。真实邮箱连接和 TextIn 公有服务需要在本机集成验证。
