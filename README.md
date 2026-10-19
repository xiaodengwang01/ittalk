# Tailki

日常文件集中在 `KOL搜集与邮件回复/`：

- `达人跟进表.xlsx`：唯一业务文件，含“达人开发主表”和“邮件记录”。
- `邮件模板.docx`：首联与后续回复模板。
- `参考资料/`：回复策略、三位达人回复、公司手册、页面策划。

复用现有文件，不建立重复主表或未使用的空目录。

主表一行对应一个平台账号，`达人ID`生成后不随排序改变。按平台和规范账号链接去重，不按经纪人邮箱合并。邮箱列只存单一邮箱，联系页另存。来源链接和抓取时间用于追溯，旧表日期、缺年份截图不当作已核实的邮件时间。粉丝数和播放量是数字；疑似中位播放大于最高播放的旧值保留并提示复核。寄样、物流和发布字段已折叠，可展开查看。

“邮件记录”一行对应一封邮件或一封未发送草稿，使用达人ID关联；截图及可辨认文字一并保留，归属冲突的记录标为待人工归属。草稿生成时间、邮件时间分别保存。日期仅有年月日时不补出精确时分；完整邮件时间统一按北京时间入表，原始时间文字另存，未知时区的旧截图保留待核实。

## 本地处理

使用系统 Python 3.9 或更高版本，以及 Codex 自带的 Node 和表格工具，无需安装项目依赖。可用 `TAILKI_NODE` 和 `TAILKI_ARTIFACT_MODULES` 指定对应运行环境。每次写入先在本机临时目录生成，核对字段、备注和截图后替换原文件；处理中若原文件变化则停止覆盖。

```bash
python3 scripts/workflow.py check
python3 scripts/workflow.py refresh
python3 scripts/workflow.py import-scrape /本机路径/抓取结果.json
python3 scripts/workflow.py import-mail /本机路径/邮件结果.json
```

默认操作日常目录中的原表；模拟检查可在命令前加 `--book /本机路径/模拟.xlsx`。`--preview` 会在本机临时目录生成预览，不在业务目录增加文件。

抓取输入使用以下结构。只有提供真实抓取时间和来源链接时才更新抓取时间；已有邮箱、人工分类与候选冲突时保留原值并提示复核。

```json
{
  "抓取时间": "2026-10-19T10:00:00+08:00",
  "creators": [{
    "平台": "TikTok",
    "账号链接": "https://www.tiktok.com/@sim_alpha",
    "达人昵称": "模拟达人 Alpha",
    "Email": "alpha@example.invalid",
    "首联内容分类": "宠物剧情类",
    "粉丝数": "8.5m",
    "来源链接": "https://www.tiktok.com/@sim_alpha"
  }]
}
```

邮件输入为记录数组或 `{"messages": [...]}`。使用表头同名字段：方向、达人ID、邮件时间、发件人、收件人、主题、正文或草稿、邮箱消息ID、会话ID、邮箱账户、原邮件定位。优先使用邮箱原生消息ID去重；`原邮件定位`须精确到单封邮件，只有会话级定位时使用会话ID及完整邮件字段，不能把整个会话链接当作单封邮件标识。无唯一ID/链接且缺少完整时间、双方地址、主题或正文时，不贸然判作重复。共享经纪邮箱、账号或会话冲突进入待人工归属。

```bash
python3 scripts/workflow.py assign-mail MAIL-000001 KOL-000001
python3 scripts/workflow.py draft-reply KOL-000001 "Re: 合作主题" /本机路径/回复正文.txt --template-version "人工回复/v1" --reply-to MAIL-000001
python3 scripts/workflow.py preview-batch MAIL-000002
```

`refresh`从现有“邮件模板.docx”读取首联与未回复跟进模板：资料可用且未联系的账号生成首联草稿；已核实、已获批并成功发送的首联，第7个工作日仍无回复时生成一次跟进草稿，工作日为周一至周五。已回复、暂停联系、邮箱无效、历史发送未经核实或发送结果未知时不催复。回复正文可由Codex结合来信拟写，再通过`draft-reply`入表；同一来信的草稿重复导入不重复建行。

## 确认和发送结果

程序不登录邮箱、不发送邮件；网页抓取、邮箱读取和定时任务尚未连接。抓取或邮箱网页读取得到的结构化结果，通过上述本地入口入表。仅通过已授权的邮箱网页操作。

发送前必须向用户展示对应批次的完整预览并获得明确确认，才可记录确认批次；不能将主表状态或一次确认作为永久许可。收件人、主题或正文变化后确认失效。确认后的实际发送仍由已授权的邮箱网页完成，下面命令只登记确认或观察到的结果：

```bash
python3 scripts/workflow.py confirm-batch 本次确认的批次ID MAIL-000002
python3 scripts/workflow.py record-result MAIL-000002 成功 --at "2026-10-19T10:00:00+08:00"
```

结果可为`成功`、`失败`、`结果待核实`；结果待核实时先查明，不能自动重发。收到回复会作废尚未发送的跟进草稿及其确认。Excel中直接修改已确认草稿后，下次处理会清除失效确认；登记发送结果前也会重新验证内容指纹。

```bash
python3 -m unittest scripts.test_workflow_rules scripts.test_workflow_adapter -v
```

所有自动检查使用模拟数据。旧的 `scripts/examples/Tailki_模拟.xlsx` 仍保留为三页的上传安全示例，不是本次两页业务表的替代文件。

业务目录全部留在本地。Git 只上传程序、说明和 `scripts/examples/Tailki_模拟.xlsx`。

日常由 Codex 在修改后提交并上传，也可双击“一键上传.command”。首次使用 Token 授权时双击“首次授权并上传.command”，只在本机安全输入框填写，凭据保存到 macOS 钥匙串。

远端：https://github.com/xiaodengwang01/ittalk
