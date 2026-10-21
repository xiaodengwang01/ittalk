# Tailki

本机 KOL 搜集、邮件审核和跟进流程。每日准备任务已按用户指令设为北京时间 10:00，包括周末；每天最多新增 8 位合格达人，不补跑旧日期。每批真实邮件仍需展示全部最终稿并取得对应版本的明确确认。

## 日常资料与事实记录

`KOL搜集与邮件回复/` 只在本机保存，保留现有名称和布局：

- `达人跟进表.xlsx`：达人开发主表、长期合作意愿两张表。
- `Tailki_KOL初次建联邮件_三个版本.docx`：三类首联模板。
- `Tailki_KOL回复策略-第二版.docx`：策略依据。
- `Tailki_KOL达人回复参考.docx`：已核实的图文参考。
- `参考资料/`：产品背景及按日期保存的带教汇报。

唯一运行数据库仍为 `~/Library/Application Support/Tailki/mail-drafts.sqlite3`，专用密码仅在 macOS 钥匙串。新邮件事实表 `mail_records` 与既有批次、草稿、发送尝试和回执共用这个库，不新建第二套名单或状态库。邮件按账号、方向和单封 Message-ID 去重；同封邮件的不同文件夹定位合并，原文或归属冲突停止覆盖。引用旧邮件不增加沟通轮次；未读与未处理分别记录。

`mail_actions` 分别保留写表、草稿清理的进度。邮件已成功但写表或清理失败，只恢复这些收尾操作，不能重新发送。

## 每日操作

双击 `scripts/今日审核.command`，会重新整理并打开唯一的“今日审核”页面，同时生成当天带教 DOCX。页面集中展示今日新增、带教待决、全部待审邮件和异常；页面中的“查看带教 DOCX”可直接打开文档，“审核意见”可复制后发回聊天。入口本身不收信、不发送、不批准邮件，也不改 Excel。

实际仓库目录为 `/Users/ludwig/Desktop/Tailki`。命令使用 Python 3.9+；表格读写使用 Codex 提供的 Node 与 Artifact Tool。详细筛选与授权规则见 `docs/daily-maintenance.md`。

1. `python3 scripts/daily_workflow.py status` 查看定时准备、当天进度和待处理来信。
2. `python3 scripts/daily_workflow.py audit-history` 只读核对收件、已发送、草稿的必要邮件头，不读取已读正文或修改标记。
3. `python3 scripts/daily_workflow.py begin YYYY-MM-DD` 复用北京时间当天批次。新准备任务不需要把历史待确认稿发完；未决项仍保留，逐位实时查重。
4. 在已登录的内置浏览器搜集候选，使用 `add 日期 本机输入.json`、`prepare 日期` 入表并生成全部首联预览。缺项或未核实不凑数。
5. `python3 scripts/mail_inbox.py` 只采集收件箱未读正文，保留未读标记。按策略文档形成摘要后，使用 `daily_workflow.py summarize 本机摘要.json` 入队；默认不承诺商务条件或发送回复。
6. `python3 scripts/daily_workflow.py review` 汇总更新统一审核页面及当天带教文档。`prepare/summarize/record-sent` 和发送收尾命令也会自动刷新；生成失败不改变已完成的业务结果，重新打开入口恢复。

页面默认保存到原本机状态目录的 `review/今日审核.html`，复用同一入口；当天 DOCX 位于业务目录 `参考资料/YYYY-MM-DD_KOL跟进.docx`。同输入重复运行复用文档。人工改过的已登记文档先导入修改再更新；遇到未登记的同名文件、结构无法识别或生成期间人工改动，停止覆盖。已发送邮件及其旧预览不会重新列为待发。

已有手工带教文档可以通过 `python3 scripts/review_center.py adopt-report /本机/已有报告.docx` 导入。一级标题按达人分节，段落包含可关联邮箱；回复稿使用“收件人：”“主题：”行，支持明确的抄送/密送和内嵌图片。图片按内容哈希留在同一本机状态目录，不再依赖聊天临时文件。原文档保留。归属不清、缺主题/正文或占位符未填时拒绝导入。

消息与关键表格字段各自绑定审核资料版本；出现新来信、人工改表或图片变化时，停止展示旧完整稿并列为待更新。复制审核意见、导入文档、同意商务方向均不替代完整最终版本的发送确认。带教文档与页面来自同一快照，文档主要收录新增概况、待决背景及已明确要求准备的回复；全部首联全文统一在页面审核。

每次交付新版本 DOCX 前，使用 documents 技能的 `render_docx.py` 渲染并逐页检查。未改动文档复用既有 QA；新哈希不得沿用旧 QA。批量网页稿只含缓存文本而未取得完整 MIME 时明确标为待核对，附件不完整的稿件不冒充最终预览。

`daily_workflow.py sync-records` 将已经合法采集的本机来信、摘要和指定联系人历史邮件头统一关联，不连接邮箱、不改表、不补扫已读正文。
`daily_workflow.py context 达人ID` 查看该达人的本机已记录单封往来及计数。计数注明覆盖范围，不能把未采集的历史当成不存在。

只执行当天时，仍可用 `start-today 日期 核对.json`；补足当天缺口用 `start-supplement 日期 核对.json`。二者不会替换旧确认或重发旧批次，当天人数合计最多 8。

## 邮箱草稿与完整确认

`python3 scripts/mail_drafts.py` 提供 `prepare/setup/audit/save/verify/refresh-source`。按实际表头及三类 DOCX 精确填充；缺字段、分类不匹配、仍有占位符、既有联系或共享邮箱歧义时停止该项。首次保存先读回校验，不覆盖既有稿，不重建已删除稿，保留网页修改。

`setup --gui` 在本机窗口输入账号及三方客户端安全密码，验证 IMAP 并存入钥匙串。IMAP 使用 `imap.qiye.aliyun.com:993`，SMTP 使用 `smtp.qiye.aliyun.com:465`，均校验证书。管理员须开放客户端及 IMAP；参见[阿里官方配置](https://help.aliyun.com/zh/document_detail/36576.html)。

SMTP 操作入口为 `python3 scripts/mail_send.py`：

1. `check-smtp` 仅登录检查，不提交邮件。
2. `review --batches 批次 --expect-count 人数 --output 本机最终预览.html` 读回当前草稿全部收件人、主题、正文、排版和附件，生成完整预览。
3. 展示完整预览并收到用户针对该版本的明确发送指令后，`approve 预览编号 本机确认.json` 记录原话、含时区的指令时间、review_hash 和 count。配置邮箱、启用定时、带教同意议价均不代替发送确认。
4. `send 预览编号` 重新核对确认版和实时历史，登记尝试后逐封投递；任何变化须重新预览确认。`resume` 只允许在此前提交项全部已核实成功时继续原确认范围内完全未提交的余项。
5. SMTP 接受后，从服务器已发送文件夹核对唯一 Message-ID、完整 MIME 和实际发送时间。结果不明暂停；`reconcile 预览编号` 只核对迟到回执，不投递。
6. `send/resume/reconcile` 命令在正常返回时自动调用发送收尾；发生异常后可运行 `finalize 预览编号` 单独恢复。`status 预览编号` 只查询。每日先运行 `finalize-pending` 恢复新流程批准批次中未完成的收尾，不重发，也不自动追溯处理旧批次。

发送入口默认写回当前主表；可在子命令前加 `--book 本机工作簿路径` 指定同结构的授权工作簿。原程序函数 `deliver` 只管理投递，直接集成该函数的调用方须随后调用 `finalize`，包括部分成功后异常的情况。

## 发送成功后的自动收尾

`mail_send.py finalize 预览编号` 本身绝不连接 SMTP：

- 重新核实原确认、已发送唯一副本、收件人、正文/附件及发送时间，登记本机单封事实。
- 按达人 ID 和当前表头定位，只填实际首联日期、必要的首联状态和原备注中的发送记录。已有不同日期或邮箱时不覆盖；长期合作等后续状态不倒退。回复的首轮日期/正文仅在对应字段全部空白时记录，后续轮次保留于本机往来记录和发送备注，不覆盖原轮次或无限扩列。
- 使用 Artifact Tool 生成目标单元格，再保留 XLSX 原部件作窄范围替换；两张表、样式、公式、折叠、批注及其他数据保留。写前备份，文件发生人工修改即停止。
- 写表读验完成后，仅把原批准版本对应草稿移入实际回收站。再次核对 UIDVALIDITY、UID 和完整 MIME；人工修改稿及无关稿保留，已消失稿不重建。
- 先记录移动尝试，再调用 UID MOVE 或 COPY。阿里服务器若 COPY 至回收站后自动移除源稿，直接核验；普通 COPY 仅在支持 UIDPLUS 且回收站副本已核实后按原 UID 清理。绝不全局 EXPUNGE。移动响应不明时重新查证两边，不重复复制；无法安全处理时报告待人工核实。

写表与清理分别记录完成状态；任一步失败都不改变实际发送事实、原确认或发送尝试。重跑收尾不会再次投递。

## 定时与权限

每日准备授权和每批发送授权独立。`configure-schedule 本机授权.json` 记录北京时间 10:00 的 prepare_only 授权并进入 pending_scheduler；通过应用自动化工具启用现有任务后，`confirm-schedule 自动化ID` 记录 active。调度器更新失败则不确认 active。每次定时运行必须先做当天只读历史核对，遇到未决联系逐项跳过，不补跑旧日期。

当前阶段默认回信只采集和摘要；进一步归档入表、拟完整回复、带教汇报依用户具体指令。报价、付款、授权和合同决定单独记录；“长期合作意愿”不表示已达成合同。

## 检查与上传

```bash
python3 -B -m unittest scripts.test_review_center scripts.test_daily_workflow scripts.test_mail_postprocess scripts.test_mail_send scripts.test_mail_inbox scripts.test_mail_drafts scripts.test_mail_text scripts.test_workflow_rules scripts.test_workflow_adapter -q
```

检查使用模拟联系人及 `.invalid` 邮箱，不连接真实 SMTP。真实名单、邮件、凭据、预览、备份和运行记录不上传。

完成程序修改和检查后运行 `python3 scripts/git_sync.py "修改说明"`，提交上传至 `https://github.com/xiaodengwang01/ittalk.git`，不强制推送。`workflow.py` 及其旧表导入/整表重写功能只供历史工具参考，不适用于当前主表。
