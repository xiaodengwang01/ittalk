# Tailki

日常工作：**查看未读回信 → 分析并生成带教审核 DOCX → 根据修改定稿 → 确认发送 → 核验并更新达人跟进表。** 每轮当次处理，没有独立待办分类或日报。

## 日常资料

```text
KOL搜集与邮件回复/
├── 达人跟进表.xlsx
├── Tailki_KOL全流程图文合集.docx
└── 参考资料/
    ├── tailki-kol-workflow/       唯一技能本体及现有配图
    ├── 历史文档/                原 DOCX、完成后的审核稿
    └── 本轮审核 DOCX、原计划表和备份
```

新版技能为“Tailki 达人建联与回信”，入口 `参考资料/tailki-kol-workflow/SKILL.md`，本机技能目录通过链接发现。日常读取技能、合集对应场景和该达人最新往来及有效决定。合集包含三类原首联和 32 个回信场景，点击场景索引跳转。所有个案金额、日期及未核实产品功能都需按本轮核对。

原有 27 份 DOCX 包括副本原样归档。原路径、归档路径、SHA-256 和字节数保存在原本机数据库的 `daily_state/document-archive:` 记录中；程序通过 `local_resources.resolve` 读取原件。历史批准、批次来源和发送证据不改写，校验不一致时停止读取。历史文档只用于追溯或吸收新批注，不作为日常逐份必读材料。

## 处理回信

1. 用 `scripts/mail_inbox.py` 只读未读正文并保持未读状态。`daily_workflow.py sync-records/context 达人ID` 复用已采集单封往来，不补扫已读正文。
2. 根据技能判断阶段和本轮目标，核对签名、已有号码、已答内容、报价及当前有效决定；查合集对应章节生成完整图文回复。
3. 每位达人固定四部分：对方来信原文、来信要点、我方回复要点及应对策略、我方完整回复正文。一级标题按达人分节，带稳定达人 ID；正文区域明确“收件人：”“主题：”，支持抄送/密送和实际内嵌图片。
4. 审核稿直接放“参考资料”。首轮命名 `YYYY-MM-DD_达人回信审核.docx`，同日其他批次加 `_02` 等序号，同轮修改复用原文件。每轮实际拟稿必须有 DOCX，没有回复稿不生成空文档。
5. 读取带教正文、修订和带锚点批注，处理后更新完整正文和梗概。通用意见沉淀进技能和合集；个案条件留原记录。新 DOCX 按 documents 技能渲染并逐页检查；文档哈希改变后重新检查。
6. 最终完整版本获得明确发送确认后走原发送链路。核验实际发送后更新达人跟进表，完成后审核稿原样归档。文稿认可和归档均不等于发送。

实际命令在 `/Users/ludwig/Desktop/Tailki` 执行，使用带 `python-docx` 的 Codex 内置 Python。已整理的审核稿可导入原本机记录，再按所选达人生成或复用本轮文件：

```bash
python3 scripts/review_center.py adopt-report /本机/审核稿.docx
python3 scripts/daily_workflow.py review --docx --creators 达人ID1 达人ID2
python3 scripts/daily_workflow.py review --docx --round 已返回的轮次编号
```

轮次按达人集合与最新来信依据识别；来信变化须重新核对，不让格式修改重新认可旧稿。无有效完整稿时不生成文档；指定达人中有缺稿、未处理批注或图片缺失时停止本轮导出。未登记的同名文件保留，新批次自动递增序号。

导入后更新表格使用 `adopt-report 路径 --updates 本机更新.json`，更新内容须绑定当前 `source_hash`、`book_hash`，每位达人提供新待发梗概及获准字段；带批注须逐条处理并标记 `comments_reviewed`，新来信须提供匹配的 `message_basis`。现有程序验证版本后窄范围写回，不自动猜测正文语义。

单纯 `daily_workflow.py review`、`prepare/summarize/record-sent` 和发送收尾只刷新原 `~/Library/Application Support/Tailki/review/今日审核.html`。双击 `scripts/今日审核.command` 可打开完整预览；页面操作本身不收信、不发信、不批准、不改表，也不额外产生日报。

发送收尾核实整轮全部对应邮件、正文/图片和写表/草稿清理完成后，自动将审核稿归档。取消拟稿时可在用户明确指令范围内运行 `review_center.py archive-report 轮次编号 --cancelled-instruction "取消原话"`；仅归档并保留未发送事实。归档失败不掩盖已经完成的发送或写表结果。

## 跟进表和内部记录

跟进表保持 2026-11-05 已确认布局：主表 11 个可见字段为达人ID、达人昵称、新增批次、账号链接、首联内容分类、内容类型、粉丝数、当前建联状况、沟通摘要、Email、往来邮件梗概。副表“长期合作意愿”已删除“待发送回复梗概”列，其余布局保留。两表按达人 ID 定位，一人只在一张表维护，不自动迁移或恢复删列。

“新增批次”保留历史日期，最早一批为“2026-10-16之前”，不等于发送日期。待发梗概、缺列联系方式及原研究资料继续存现有本机记录；原计划表和 `达人跟进表_2026-11-05备份.xlsx` 保留。表格只保存单封邮件日期与梗概，全文留文档及原本机记录。

唯一运行库仍为 `~/Library/Application Support/Tailki/mail-drafts.sqlite3`。`mail_records` 按账号、方向和 Message-ID 去重；引用不增轮次，覆盖不全不能推定没有历史。草稿、批准、发送尝试及回执继续沿用；`mail_actions` 分别记录写表与草稿清理的恢复进度。新增 `reply-report:` 仅绑定文档轮次、依据及校验值，不是第二套达人台账。

## 搜集与首联

搜集在当前授权任务范围内执行，每日最多新增 5 位，含补充批次；历史已完成批次不重写。不因旧未发送稿阻断新的准备，但每位入选前实时查重。

```bash
python3 scripts/daily_workflow.py audit-history
python3 scripts/daily_workflow.py begin YYYY-MM-DD
python3 scripts/candidate_pool.py research --limit 10
python3 scripts/candidate_pool.py ingest /本机/当天核实结果.json
python3 scripts/candidate_pool.py select YYYY-MM-DD
```

先查看候选池，再实际核实公开邮箱、英语市场、内容适配与近期表现。TikTok 按主页最新顺序看近期非置顶封面及播放量，优先 10 条；不逐条打开、不编造具体剧情或发布时间。`ingest` 记录真实 `checked_at/routes_checked` 与来源；不把旧日期改成今天冒充核对。缺资料、旧线索、共享邮箱歧义留池，合格不足如实记录。

`select` 查两表、候选账号与邮箱历史、实际收件/已发送/草稿，入选后冻结完整首联预览，不换人或追加。中断恢复原预留 ID。独立补充须有当天指令且合计仍最多 5 位；`start-today` 不改变定时状态。

三类首联从合集 `TAILKI_FIRST_PET_STORY`、`TAILKI_FIRST_PET_PRODUCT`、`TAILKI_FIRST_GENERAL` 书签范围提取，只含主题和正文。缺书签、重复、交叉或不完整直接停止；历史冻结批次继续读取旧原件，不改写正文和依据。新准备检查跟进表、合集和新版技能三个文件。

## 发送、核验和恢复

现有阿里 IMAP/SMTP 链路保持。密码只在本机输入并存 macOS 钥匙串。连接检查、保存草稿、定时准备、带教认可和商务初步同意均不授予发送权限。

`mail_send.py review --batches 批次 --expect-count 人数 --output 本机预览.html` 读回邮箱最新完整稿，包括实际收件人、抄送、密送、正文、排版及附件。向用户展示所有最终版本，取得针对该版本的明确发送指令后才 `approve`、`send`。内容变化重新确认，网页人工修改不覆盖，被删草稿不重建，结果不明不重发。

SMTP 接受后，核对服务端唯一已发送副本、完整 MIME 和实际时间；`reconcile` 只读查证。`finalize` 不投递邮件，重新核验后登记事实、按表头安全写回、读验，再只将对应批准版本的草稿移入已删除邮件。写表冲突保留人工值；移动结果不明核对两边，不重复复制、不全局 EXPUNGE、不清空回收站。`finalize-pending` 只恢复新流程已核实发送的收尾，不追溯重发旧批次。

现有自动化任务本次保留 **暂停** 状态及原时间配置，仅替换资料和流程说明；不会因整理文件被启用。定时授权与每批发送确认独立。

## 检查与同步

```bash
python3 -m unittest discover -s scripts -p 'test_*.py'
```

检查仅使用 SIM 编号、`.invalid` 邮箱及模拟资料，不连接真实邮箱、不投递。`workflow.py` 的旧整表重写入口不适用于当前跟进表。

日常收尾不自动提交、同步或上传。仅用户明确要求本次上传时运行：

```bash
python3 scripts/git_sync.py --scope tailki "更新 Tailki 整理后的流程、合集和技能"
```

按 2026-11-06 确认的范围，远程只上传全流程合集、新版技能与三张配图、项目说明、必要程序和模拟检查。原 27 份 DOCX 及后续归档、每轮审核稿、真实达人跟进表、计划表及备份、邮件记录、登录资料和运行数据库均保留本机。唯一允许上传的表格是使用 SIM 编号和 `.invalid` 邮箱的 `scripts/examples/Tailki_模拟.xlsx`。

上传程序与 `.gitignore` 使用精确文件清单，不把整个业务目录加入 Git，不夹带 docs、Easel 等已有改动。合集、配图和模拟表使用已审版本校验；修改后核对内容再更新校验值。上传前检查源文件及 DOCX/XLSX 内部 XML，本机 `scripts/gitleaks` 检查工具本身不上传。旧 docs/Easel 上传入口仅保留兼容性，Tailki 必须指定 `--scope tailki`。

仓库为 `https://github.com/xiaodengwang01/ittalk.git`，不强制推送，上传失败说明原因并保留本地提交。新克隆只含通用方法和程序，真实业务资料需从本机保留目录使用；不得从模拟表推断真实发送权限。
