# Tailki 本地邮件工作流

用于在专用 Mac 上准备达人联络表，并逐步搭建网页邮箱收发和回信入表流程。

远端仓库：https://github.com/xiaodengwang01/ittalk

## 目前已完成

- 模拟表生成程序及日常 Git 上传入口。
- Personal Access Token 安全输入与 macOS 钥匙串授权入口。

邮箱尚未接入，邮件收发和定时检查尚未启用。生成的模拟表中发送许可全部为“否”。

## 放置真实材料

当前“`KOL搜集与邮件回复`”目录中的达人跟进表、邮件模板和参考文档保留在本机，该目录不上传。也可以使用“待导入”“备份”“发送预览”“运行记录”存放后续材料，这四个目录只允许上传空目录占位文件。

本仓库为公开仓库，仅纳入程序、说明和明确标注的模拟表。`Tailki_模拟.xlsx` 是可选的，当前工作目录已移除该示例；如需重新生成，见下文。真实表格不要使用这个文件名；只要该文件存在，上传前就会检查是否仍为模拟数据。

## 日常更新

已约定由 Codex 在每次完成本项目的修改和必要检查后提交并上传。也可以双击“一键上传.command”，或执行：

```sh
python3 scripts/git_sync.py "说明本次修改"
```

该入口会检查目标仓库、模拟数据和待提交文件，创建本地提交并上传到 `origin/main`。无新改动时不创建空提交，但仍尝试上传已有提交。认证失败或远端分支发生冲突时停止，并保留本地提交，不强制覆盖远端。

## 首次用 Personal Access Token 授权

1. 在 GitHub 创建 Fine-grained personal access token，Resource owner 选择 `xiaodengwang01`，Repository access 选择 Only select repositories → `ittalk`，Repository permissions 中设置 Contents → Read and write，并设置适当的有效期。Metadata 的只读权限会自动附带。
2. 双击“首次授权并上传.command”，在 macOS 安全输入框中粘贴 Token，点击“授权并上传”。只需要输入 Token，仓库地址和用户名已配好。
3. 程序将 Token 存入本机 macOS 钥匙串，调用 `scripts/git_sync.py` 提交并上传，随后核对远端提交。成功后会弹窗提示，以后直接使用“一键上传.command”。

授权仅修改本仓库的 Git 配置，钥匙串记录按仓库路径区分。Token 不写入项目文件、上传日志或命令参数，也不要粘贴到聊天中。Token 到期后重新运行首次授权入口即可。取消输入不会上传；认证失败或远端冲突时保留本地提交，不强制推送。

GitHub 官方说明：https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens

## 重新生成模拟表

`scripts/create_mock_workbook.mjs` 使用 Codex 桌面应用提供的表格运行库。使用该运行库配套的 Node 执行，并指定一个尚不存在的输出文件路径：

```sh
node scripts/create_mock_workbook.mjs /tmp/Tailki_模拟_重新生成.xlsx
```

生成程序会拒绝覆盖现有目标文件。只有明确标注的模拟表可以作为 `Tailki_模拟.xlsx` 纳入仓库。
