"""Exercise scoped publication against disposable local Git repositories."""
from contextlib import redirect_stdout
from io import StringIO
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

try:
    from . import git_sync
except ImportError:
    import git_sync


IGNORE_RULES = Path(__file__).resolve().parent.parent.joinpath(".gitignore").read_text()


class DocsSyncTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="tailki-docs-sync-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "project with spaces"
        self.remote = Path(self.temp.name) / "remote.git"
        self.root.mkdir()
        self.command("init", "--bare", str(self.remote), cwd=self.root)
        self.command("init", "-b", "main", cwd=self.root)
        self.command("config", "user.email", "test@example.invalid")
        self.command("config", "user.name", "Simulated uploader")
        self.write("scripts/local.py", "local code\n")
        self.write("private.md", "local instructions\n")
        self.write("docs/原文档.md", "original docs\n")
        self.command("add", ".")
        self.command("commit", "-m", "old full-project upload")
        self.command("remote", "add", "origin", str(self.remote))
        self.command("push", "-u", "origin", "main")
        self.write(".gitignore", IGNORE_RULES)
        self.addCleanup(patch.stopall)
        patch.object(git_sync, "ROOT", self.root).start()
        patch.object(git_sync, "REMOTE", str(self.remote)).start()

    def command(self, *args, cwd=None, check=True):
        return subprocess.run(["git", *args], cwd=cwd or self.root,
                              text=True, capture_output=True, check=check,
                              env={**os.environ, "GIT_TERMINAL_PROMPT": "0"}, timeout=15)

    def write(self, name, text):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text)
        return path

    def publish(self, scope="docs"):
        with redirect_stdout(StringIO()):
            git_sync.check_workspace(scope)
            git_sync.sync("Publish " + scope, scope)

    def remote_files(self):
        return git_sync.names(self.command("--git-dir=" + str(self.remote), "ls-tree", "-r", "--name-only", "-z", "main").stdout)

    def add_easel_subtree(self):
        self.publish()
        upstream = Path(self.temp.name) / "official-easel"
        upstream.mkdir()
        self.command("init", "-b", "main", cwd=upstream)
        self.command("config", "user.email", "upstream@example.invalid", cwd=upstream)
        self.command("config", "user.name", "Simulated upstream", cwd=upstream)
        (upstream / "README.md").write_text("upstream one\n")
        (upstream / "LICENSE").write_text("example license\n")
        (upstream / ".gitignore").write_text(".venv/\n.env\noutputs/*\n!outputs/.gitkeep\nprofiles/*\n!profiles/_template/\n")
        (upstream / ".env.example").write_text("API_KEY=REPLACE_ME\n")
        (upstream / "assets/readme").mkdir(parents=True)
        (upstream / "assets/readme/logo.txt").write_text("official asset\n")
        self.command("add", ".", cwd=upstream)
        self.command("commit", "-m", "upstream baseline", cwd=upstream)
        baseline = self.command("rev-parse", "HEAD", cwd=upstream).stdout.strip()
        self.command("remote", "add", git_sync.UPSTREAM_REMOTE, str(upstream))
        self.command("fetch", git_sync.UPSTREAM_REMOTE, "main")
        self.command("update-ref", git_sync.BASELINE_REF, baseline)
        patch.object(git_sync, "UPSTREAM", str(upstream)).start()
        self.command("subtree", "add", "--prefix=Easel", "--squash", str(upstream), baseline)
        return upstream

    def test_subtree_import_and_scoped_updates_preserve_other_area(self):
        self.add_easel_subtree()
        self.write("docs/原文档.md", "private pending docs edit\n")
        self.publish("easel")
        self.assertIn("Easel/README.md", self.remote_files())
        self.assertIn("Easel/.env.example", self.remote_files())
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "show", "main:docs/原文档.md").stdout, "original docs\n")
        self.write("Easel/README.md", "personal code change\n")
        self.publish("docs")
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "show", "main:Easel/README.md").stdout, "upstream one\n")
        self.assertEqual((self.root / "Easel/README.md").read_text(), "personal code change\n")
        self.publish("easel")
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "show", "main:Easel/README.md").stdout, "personal code change\n")

    def test_empty_docs_never_deletes_easel(self):
        self.add_easel_subtree()
        self.publish("easel")
        (self.root / "docs/原文档.md").unlink()
        self.publish("docs")
        self.assertIn("Easel/README.md", self.remote_files())
        self.assertFalse(any(name.startswith("docs/") for name in self.remote_files()))

    def test_private_runtime_and_business_paths_remain_ignored(self):
        self.add_easel_subtree()
        self.publish("easel")
        for name in ("Easel/.env", "Easel/.venv/secret.txt", "Easel/outputs/run.json", "Easel/profiles/private/identity.md", "KOL搜集与邮件回复/private.txt"):
            self.write(name, "private local fixture")
            self.assertEqual(self.command("check-ignore", name, check=False).returncode, 0, name)
        self.publish("easel")
        self.assertFalse(any(name.endswith(("secret.txt", "run.json", "identity.md", "private.txt")) for name in self.remote_files()))

    def test_personal_assets_are_blocked_even_if_force_staged(self):
        self.add_easel_subtree()
        self.publish("easel")
        self.write("Easel/assets/personal.jpg", "private fixture")
        self.command("add", "Easel/assets/personal.jpg")
        tree = self.command("write-tree").stdout.strip()
        with self.assertRaisesRegex(RuntimeError, "不是已核实"):
            git_sync.check_tree(tree)

    def test_changed_example_configuration_cannot_upload_real_values(self):
        self.add_easel_subtree()
        self.publish("easel")
        self.write("Easel/.env.example", "API_KEY=private_value\n")
        with self.assertRaisesRegex(RuntimeError, "示例配置不是已核实"):
            self.publish("easel")
        self.assertEqual(self.command("diff", "--cached", "--name-only").stdout, "")
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "show", "main:Easel/.env.example").stdout, "API_KEY=REPLACE_ME\n")

    def test_cross_scope_unpushed_commits_are_not_uploaded(self):
        self.add_easel_subtree()
        self.publish("easel")
        self.write("docs/原文档.md", "committed docs change")
        self.command("add", "docs/原文档.md")
        self.command("commit", "-m", "unpublished docs")
        with self.assertRaisesRegex(RuntimeError, "以外的改动"):
            self.publish("easel")
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "show", "main:docs/原文档.md").stdout, "original docs\n")

    def test_other_scope_staged_during_scan_is_preserved(self):
        self.add_easel_subtree()
        self.publish("easel")
        self.write("Easel/README.md", "easel change\n")
        self.write("docs/原文档.md", "concurrent docs change\n")
        original_scan = git_sync.scan_secrets

        def concurrent_stage(*, staged=False, base=None):
            original_scan(staged=staged, base=base)
            if staged:
                self.command("add", "docs/原文档.md")

        with patch.object(git_sync, "scan_secrets", side_effect=concurrent_stage):
            with self.assertRaisesRegex(RuntimeError, "暂存内容发生变化"):
                self.publish("easel")
        staged = git_sync.names(self.command("diff", "--cached", "--name-only", "-z").stdout)
        self.assertIn("docs/原文档.md", staged)
        self.assertNotIn("Easel/README.md", staged)
        self.assertEqual((self.root / "Easel/README.md").read_text(), "easel change\n")

    def test_secret_in_source_blocks_upload_without_disclosing_value(self):
        self.add_easel_subtree()
        self.publish("easel")
        secret = "ghp_" + "a1B2c3D4e5F6g7H8i9J0k1L2m3N4o5P6q7R8"
        self.write("Easel/example.py", "token = '" + secret + "'\n")
        with self.assertRaisesRegex(RuntimeError, "密钥检查") as error:
            self.publish("easel")
        self.assertNotIn(secret, str(error.exception))
        self.assertNotIn("Easel/example.py", self.remote_files())
        self.assertEqual(self.command("diff", "--cached", "--name-only").stdout, "")

    def test_upstream_update_can_be_reviewed_and_conflicts_stop_sync(self):
        upstream = self.add_easel_subtree()
        self.publish("easel")
        self.write("Easel/README.md", "personal edit\n")
        self.command("add", "Easel/README.md")
        self.command("commit", "-m", "personal change")
        self.publish("easel")
        (upstream / "README.md").write_text("official edit\n")
        self.command("add", "README.md", cwd=upstream)
        self.command("commit", "-m", "official change", cwd=upstream)
        self.command("fetch", git_sync.UPSTREAM_REMOTE, "main")
        result = self.command("subtree", "pull", "--prefix=Easel", "--squash", git_sync.UPSTREAM_REMOTE, "main", check=False)
        self.assertNotEqual(result.returncode, 0)
        with self.assertRaisesRegex(RuntimeError, "未完成"):
            git_sync.check_workspace("easel")
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "show", "main:Easel/README.md").stdout, "personal edit\n")

    def test_verified_upstream_update_is_published_after_review(self):
        upstream = self.add_easel_subtree()
        self.publish("easel")
        (upstream / "new-feature.txt").write_text("official feature\n")
        self.command("add", "new-feature.txt", cwd=upstream)
        self.command("commit", "-m", "official feature", cwd=upstream)
        self.command("subtree", "pull", "--prefix=Easel", "--squash", git_sync.UPSTREAM_REMOTE, "main")
        self.publish("easel")
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "show", "main:Easel/new-feature.txt").stdout, "official feature\n")

    def test_forged_subtree_history_is_rejected(self):
        upstream = self.add_easel_subtree()
        self.publish("easel")
        baseline = self.command("rev-parse", "HEAD", cwd=upstream).stdout.strip()
        # Build a real object with content unrelated to the claimed upstream tree.
        payload = self.write("forged-private.tmp", "SIMULATED_PRIVATE_CONTENT")
        blob = self.command("hash-object", "-w", str(payload)).stdout.strip()
        tree_input = f"100644 blob {blob}\tprivate.txt\n"
        tree = subprocess.run(["git", "mktree"], cwd=self.root, input=tree_input, text=True,
                              capture_output=True, check=True).stdout.strip()
        message = f"fake squash\n\ngit-subtree-dir: Easel\ngit-subtree-split: {baseline}\n"
        forged = self.command("commit-tree", tree, "-m", message).stdout.strip()
        self.command("merge", "--no-ff", "--allow-unrelated-histories", "-s", "ours", forged, "-m", "fake history")
        with self.assertRaisesRegex(RuntimeError, "内容不一致"):
            self.publish("easel")

    def test_index_migration_preserves_dirty_docs_and_existing_runtime(self):
        self.publish()
        upstream = Path(self.temp.name) / "migration-upstream"
        upstream.mkdir()
        self.command("init", "-b", "main", cwd=upstream)
        self.command("config", "user.email", "upstream@example.invalid", cwd=upstream)
        self.command("config", "user.name", "Simulated upstream", cwd=upstream)
        (upstream / "README.md").write_text("source baseline\n")
        self.command("add", ".", cwd=upstream)
        self.command("commit", "-m", "source baseline", cwd=upstream)
        self.command("clone", str(upstream), str(self.root / "Easel"))
        self.write("Easel/.env", "simulated local configuration")
        self.write("Easel/.venv/runtime", "local runtime")
        self.write("Easel/outputs/local-data", "personal generated data")
        self.write("Easel/启动Easel.command", "local launcher")
        self.write("docs/原文档.md", "unsaved local docs change\n")
        runtime = self.root / "Easel/.venv/runtime"
        runtime_inode = runtime.stat().st_ino
        worktree = Path(self.temp.name) / "isolated-import"
        self.command("worktree", "add", "--detach", str(worktree), "HEAD")
        self.command("subtree", "add", "--prefix=Easel", "--squash", str(upstream), "main", cwd=worktree)
        migration = self.command("rev-parse", "HEAD", cwd=worktree).stdout.strip()
        (self.root / "Easel/.git").rename(self.root / ".git/easel-upstream-backup")
        self.command("reset", "--mixed", "--quiet", migration)
        self.assertEqual((self.root / "docs/原文档.md").read_text(), "unsaved local docs change\n")
        self.assertEqual((self.root / "Easel/.env").read_text(), "simulated local configuration")
        self.assertEqual((self.root / "Easel/outputs/local-data").read_text(), "personal generated data")
        self.assertEqual(runtime.stat().st_ino, runtime_inode)
        self.assertTrue((self.root / "Easel/启动Easel.command").is_file())
        self.assertEqual(self.command("--git-dir=.git/easel-upstream-backup", "rev-parse", "HEAD").stdout.strip(),
                         self.command("rev-parse", "HEAD", cwd=upstream).stdout.strip())
        self.assertEqual(self.command("diff", "--name-only", "--", "Easel/").stdout, "")

    def test_migration_keeps_local_files_and_only_publishes_docs(self):
        self.write("scripts/local.py", "modified local code\n")
        self.write("docs/新增图片.png", "simulated image")
        self.write("docs/.env", "SIMULATED_SECRET")
        self.write("docs/联系人.xlsx", "SIMULATED_WORKBOOK")
        self.publish()
        self.assertEqual(set(self.remote_files()), {"docs/原文档.md", "docs/新增图片.png"})
        self.assertEqual((self.root / "scripts/local.py").read_text(), "modified local code\n")
        self.assertTrue((self.root / "private.md").is_file())
        self.assertTrue((self.root / ".gitignore").is_file())
        self.assertEqual(self.command("status", "--porcelain", "--", "docs/").stdout, "")
        for name in ("scripts/new.py", "new-root.txt", "KOL搜集与邮件回复/private.md"):
            self.write(name, "simulated local data")
            self.assertEqual(self.command("check-ignore", name, check=False).returncode, 0)

    def test_empty_docs_clears_remote_and_repeated_sync_is_idempotent(self):
        (self.root / "docs/原文档.md").unlink()
        self.publish()
        self.assertEqual(self.remote_files(), [])
        head = self.command("rev-parse", "HEAD").stdout
        self.publish()
        self.assertEqual(self.command("rev-parse", "HEAD").stdout, head)
        self.assertTrue((self.root / "scripts/local.py").is_file())

    def test_updates_additions_and_deletions_after_empty_sync(self):
        (self.root / "docs/原文档.md").unlink()
        self.publish()
        self.write("docs/新目录/内容.txt", "version one")
        self.publish()
        self.assertEqual(self.remote_files(), ["docs/新目录/内容.txt"])
        self.write("docs/新目录/内容.txt", "version two")
        self.publish()
        result = self.command("--git-dir=" + str(self.remote), "show", "main:docs/新目录/内容.txt")
        self.assertEqual(result.stdout, "version two")
        (self.root / "docs/新目录/内容.txt").unlink()
        self.publish()
        self.assertEqual(self.remote_files(), [])

    def test_previously_staged_files_are_preserved_and_block_upload(self):
        self.write("staged-local.txt", "simulated local content")
        # Simulate an entry staged before the new ignore rules were installed.
        (self.root / ".gitignore").unlink()
        self.command("add", "staged-local.txt")
        self.write(".gitignore", IGNORE_RULES)
        with self.assertRaisesRegex(RuntimeError, "已暂存"):
            self.publish()
        self.assertIn("staged-local.txt", self.command("diff", "--cached", "--name-only").stdout)
        self.assertTrue((self.root / "staged-local.txt").is_file())

    def test_unknown_remote_changes_stop_before_staging(self):
        self.publish()
        other = Path(self.temp.name) / "other"
        self.command("clone", "--branch", "main", str(self.remote), str(other))
        self.command("config", "user.email", "other@example.invalid", cwd=other)
        self.command("config", "user.name", "Other uploader", cwd=other)
        (other / "docs/remote.txt").write_text("remote change")
        self.command("add", ".", cwd=other)
        self.command("commit", "-m", "remote change", cwd=other)
        self.command("push", "origin", "main", cwd=other)
        self.write("docs/local.txt", "local change")
        head = self.command("rev-parse", "HEAD").stdout
        with self.assertRaisesRegex(RuntimeError, "远端存在"):
            self.publish()
        self.assertEqual(self.command("rev-parse", "HEAD").stdout, head)
        self.assertEqual(self.command("diff", "--cached", "--name-only").stdout, "")
        self.assertIn("docs/remote.txt", self.remote_files())

    def test_unpushed_non_docs_history_is_rejected_even_if_tip_is_clean(self):
        self.publish()
        (self.root / ".gitignore").unlink()
        self.write("unapproved.txt", "simulated local-only data")
        self.command("add", "unapproved.txt")
        self.command("commit", "-m", "accidental local commit")
        self.command("rm", "--cached", "unapproved.txt")
        self.command("commit", "-m", "remove local-only data")
        self.write(".gitignore", IGNORE_RULES)
        with self.assertRaisesRegex(RuntimeError, "允许范围之外"):
            self.publish()
        self.assertEqual(self.remote_files(), ["docs/原文档.md"])

    def test_push_failure_preserves_commit_and_local_files(self):
        original_git = git_sync.git

        def fail_push(*args, **kwargs):
            if "push" in args:
                raise RuntimeError("simulated network failure")
            return original_git(*args, **kwargs)

        head = self.command("rev-parse", "HEAD").stdout
        with patch.object(git_sync, "git", side_effect=fail_push):
            with self.assertRaisesRegex(RuntimeError, "network failure"):
                self.publish()
        self.assertNotEqual(self.command("rev-parse", "HEAD").stdout, head)
        self.assertTrue((self.root / "scripts/local.py").is_file())
        self.publish()
        self.assertEqual(self.remote_files(), ["docs/原文档.md"])

    def test_symlink_and_ignored_staged_file_are_rejected(self):
        link = self.root / "docs/local-link"
        link.symlink_to(self.root / "scripts/local.py")
        with self.assertRaisesRegex(RuntimeError, "普通文件"):
            git_sync.check_workspace()
        link.unlink()
        secret = self.write("docs/.env", "SIMULATED_SECRET")
        self.command("hash-object", "-w", str(secret))
        blob = self.command("hash-object", str(secret)).stdout.strip()
        self.command("update-index", "--add", "--cacheinfo", "100644", blob, "docs/.env")
        with self.assertRaisesRegex(RuntimeError, "已暂存"):
            git_sync.check_workspace()

    def test_check_only_does_not_stage_commit_or_push(self):
        head = self.command("rev-parse", "HEAD").stdout
        before = self.command("status", "--porcelain").stdout
        with patch("sys.argv", ["git_sync.py", "--check-only"]), redirect_stdout(StringIO()):
            git_sync.main()
        self.assertEqual(self.command("status", "--porcelain").stdout, before)
        self.assertEqual(self.command("rev-parse", "HEAD").stdout, head)

    def test_tags_are_not_uploaded_when_follow_tags_is_configured(self):
        self.command("config", "push.followTags", "true")
        self.command("tag", "-a", "local-only-tag", "-m", "keep local")
        self.publish()
        self.assertEqual(self.command("--git-dir=" + str(self.remote), "tag").stdout, "")

    def test_push_url_mismatch_is_rejected(self):
        self.command("remote", "set-url", "--push", "origin", str(Path(self.temp.name) / "unexpected.git"))
        with self.assertRaisesRegex(RuntimeError, "上传地址"):
            git_sync.check_workspace()

    def tailki_fixture(self, document_text="Generic consolidated workflow"):
        self.publish()
        self.write("README.md", "Reusable Tailki instructions\n")
        self.write("AGENTS.md", "Publish only reviewed artifacts\n")
        self.write(git_sync.SKILL_ROOT + "SKILL.md", "Generic workflow rules\n")
        self.write("scripts/daily_workflow.py", "# generic workflow implementation\n")
        self.tailki_document(document_text)

    def tailki_document(self, text):
        path = self.root / git_sync.HANDBOOK
        path.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", "<document>" + text + "</document>")
        blob = self.command("hash-object", str(path)).stdout.strip()
        patch.dict(git_sync.AUDITED_TAILKI_BLOBS, {git_sync.HANDBOOK: blob}).start()

    def test_tailki_publishes_only_curated_files_and_preserves_unrelated_deletion(self):
        self.tailki_fixture()
        (self.root / "docs/原文档.md").unlink()
        excluded = [
            "KOL搜集与邮件回复/参考资料/历史文档/old.docx",
            "KOL搜集与邮件回复/参考资料/2026-11-06_达人回信审核.docx",
            "KOL搜集与邮件回复/达人跟进表.xlsx",
            "KOL搜集与邮件回复/参考资料/计划.xlsx",
            git_sync.SKILL_ROOT + "private.txt", "scripts/private.py",
            "scripts/examples/真实表.xlsx", "scripts/gitleaks",
        ]
        for name in excluded:
            self.write(name, "simulated private data")
            self.assertEqual(self.command("check-ignore", name, check=False).returncode, 0, name)
        self.publish("tailki")
        files = self.remote_files()
        self.assertIn(git_sync.HANDBOOK, files)
        self.assertIn(git_sync.SKILL_ROOT + "SKILL.md", files)
        self.assertIn("scripts/daily_workflow.py", files)
        self.assertIn("README.md", files)
        self.assertIn("docs/原文档.md", files)
        self.assertTrue(all(name not in files for name in excluded))
        self.assertIn(" D docs/原文档.md\0", self.command("status", "--porcelain", "-z").stdout)
        head = self.command("rev-parse", "HEAD").stdout
        self.publish("tailki")
        self.assertEqual(self.command("rev-parse", "HEAD").stdout, head)

    def test_tailki_archive_cannot_enter_commit_even_if_force_staged(self):
        self.tailki_fixture()
        path = "KOL搜集与邮件回复/参考资料/历史文档/old.docx"
        self.write(path, "private source document")
        self.command("add", "-f", path)
        with self.assertRaisesRegex(RuntimeError, "允许范围之外"):
            git_sync.check_tree(self.command("write-tree").stdout.strip())

    def test_tailki_changed_binary_requires_review(self):
        self.tailki_fixture()
        self.write(git_sync.SKILL_ROOT + "assets/tailki-collar.png", "unreviewed image")
        with self.assertRaisesRegex(RuntimeError, "尚未完成内容核对"):
            self.publish("tailki")
        self.assertEqual(self.command("diff", "--cached", "--name-only").stdout, "")
        self.assertNotIn(git_sync.HANDBOOK, self.remote_files())

    def test_tailki_checks_secrets_inside_docx(self):
        secret = "ghp_" + "a1B2c3D4e5F6g7H8i9J0k1L2m3N4o5P6q7R8"
        self.tailki_fixture("GitHub access: " + secret)
        head = self.command("rev-parse", "HEAD").stdout
        with self.assertRaisesRegex(RuntimeError, "内部内容未通过密钥检查") as error:
            self.publish("tailki")
        self.assertNotIn(secret, str(error.exception))
        self.assertEqual(self.command("rev-parse", "HEAD").stdout, head)
        self.assertEqual(self.command("diff", "--cached", "--name-only").stdout, "")

    def test_tailki_never_pushes_unrelated_pending_commits(self):
        self.tailki_fixture()
        self.write("docs/原文档.md", "unrelated pending change")
        self.command("add", "docs/原文档.md")
        self.command("commit", "-m", "unrelated change")
        with self.assertRaisesRegex(RuntimeError, "以外的改动"):
            self.publish("tailki")
        self.assertNotIn(git_sync.HANDBOOK, self.remote_files())
        self.assertEqual(self.command("diff", "--cached", "--name-only").stdout, "")


if __name__ == "__main__":
    unittest.main()
