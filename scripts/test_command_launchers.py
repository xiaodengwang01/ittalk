"""Exercise the review launcher in an isolated project; never access real mail."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest


@unittest.skipUnless(shutil.which("zsh"), "The macOS launcher requires zsh")
class ReviewLauncherTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="tailki-launcher-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "project with spaces"
        scripts = self.root / "scripts"
        scripts.mkdir(parents=True)
        self.launcher = scripts / "今日审核.command"
        shutil.copy2(Path(__file__).with_name("今日审核.command"), self.launcher)
        self.program = scripts / "review_center.py"
        self.program.write_text(
            "import json, pathlib, sys\n"
            "pathlib.Path('invocation.json').write_text(json.dumps(sys.argv[1:]))\n"
            "raise SystemExit(0)\n"
        )
        self.env = {**os.environ, "HOME": str(Path(self.temp.name) / "home"),
                    "TAILKI_REVIEW_PYTHON": sys.executable}

    def launch(self):
        return subprocess.run([shutil.which("zsh"), str(self.launcher)],
                              cwd=self.temp.name, env=self.env,
                              stdin=subprocess.DEVNULL, capture_output=True,
                              text=True, timeout=15)

    def test_opens_review_from_an_unrelated_directory_with_spaces(self):
        result = self.launch()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads((self.root / "invocation.json").read_text()),
                         ["build", "--open"])

    def test_generation_error_preserves_exit_code_and_diagnostic(self):
        self.program.write_text("import sys\nprint('simulated build conflict', file=sys.stderr)\nraise SystemExit(23)\n")
        result = self.launch()
        self.assertEqual(result.returncode, 23, result.stderr)
        self.assertIn("simulated build conflict", result.stderr)
        self.assertIn("原资料已保留", result.stderr)

    def test_missing_bundled_runtime_uses_compatible_python_on_path(self):
        self.env.pop("TAILKI_REVIEW_PYTHON")
        binary = Path(self.temp.name) / "bin"
        binary.mkdir()
        (binary / "python3").symlink_to(sys.executable)
        self.env["PATH"] = str(binary) + os.pathsep + os.environ.get("PATH", "")
        result = self.launch()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue((self.root / "invocation.json").exists())

    def test_unusable_runtime_explains_dependency_without_starting_build(self):
        binary = Path(self.temp.name) / "bin"
        binary.mkdir()
        broken = binary / "python3"
        broken.write_text("#!/bin/sh\nexit 1\n")
        broken.chmod(0o755)
        (binary / "dirname").symlink_to(shutil.which("dirname"))
        self.env.update(PATH=str(binary), TAILKI_REVIEW_PYTHON=str(broken))
        result = self.launch()
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("python-docx", result.stderr)
        self.assertFalse((self.root / "invocation.json").exists())

    def test_missing_program_has_actionable_error(self):
        self.program.unlink()
        result = self.launch()
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("今日审核程序缺失", result.stderr)


if __name__ == "__main__":
    unittest.main()
