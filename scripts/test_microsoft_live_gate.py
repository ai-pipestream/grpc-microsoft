"""Offline regression checks: python3 scripts/test_microsoft_live_gate.py.

Runs the workflow's actual shell blocks with synthetic credentials/XML only.
Uses the Python standard library; never starts Gradle or contacts a service.
"""
import os
from pathlib import Path
import subprocess
import tempfile
import textwrap
import unittest

ROOT = Path(__file__).resolve().parents[1]
WORKFLOW = (ROOT / ".github/workflows/ci.yml").read_text()
microsoft = "microsoft"
Microsoft = "Microsoft"
KEYS = ['MICROSOFT_TENANT_ID', 'MICROSOFT_CLIENT_ID', 'MICROSOFT_CLIENT_SECRET']


def run_block(step_name):
    tail = WORKFLOW.split("      - name: " + step_name + "\n", 1)[1]
    block = tail.split("        run: |\n", 1)[1]
    lines = []
    for line in block.splitlines():
        if line.strip() and not line.startswith("          "):
            break
        lines.append(line)
    return textwrap.dedent("\n".join(lines))


class LiveGateTest(unittest.TestCase):
    def preflight(self, values, required, expected_code, ready):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            env = {"PATH": os.defpath, "GITHUB_OUTPUT": str(folder / "output"),
                   "GITHUB_STEP_SUMMARY": str(folder / "summary"),
                   "REQUIRE_LIVE": str(required).lower(), **dict(zip(KEYS, values))}
            result = subprocess.run(
                ["bash", "-e", "-c", run_block("Check live credentials without printing them")],
                env=env, capture_output=True, text=True)
            self.assertEqual(result.returncode, expected_code)
            self.assertEqual((folder / "output").read_text().strip(), "ready=" + ready)
            summary = (folder / "summary").read_text() if (folder / "summary").exists() else ""
            output = result.stdout + result.stderr + summary
            self.assertNotIn("synthetic-secret", output)
            if ready == "false":
                self.assertIn("FAILED" if required else "SKIPPED", summary)
                self.assertIn("no live evidence", summary)

    def test_missing_optional(self):
        self.preflight([""] * len(KEYS), False, 0, "false")

    def test_missing_required(self):
        self.preflight([""] * len(KEYS), True, 1, "false")

    def test_each_credential_required(self):
        for key in range(len(KEYS)):
            with self.subTest(key=KEYS[key]):
                values = ["synthetic-secret"] * len(KEYS)
                values[key] = ""
                self.preflight(values, True, 1, "false")

    def test_whitespace_optional(self):
        self.preflight([" \t\n"] * len(KEYS), False, 0, "false")

    def test_whitespace_required(self):
        self.preflight([" \t\n"] * len(KEYS), True, 1, "false")

    def test_present_optional(self):
        self.preflight(["synthetic-secret"] * len(KEYS), False, 0, "true")

    def test_present_required(self):
        self.preflight(["synthetic-secret"] * len(KEYS), True, 0, "true")

    def evidence(self, xml, success):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            report = folder / (f"grpc-{microsoft}-service/build/test-results/liveSmokeTest/"
                               f"TEST-ai.pipestream.{microsoft}.{Microsoft}LiveSmokeIT.xml")
            if xml is not None:
                report.parent.mkdir(parents=True)
                report.write_text(xml)
            result = subprocess.run(
                ["bash", "-e", "-c", run_block("Require executed live evidence")],
                cwd=folder, env={"PATH": os.defpath, "GITHUB_STEP_SUMMARY": str(folder / "summary")},
                capture_output=True, text=True)
            self.assertEqual(result.returncode == 0, success, result.stderr)
            self.assertEqual((folder / "summary").exists(), success)

    def test_xml_missing(self):
        self.evidence(None, False)

    def test_xml_empty(self):
        self.evidence("<testsuite/>", False)

    def test_xml_skipped(self):
        self.evidence("<testsuite><testcase><skipped/></testcase></testsuite>", False)

    def test_xml_failed(self):
        self.evidence("<testsuite><testcase><failure/></testcase></testsuite>", False)

    def test_xml_error(self):
        self.evidence("<testsuite><testcase><error/></testcase></testsuite>", False)

    def test_xml_executed(self):
        self.evidence("<testsuite><testcase/></testsuite>", True)

    def test_fork_and_dispatch_guards(self):
        for job in (f"live-{microsoft}-preflight", f"live-{microsoft}"):
            body = WORKFLOW.split(f"  {job}:\n", 1)[1].split("    steps:", 1)[0]
            self.assertIn("github.event_name != 'pull_request' || "
                          "github.event.pull_request.head.repo.full_name == github.repository", body)
        self.assertIn(f"inputs.require_live_{microsoft}", WORKFLOW)
        self.assertIn(f"needs.live-{microsoft}-preflight.outputs.ready == 'true'", WORKFLOW)
        self.assertIn(":liveSmokeTest --rerun-tasks", WORKFLOW)


if __name__ == "__main__":
    unittest.main(verbosity=2)
