"""Shared helpers. Tests run in temporary copies of the demo experiment (examples/demo_binder_screen), which is
simulated data, so nothing here needs private files."""

import contextlib
import io
import os
import shutil
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

DEMO = os.path.join(ROOT, "examples", "demo_binder_screen")


def run_cli(*argv):
    """Run the CLI in-process -> (exit code, stdout, stderr). A message exit is code 1."""
    from bli2prism.cli import main
    out, err = io.StringIO(), io.StringIO()
    code = 0
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try:
            code = main([str(a) for a in argv]) or 0
        except SystemExit as e:
            code = e.code if isinstance(e.code, int) else 1
            if isinstance(e.code, str):
                err.write(e.code)
    return code, out.getvalue(), err.getvalue()


@pytest.fixture()
def demo_run(tmp_path):
    """A private, writable copy of the demo experiment's run folder (raw files, results CSV, setup workbook)."""
    d = str(tmp_path / "run")
    shutil.copytree(DEMO, d, ignore=shutil.ignore_patterns("expected_output", "rebuild"))
    return d
