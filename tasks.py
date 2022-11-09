# coding: utf-8

import sys
import subprocess
from invoke import task
from pathlib import Path


REPO_HOME = Path.cwd()


def get_python_interpreters():
    available_pythons = [
        pyver
        for pyver in subprocess.check_output("py -0").decode("utf-8").split()
        if pyver.startswith("-")
    ]
    for python in available_pythons:
        exe = (
            subprocess.check_output(
                ["py", python, "-c", "import sys;print(sys.executable)"]
            )
            .decode(sys.getdefaultencoding())
            .strip()
        )
        yield (exe, python[-2:])


@task
def build_wheels(c, release=True, strip=True, sdist=True):
    i_args = {
        exe: "i686-pc-windows-msvc" if arch == "32" else "x86_64-pc-windows-msvc"
        for (exe, arch) in get_python_interpreters()
    }
    for (interpreter_path, target) in i_args.items():
        print(f"Building wheel for Python {interpreter_path} using target {target}")
        build_command = " ".join(
            [
                "maturin build",
                "--release" if release else "",
                "--strip" if strip else "",
                "--sdist",
                f"-i {interpreter_path}",
                f"--target {target}",
            ]
        )
        c.run(build_command)


@task
def upload_wheels(c):
    with c.cd(REPO_HOME):
        c.run(r'twine upload  "./target/wheels/*" --non-interactive --skip-existing')
