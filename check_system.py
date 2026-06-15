from __future__ import annotations

import ast
import re
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SERVER = ROOT / "ui" / "server.py"
TEMPLATE_DIRS = [ROOT / "ui" / "templates", ROOT / "templates"]

REQUIRED_FUNCTIONS = {
    "update_offer_meta",
    "complete_application",
    "set_no_plate_for_offer",
    "delete_offer",
    "add_no_cache_headers",
}


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


def ok(message: str) -> None:
    print(f"OK: {message}")


def check_py_compile() -> None:
    result = subprocess.run(
        [sys.executable, "-m", "py_compile", str(SERVER)],
        cwd=ROOT,
        text=True,
        capture_output=True,
    )
    if result.returncode != 0:
        print(result.stdout)
        print(result.stderr)
        fail("ui/server.py does not compile")
    ok("ui/server.py compiles")


def check_templates_exist(source: str) -> None:
    templates = sorted(set(re.findall(r'render_template\("([^"]+)"', source)))
    missing = []

    for template in templates:
        if not any((directory / template).exists() for directory in TEMPLATE_DIRS):
            missing.append(template)

    if missing:
        fail("missing templates: " + ", ".join(missing))
    ok(f"all {len(templates)} referenced templates exist")


def check_functions(source: str) -> None:
    tree = ast.parse(source)
    found = {node.name for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)}
    missing = sorted(REQUIRED_FUNCTIONS - found)

    if missing:
        fail("missing functions: " + ", ".join(missing))

    count = sum(1 for node in ast.walk(tree) if isinstance(node, ast.FunctionDef) and node.name == "add_no_cache_headers")
    if count != 1:
        fail(f"expected one add_no_cache_headers definition, found {count}")

    ok("required functions exist and add_no_cache_headers is unique")


def check_customer_name_parsing() -> None:
    from mailgen import guess_aanhef_en_achternaam, split_dealer_customer_name

    surname, initials = split_dealer_customer_name("De Jong, D.")
    if (surname, initials) != ("De Jong", "D."):
        fail(f"dealer name split failed: {(surname, initials)!r}")

    aanhef, achternaam = guess_aanhef_en_achternaam("De Jong, D.")
    if achternaam != "De Jong":
        fail(f"dealer surname parsing failed: {(aanhef, achternaam)!r}")

    aanhef, achternaam = guess_aanhef_en_achternaam("Mevr. J van der Meer")
    if (aanhef, achternaam) != ("mevrouw", "van der Meer"):
        fail(f"regular surname parsing failed: {(aanhef, achternaam)!r}")

    ok("customer name parsing handles dealer format")


def main() -> None:
    if not SERVER.exists():
        fail("ui/server.py not found")

    source = SERVER.read_text(encoding="utf-8")
    check_py_compile()
    check_templates_exist(source)
    check_functions(source)
    check_customer_name_parsing()
    ok("system checks passed")


if __name__ == "__main__":
    main()
