#!/usr/bin/env python3
"""Build families into OUT, then check the whole corpus.

    python3 -I make_corpus.py OUT [f00_demo ...]     (no names: every families/f*.py)
"""

import importlib.util
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import check_corpus  # noqa: E402


def main(argv: list[str]) -> int:
    out, wanted = Path(argv[1]).resolve(), argv[2:]
    fams = sorted((HERE / "families").glob("f*.py"))
    unknown = set(wanted) - {f.stem for f in fams}
    if unknown:
        raise SystemExit(f"unknown families: {sorted(unknown)}")
    for f in fams:
        if wanted and f.stem not in wanted:
            continue
        shutil.rmtree(out / f.stem, ignore_errors=True)  # a family owns out/<stem>/ and nothing else
        spec = importlib.util.spec_from_file_location(f.stem, f)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        mod.build(out)
        print(f"built {f.stem}: {len(list((out / f.stem).glob('*.key.json')))} documents")
    return check_corpus.main(out)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
