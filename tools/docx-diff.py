#!/usr/bin/env python3
"""Compare a DOCX written by the probe with the original (spec T3, Mac verification).

    python3 tools/docx-diff.py probe-fixture.docx probe-fixture-edit-123456.docx

Reports every package entry whose uncompressed bytes differ, then, for each differing XML
part, a unified diff with one tag per line, and whether the root start tag changed.
Exit status 0 when nothing differs, 1 otherwise.
"""
import difflib
import hashlib
import re
import sys
import zipfile


def entries(path):
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


def tag_lines(xml: bytes):
    text = xml.decode("utf-8", errors="replace")
    return re.sub(r">\s*<", ">\n<", text).splitlines()


def root_tag(xml: bytes) -> str:
    text = xml.decode("utf-8", errors="replace")[:8192]
    m = re.search(r"<\?xml[^>]*\?>\s*<[^!?][^>]*>", text, re.S)
    return m.group(0) if m else text[:400]


def main(original, copy):
    a, b = entries(original), entries(copy)
    changed = []
    for name in sorted(set(a) | set(b)):
        if name not in b:
            print(f"MISSING in copy: {name}")
            changed.append(name)
        elif name not in a:
            print(f"ADDED in copy:   {name}")
            changed.append(name)
        elif hashlib.sha256(a[name]).digest() != hashlib.sha256(b[name]).digest():
            print(f"DIFFERS:         {name}")
            changed.append(name)
    same = len(set(a) & set(b)) - sum(1 for n in changed if n in a and n in b)
    print(f"\n{same} entries identical, {len(changed)} differ\n")

    for name in changed:
        if name in a and name in b and name.endswith((".xml", ".rels")):
            ra, rb = root_tag(a[name]), root_tag(b[name])
            print(f"=== {name}: root tag {'UNCHANGED' if ra == rb else 'CHANGED'}")
            if ra != rb:
                print(f"--- before\n{ra}\n--- after\n{rb}\n")
            diff = difflib.unified_diff(
                tag_lines(a[name]), tag_lines(b[name]), "original", "copy", lineterm="", n=2
            )
            print("\n".join(diff))
            print()
    return 1 if changed else 0


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    sys.exit(main(sys.argv[1], sys.argv[2]))
