#!/usr/bin/env python3
"""Search the official DraftSight help the way it should be searched.

DraftSight ships 6,457 pages of CHM help. Unpacking them costs about 12 s and 63 MB; *reading* them
costs a context window. So: extract once, then resolve pages **by filename**, because the API tree
names its pages `Interface~Member`.

    python examples/docs_lookup.py                 # three worked lookups + a status report
    python examples/docs_lookup.py ExportToPng     # one lookup of your choosing

From a shell the same capability is `ds-docs decompile` / `ds-docs lookup ...`.

Cost, measured on the box this was written on: ~2 s for this whole script once the tree exists,
~12 s extra the first time to unpack the CHMs with 7-Zip, and 0.3 s for a filename lookup versus
1.7 s for a full-text one.
"""
import sys

from draftsight_com import docs

# Help pages are UTF-8; a Korean Windows console is cp949 and raises UnicodeEncodeError on the first
# non-breaking space. The ds-docs CLI already reconfigures its own stdout; a script has to do the same.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> int:
    # Once per machine: gitignored, vendor copyrighted, fast with 7-Zip and tolerable with the
    # hh.exe fallback. Note where the tree lands: docs.py anchors it to the *package* directory, so
    # an editable install decompiles into your checkout and a wheel install into site-packages.
    if not docs.find_pages("IDocumentExporter~ExportToPng"):
        print("decompiling the installed CHMs (once per machine) ...", flush=True)
        print(docs.decompile())

    if len(sys.argv) > 1:
        print(docs.lookup(sys.argv[1], pages=1, max_lines=40))
        return 0

    queries = [
        # A member's real signature, and what each parameter means.
        dict(pattern="IDocumentExporter~ExportToPng", max_lines=30),
        # Is this a real RunCommand name? ZoomExtents is not; ZoomFit is.
        dict(pattern="ZoomFit", cmdref=True, max_lines=20),
        # An enum family, found by full text rather than by page name.
        dict(pattern="dsLnWt", content=True, pages=2, max_lines=15),
    ]
    for q in queries:
        print("=" * 72)
        print("lookup:", q["pattern"], {k: v for k, v in q.items() if k != "pattern"})
        print(docs.lookup(**q))
    print("=" * 72)
    print(docs.status())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
