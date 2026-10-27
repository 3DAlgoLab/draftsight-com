"""The official DraftSight help as a search target: decompile the installed CHMs, then look things up.

The installed Help folder unpacks to 6,457 pages / 63 MB (measured). Pages are 1-4 KB each, which sounds
readable and is not: picking the right one out of 6,457 by opening them is how context gets destroyed.
So the tree is generated once per machine, gitignored (vendor-copyrighted), and only ever *searched*.

Shell:

    ds-docs decompile                                  # build docs/ from the installed CHMs (12s measured)
    ds-docs lookup ISketchManager~Align3DEntities      # prose for one API member, one page
    ds-docs lookup -t ZoomFit                          # command reference only (valid RunCommand names)
    ds-docs lookup -c ExportToPng -n 3                 # full-text search, a few pages
    ds-docs status                                     # sources, extractor, and what is generated

Python:

    from draftsight_com import docs
    print(docs.lookup("IDocumentExporter~ExportToPng"))   # signature, parameter meaning, examples
    docs.decompile()                                      # idempotent; existing sets come back "cached"

`lookup` resolves a page **by filename** first: the API tree names pages `Interface~Member`, e.g.
`docs/draftsightapi/Interop.dsAutomation~Interop.dsAutomation.ISketchManager~Align3DEntities.html`, so a
name query is exact and needs no reading. Content search is the fallback. Output is capped at `-n` pages
stripped to `-l` lines, so one lookup costs tens of lines no matter how big the tree is.

Order of trust for any API question: the type library (exact signatures and enum numbers), then these
pages (what parameters mean, which RunCommand names exist, vendor examples), then a live probe (what
this build actually does).
"""

import argparse
import html
import os
import re
import shutil
import subprocess
import sys
import time

# --- sources -----------------------------------------------------------------

HELP_DIR_ENV = "DRAFTSIGHT_HELP_DIR"
DOCS_ENV = "DRAFTSIGHT_DOCS"

# Install root can be on any drive, and a root can exist as a Fonts-only stub, so probe for a real CHM
# rather than trusting the directory.
HELP_ROOTS = (r"D:\Program Files\Dassault Systemes\DraftSight",
              r"C:\Program Files\Dassault Systemes\DraftSight")
API_CHM = "draftsightapi.chm"

# Value-ordered. draftsightapi.chm is the COM API reference, DraftSight.chm carries the command and
# command-variable reference; the rest are niche and cheap to keep.
DEFAULT_CHMS = (API_CHM, "DraftSight.chm", "DraftSight_Lisp_Reference.chm", "DSAPIHelp.chm",
                "DraftSightMech.chm", "DraftSightSW.chm", "NestingManager.chm", "DraftSightConnected.chm")

HH_EXE = r"C:\Windows\hh.exe"
SEVENZIP_CANDIDATES = (r"C:\Program Files\7-Zip\7z.exe",
                       r"C:\Program Files (x86)\7-Zip\7z.exe",
                       os.path.expandvars(r"%LOCALAPPDATA%\Programs\7-Zip\7z.exe"))

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DEFAULT_DOCS = os.path.join(SKILL_ROOT, "docs")

# Subtree inside docs/DraftSight holding the command reference. Lookups that validate a RunCommand
# string should scope here instead of the whole tree, so a name collision in the API reference cannot
# approve a command that does not exist.
CMDREF = os.path.join("DraftSight", "cmdref")


def help_dir(missing_ok=False):
    """Return the installed `Help\\english` folder holding the CHMs, or None when `missing_ok`."""
    env = os.environ.get(HELP_DIR_ENV)
    cands = [env] if env else [os.path.join(r, "Help", "english") for r in HELP_ROOTS]
    for c in cands:
        if c and os.path.isfile(os.path.join(c, API_CHM)):
            return c
    if missing_ok:
        return None
    raise RuntimeError("DraftSight Help\\english not found (probed: %s). Set %s to <install root>\\Help\\english"
                       % (", ".join(cands), HELP_DIR_ENV))


def docs_dir(create=False):
    """Return the generated help tree directory (gitignored, regenerable with `decompile`)."""
    d = os.environ.get(DOCS_ENV) or DEFAULT_DOCS
    if create:
        os.makedirs(d, exist_ok=True)
    return d


def chm_sources(help=None):
    """Return every CHM in the help folder as (name, full path), sorted by the default value order."""
    help = help or help_dir()
    have = {f.lower(): f for f in os.listdir(help) if f.lower().endswith(".chm")}
    ordered = [(have[n.lower()], os.path.join(help, have[n.lower()])) for n in DEFAULT_CHMS if n.lower() in have]
    extra = sorted(v for k, v in have.items() if k not in {n.lower() for n in DEFAULT_CHMS})
    return ordered + [(os.path.basename(p), p) for p in extra]


# --- extraction --------------------------------------------------------------

def sevenzip():
    """Path to a 7-Zip binary that can unpack CHM containers, or None."""
    found = shutil.which("7z") or shutil.which("7za")
    if found:
        return found
    return next((c for c in SEVENZIP_CANDIDATES if os.path.isfile(c)), None)


def _unpack_7z(z, chm, out, timeout):
    # CHM is a container 7-Zip understands: one quiet pass, exit code 0 means it wrote the tree.
    r = subprocess.run([z, "x", "-y", "-bso0", "-bsp0", "-o" + out, chm],
                       capture_output=True, text=True, timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError("exit %d: %s" % (r.returncode, " ".join((r.stderr or r.stdout or "").split())[:160]))


def _unpack_hh(chm, out, timeout):
    """Decompile with Windows' own hh.exe, launched through PowerShell, and only this way.

    Verified on this machine: hh.exe launched directly (CreateProcess, `cmd /c`, ShellExecuteEx, with or
    without CREATE_NEW_CONSOLE, output dir existing or not) returns exit code 0 in 30 ms and writes
    nothing at all - only `Start-Process -Wait` makes it do the work. A silent no-op is the worst failure
    mode there is, so the fallback is the proven invocation rather than the obvious one.
    """
    cmd = ("Start-Process -FilePath '%s' -ArgumentList '-decompile','%s','%s' -Wait"
           % (HH_EXE, out, chm))
    subprocess.run(["powershell", "-NoProfile", "-Command", cmd], capture_output=True, text=True,
                   timeout=timeout)


def _extract(chm, out, timeout):
    """Unpack one CHM into `out`; return (tool, note). 7-Zip first, `hh.exe -decompile` as fallback.

    7-Zip reads CHM containers natively and produced the same 4,353 API pages as hh.exe in a fifth of the
    time. Some books still defeat it - `DraftSightInstallAdminLicensing.chm` does on this machine - so a
    7-Zip failure is not fatal, it retries with hh.exe and reports which tool won in `note`.
    """
    note = ""
    z = sevenzip()
    if z:
        try:
            _unpack_7z(z, chm, out, timeout)
            return "7z", ""
        except Exception as e:                                # noqa: BLE001 - hh.exe is the fallback
            note = "7z failed: %s" % e
    try:
        _unpack_hh(chm, out, timeout)
    except Exception as e:                                    # noqa: BLE001 - surface both attempts
        raise RuntimeError(" ; ".join(x for x in (note, "hh.exe: %s" % e) if x))
    return ("hh.exe", note) if not note else ("hh.exe*", note)


def decompile(names=None, all_chms=False, help=None, out_dir=None, force=False, timeout=900):
    """Decompile CHM help files into a greppable tree. Idempotent: a set that already has pages is skipped.

    `names` selects CHMs (case-insensitive, `.chm` optional); `all_chms=True` takes everything in the
    help folder. Returns one row per CHM: name, dir, pages, bytes, seconds, tool, status. A row whose
    status is not "ok" or "cached" is a real problem - `report()` renders the rows as a table.
    """
    help = help or help_dir()
    out_root = out_dir or docs_dir(create=True)
    sources = dict(chm_sources(help))
    if all_chms:
        picked = list(sources)
    elif names:
        want = {n[:-4].lower() if n.lower().endswith(".chm") else n.lower() for n in names}
        have = {n[:-4].lower(): n for n in sources}
        picked = [have[k] for k in sorted(want) if k in have]
        for k in sorted(want - set(have)):
            print("skip: %s.chm not found in %s" % (k, help), file=sys.stderr)
    else:
        picked = [n for n in DEFAULT_CHMS if n in sources]

    rows = []
    for name in picked:
        out = os.path.join(out_root, name[:-4])
        pages = _pages(out)
        if pages and not force:
            rows.append({"chm": name, "dir": out, "pages": pages, "bytes": _bytes(out),
                         "seconds": 0.0, "tool": "cached", "status": "cached"})
            continue
        if pages and force:
            shutil.rmtree(out, ignore_errors=True)
        os.makedirs(out, exist_ok=True)
        t0 = time.time()
        try:
            tool, note = _extract(sources[name], out, timeout)
        except Exception as e:                                # noqa: BLE001 - report per file, keep going
            rows.append({"chm": name, "dir": out, "pages": 0, "bytes": 0, "seconds": round(time.time() - t0, 1),
                         "tool": "-", "status": "failed: %s" % e})
            continue
        pages = _pages(out)
        rows.append({"chm": name, "dir": out, "pages": pages, "bytes": _bytes(out),
                     "seconds": round(time.time() - t0, 1), "tool": tool,
                     "status": ("ok" if pages else "empty: no html pages")
                                 + ((" (" + note + ")") if note else "")})
    return rows


def report(rows):
    """Render `decompile` rows as a compact table, plus the total."""
    if not rows:
        return "nothing decompiled"
    lines = []
    for r in rows:
        lines.append("%-34s %6d pages  %6.1f MB  %6.1fs  %-7s %s"
                     % (r["chm"], r["pages"], r["bytes"] / 1e6, r["seconds"], r["tool"], r["status"]))
    lines.append("total %d pages, %.1f MB -> %s" % (sum(r["pages"] for r in rows),
                                                    sum(r["bytes"] for r in rows) / 1e6,
                                                    os.path.dirname(rows[0]["dir"])))
    return "\n".join(lines)


# --- html -> text ------------------------------------------------------------

_SCRIPT_STYLE = re.compile(r"<(script|style)[^>]*>.*?</\1>", re.I | re.S)
_TAG = re.compile(r"<[^>]+>")
_SPACES = re.compile(r"[ \t]+")


def html2text(text):
    """Strip markup so a help page can be read or grepped as prose (entities decoded, markup junk dropped).

    Every line is flushed to its text: a decompiled page is 119 lines of which 41 carry anything, and the
    rest is table and nav indentation that would otherwise eat the context budget of a lookup. Code keeps
    its line structure, not its indentation - the pages render snippets inside layout tables anyway.
    """
    text = _SCRIPT_STYLE.sub(" ", text)
    text = _TAG.sub(" ", text)
    text = html.unescape(text)          # after tag stripping, so escaped markup stays as literal text
    text = _SPACES.sub(" ", text.lstrip("\ufeff"))
    return "\n".join(l.strip() for l in text.splitlines() if l.strip())


def _read(path):
    """Read one help page as text; the tree is Windows-1252-ish html with stray bytes, so never raise on decode."""
    with open(path, "rb") as fh:
        return fh.read().decode("utf-8", "ignore")


def page(path, max_lines=70, raw=False):
    """Return one page as prose (or raw html when `raw`), truncated to `max_lines`."""
    blob = _read(path)
    text = blob if raw else html2text(blob)
    if raw:
        return text[:4000]
    lines = text.splitlines()
    out = "\n".join(lines[:max_lines])
    if len(lines) > max_lines:
        out += "\n... %d more lines (page max_lines=%d shows the whole page)" % (len(lines) - max_lines, len(lines))
    return out


def html2txt(path, max_lines=0):
    """Public one-file converter: prose for a decompiled page. `max_lines=0` prints the whole page."""
    text = html2text(_read(path))
    return "\n".join(text.splitlines()[:max_lines]) if max_lines else text


# --- lookup ------------------------------------------------------------------

_HTML = (".htm", ".html", ".htm2", ".xhtml")


def _pages(d):
    return sum(1 for dp, _, fs in os.walk(d) for f in fs if f.lower().endswith(_HTML)) if os.path.isdir(d) else 0


def _bytes(d):
    return sum(os.path.getsize(os.path.join(dp, f)) for dp, _, fs in os.walk(d) for f in fs
               if os.path.isfile(os.path.join(dp, f))) if os.path.isdir(d) else 0


def _walk(scope):
    for dp, _, fs in os.walk(scope):
        for f in fs:
            if f.lower().endswith(_HTML):
                yield os.path.join(dp, f)


def find_pages(pattern, scope=None, content=False, limit=10, docs=None):
    """Return up to `limit` page paths matching `pattern`: filename match first, then page content.

    Filenames in the decompiled API tree embed `Interface~Member`, so a filename hit is exact; hits are
    ranked shortest-path-first, which puts the member page ahead of its `_members` index page.
    """
    docs = docs or docs_dir()
    scope = scope or docs
    if content:
        needle = pattern.lower()
        hits = [p for p in _walk(scope) if needle in _read(p).lower()]
        return hits[:limit]
    needle = pattern.lower()
    hits = [p for p in _walk(scope) if needle in os.path.relpath(p, docs).lower().replace(os.sep, "/")]
    hits.sort(key=lambda p: (len(p), p))
    return hits[:limit]


def lookup(pattern, pages=1, max_lines=70, cmdref=False, scope=None, content=False, raw=False, docs=None):
    """Context-safe lookup in the decompiled help tree; returns the prose of at most `pages` pages.

    `cmdref=True` restricts the search to the command reference, the only honest way to check whether a
    `RunCommand` name exists (DraftSight is not AutoCAD: `ZoomExtents` does not exist, `ZoomFit` does).
    `content=True` searches page text instead of filenames. Returns text beginning with `no match` when
    nothing hits, and a guidance message when the tree has not been generated yet.
    """
    docs = docs or docs_dir()
    if not os.path.isdir(docs):
        return ("docs/ is not generated. It is deliberately not shipped (vendor-copyrighted: 63 MB, built in\n"
                "12 s):\n  ds-docs decompile\n"
                "For signatures and known quirks you may not need it: see references/draftsight-api.md in\n"
                "this skill, or probe a live session through draftsight_com.")
    if cmdref:
        scope = os.path.join(docs, CMDREF)
    elif not scope:
        scope = docs
    elif not os.path.isabs(scope):
        scope = os.path.join(docs, scope)
    if not os.path.isdir(scope):
        return "no such scope: %s" % scope

    hits = find_pages(pattern, scope=scope, content=content, limit=pages, docs=docs)
    if not hits and not content:
        hits = find_pages(pattern, scope=scope, content=True, limit=pages, docs=docs)
        note = "no page name matched; searched page content instead\n" if hits else ""
    else:
        note = ""
    if not hits:
        return "no match for: %s  (scope: %s)" % (pattern, os.path.relpath(scope, docs).replace(os.sep, "/"))
    out = [note] if note else []
    for p in hits:
        out.append("=== %s\n%s" % (os.path.relpath(p, docs).replace(os.sep, "/"), page(p, max_lines, raw)))
    return "\n".join(out)


def status():
    """Return a short report: help sources found, extractor available, and what is generated under docs/."""
    help = help_dir(missing_ok=True)
    out = ["help folder : %s" % (help or "NOT FOUND (set %s)" % HELP_DIR_ENV),
           "extractor   : %s" % (sevenzip() or "hh.exe via PowerShell Start-Process -Wait")]
    docs = docs_dir()
    if not os.path.isdir(docs):
        out.append("docs        : %s (not generated - run: ds-docs decompile)" % docs)
        return "\n".join(out)
    out.append("docs        : %s" % docs)
    for name in sorted(os.listdir(docs)):
        d = os.path.join(docs, name)
        if os.path.isdir(d):
            out.append("  %-30s %5d pages  %6.1f MB" % (name, _pages(d), _bytes(d) / 1e6))
    return "\n".join(out)


# --- cli ---------------------------------------------------------------------

def cli(argv=None):
    """Entry point for `ds-docs`: status | decompile | lookup | html2txt."""
    # Help pages carry non-ASCII entities, and a Korean console is cp949: without this the print aborts
    # mid-page with UnicodeEncodeError.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(prog="ds-docs", description="Decompile and search the official DraftSight help.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("status", help="what is available and what is generated")

    p = sub.add_parser("decompile", help="unpack the installed CHMs into docs/")
    p.add_argument("chm", nargs="*", help="CHM names (default: the shipped set)")
    p.add_argument("-a", "--all", action="store_true", dest="all_chms", help="every CHM in the help folder")
    p.add_argument("-f", "--force", action="store_true", help="re-extract even if already present")
    p.add_argument("--help-dir", help="override <install root>\\Help\\english")
    p.add_argument("--docs", help="output directory (default: %s)" % DEFAULT_DOCS)

    p = sub.add_parser("lookup", help="prose for an API member, enum family, or command name")
    p.add_argument("pattern")
    p.add_argument("-n", "--pages", type=int, default=1, help="max pages to print (default 1)")
    p.add_argument("-l", "--lines", type=int, default=70, help="max lines per page (default 70)")
    p.add_argument("-t", "--cmdref", action="store_true", help="command reference only (RunCommand names)")
    p.add_argument("-c", "--content", action="store_true", help="search page content, not page names")
    p.add_argument("-r", "--raw", action="store_true", help="print raw html instead of prose")
    p.add_argument("-s", "--scope", help="subtree of docs/ to search (e.g. draftsightapi)")
    p.add_argument("--docs", help="help tree (default: %s)" % DEFAULT_DOCS)

    p = sub.add_parser("html2txt", help="strip one decompiled page to prose")
    p.add_argument("path", nargs="+")
    p.add_argument("-l", "--lines", type=int, default=0, help="max lines (default: whole page)")

    a = ap.parse_args(argv)
    if a.cmd == "status":
        print(status())
        return 0
    if a.cmd == "decompile":
        rows = decompile(a.chm or None, all_chms=a.all_chms, help=a.help_dir, out_dir=a.docs, force=a.force)
        print(report(rows))
        return 0 if all(r["status"] in ("ok", "cached") for r in rows) and rows else 1
    if a.cmd == "lookup":
        out = lookup(a.pattern, pages=a.pages, max_lines=a.lines, cmdref=a.cmdref, scope=a.scope,
                     content=a.content, raw=a.raw, docs=a.docs)
        print(out)
        if out.startswith("docs/ is not generated"):
            return 3
        return 1 if out.startswith("no match") or out.startswith("no such scope") else 0
    for p in a.path:
        print(html2txt(p, a.lines))
    return 0


if __name__ == "__main__":
    raise SystemExit(cli())
