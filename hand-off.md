# Hand-off: `draftsight-com` as a `pi install git:` package

Repo now lives at `D:\dev-pi\draftsight-com`. It used to sit inside
`~/.pi/agent/skills/`, pi's auto-scanned personal skills dir, which is exactly the conflict this
document set out to remove. The move is done and the extension works. **The blocking regression
recorded by earlier revisions of this file was not a pi bug — it was a Python bug in our own
probe.** Read [False alarm](#false-alarm-what-the-blocking-regression-actually-was) before
believing anything else written here.

---

## Ledger

| Check | Status | Evidence |
|---|---|---|
| Repo moved to `D:\dev-pi\draftsight-com` | **verified** | `git status` clean-apart-from-known-files at the new path; `.git` intact, `origin/main` reachable |
| Kernel imports from the new path | **verified** | `draftsight_com.__file__` = `D:\dev-pi\draftsight-com\src\draftsight_com\__init__.py`, `sys.executable` = `~\.pi\.venv` |
| Editable install re-pointed | **verified** | `.pth` now contains `D:\dev-pi\draftsight-com\src`; `findstr` finds no old path in any `.pth`/`.json` under `~\.pi` |
| Old skill dir no longer discoverable | **verified** | contents deleted; only a locked empty shell remained (no `SKILL.md`), removal retried by a detached helper |
| `package.json` makes pi treat the repo as a package | **verified** | isolated `pi install <path>` exited 0 and wrote the entry |
| `"pi": { "skills": ["."] }` discovers root `SKILL.md` | **verified** | model turn answered `SKILLS draftsight-com`; command list shows `/skill:draftsight-com` |
| Extension loads from the package | **verified** | factory log line appears on every run; `session_start` fires (`hasUI=false mode=print`) |
| `/draftsight-com` is registered | **verified** | `pi --print "/draftsight-com"` printed `action: ready / python: ~\.pi\.venv\Scripts\python.exe / windows: true` |
| Extension installs the Python half | **verified** | forced run: `install exit=0 tail=+ pywin32==312`, then `import draftsight_com, win32com, pythoncom` OK in the previously empty venv |
| Extension skips install when a candidate is healthy | **verified** | probe log `-> missing` then `-> ready`, no install attempted |
| Docs regenerate at the new path | **verified** | `ds-docs decompile` -> `total 6457 pages, 63.5 MB -> D:\dev-pi\draftsight-com\docs`; `lookup basepoint` returns real API prose |
| README claims match behaviour | **verified** | its install description is the behaviour now proven above |
| Pin is exclusive, not a hint | **verified** | pin to an empty venv: `installed`, then `import draftsight_com, win32com, pythoncom` OK *in that venv*; report line `pinned: <path>` |
| Discovery still resolves the kernel | **verified** | no pin -> `action: ready`, `python: ~\.pi\.venv\Scripts\python.exe`, `pinned: (discovered)`, nothing reinstalled |
| Bogus pin fails loudly, never falls back | **verified** | pin=`notepad.exe` -> `action: install-failed`, `python:`/`pinned:` both name it, stderr line + manual command. So `no-python` is a **discovery-only** action; a pin that is not a Python reads `install-failed` |
| `pi install git:github.com/3DAlgoLab/draftsight-com` from a fresh agent dir | **verified** | fresh agent dir, `PI_OFFLINE=1`: exit 0, `settings.json` = `["git:github.com/3DAlgoLab/draftsight-com"]`, clone at `6d5f3c8`; `/draftsight-com` with no pin -> `action: ready`, `python: ~\.pi\.venv\Scripts\python.exe`, `pinned: (discovered)` |
| Same install drives a *pinned* empty venv | **verified** | pinned a venv created seconds earlier: `_editable_impl_prime_agent_skill_draftsight_com.pth` -> `<clone>\src` plus `pywin32-312.dist-info` appeared in its site-packages, and that venv printed `IMPORTS OK` with `__file__` inside the clone |
| Live config untouched by the tests | **verified** | `~/.pi/agent/settings.json` has zero `draftsight` matches and still ends at `npm:pi-web-access` |
| `/draftsight-com install` right after a session-start install | **known nuance** | `force` re-probes, so the report honestly reads `ready`, not `installed`; `installed` is only returned by the pass that actually installed. Judge an install by site-packages and `__file__`, not by that word |

---

## False alarm: what the "blocking regression" actually was

The symptom was real, the diagnosis was wrong. `CMD: no` came from two independent things:

**1. A TypeError in the probe, not a load failure.** The probe was written as

```python
print('IMPORT ' + (u.find_spec('draftsight_com') is not None))   # str + bool -> TypeError
```

In Python that raises, so *every* interpreter exited 1, every candidate classified `unreachable`,
nothing was installable, and `ensure()` returned `no-python`. The fix is `str(...)`, and the
classifier now refuses to trust output that lacks `IMPORT ` at all, so a broken probe reads as
`broken` (still installable) instead of poisoning the whole pool.

**2. The model cannot see extension commands in its listed command set in print mode.** Given
"list every slash command", the model correctly enumerated `/settings /model /compact ...
/skill:draftsight-com` and omitted `/draftsight-com` — while `pi --print "/draftsight-com"`
executed it and printed its report. So *asking the model whether a command exists* is not a
registration test in `--print` mode. **Invoke it instead.** Earlier revisions of this file called
that answer evidence of a discovery bug; it was an artifact of the probe question.

Nothing about pi's package discovery was at fault: `pi.skills: ["."]` coexisting with the
conventional `extensions/` dir is fine, and the `skills/draftsight-com/` fallback layout
suggested here earlier is **not needed**.

---

## Repo state

Uncommitted at `D:\dev-pi\draftsight-com` (nothing staged, `main` level with `origin/main`):

```
 M SKILL.md              setup section: install into the interpreter that runs the session
 M pyproject.toml        readme ref, `pywin32; sys_platform == 'win32'`, [docs] extra, sdist include/exclude
?? package.json          pi manifest: skills ["."], extensions convention, scripts ds-docs
?? README.md             intro, both install paths, "How the install works", troubleshooting
?? extensions/index.ts   interpreter bootstrap + /draftsight-com
?? hand-off.md           this file
```

`docs/` (63.5 MB, generated) and `node_modules/` (written by `pi install`) are gitignored — do not
commit either. Pushed history: `fc3bbe9` circle/arc entities, `7e2316b` merged `knowledge/` into
`references/`.

---

## Facts verified about pi 1.4.0

- Personal skills are discovered by **recursive** scan of `~/.pi/agent/skills`. A package placed
  in that tree is discovered twice; the resolution is first-wins plus a warning, not a merge.
- Package resource discovery honours `.gitignore` / `.ignore` / `.fdignore` as ignore manifests.
- `pi.exec(cmd, args, opts)` resolves to `{ code, stdout, stderr, killed }`; a spawn failure
  **throws** rather than returning a nonzero code. Catch it or one bad candidate looks like a
  healthy-but-broken Python.
- `ctx.ui.notify` accepts `"info" | "warning" | "error"` — there is no `"warn"`.
- In `--print` mode there is no dialog surface: `ctx.hasUI` is false and `ctx.mode === "print"`.
  `ctx.ui.notify` disappears. **stderr survives**, which is where the command report goes.
- A relative path handed to `pi install` is written verbatim into `settings.json`
  (`"../../../.pi/agent/skills/draftsight-com"`), and later `pi remove` from another CWD will not
  match it. Always install by git URL or absolute path.
- `pi` is a `.cmd` shim: not exec-able without a shell, and it truncates a prompt at the first
  newline. Keep probe prompts on one line and assert they arrived.
- Console codepage here is cp949 while pi emits box-drawing characters: capture child stdout as
  UTF-8 or die in `UnicodeDecodeError`.
- Git Bash rewrites a bare `/draftsight-com` argument into `C:/Program Files/Git/draftsight-com`,
  so pi receives a *path* and answers it as a question — which silently spends a model turn instead
  of running the command. Scope the escape to the argument
  (`MSYS2_ARG_CONV_EXCL='/draftsight-com'`); the blunt `MSYS_NO_PATHCONV=1` breaks the `pi` shim
  itself, whose own `/c/...` launcher path then reaches node unconverted.
- Probing a pinned path *executes* it (`<exe> -c <probe>`). Pin something that is not a Python and
  you launch it: `notepad.exe` blocked the probe on a GUI process until it was killed, then `uv`
  refused it as an interpreter. A bogus pin therefore reads `install-failed`, never `no-python`.
- Custom providers come from `agent/models.json`; copy it into an isolated agent dir or
  `dpi1/coder1` silently stops existing. `--model dpi1/coder1:off` is rejected.

---

## Facts verified about the local Python layout

| Interpreter | What it is | pip | holds the package |
|---|---|---|---|
| `~\.pi\.venv\Scripts\python.exe` | the pi/fabric kernel, uv-created | **no** | yes, editable via `.pth` |
| `python` on PATH | uv-managed 3.12 | yes | no |
| `py -3` | 3.14 | yes | no |
| temp test venvs | uv-created | **no** | only after bootstrap |

`~\.pi\.venv\Lib\site-packages\_editable_impl_*.pth` holds a **literal path**. Moving the repo
without re-running `uv pip install --editable` breaks `import draftsight_com` in the kernel
silently. That reinstall is step 2 of any future move, non-negotiable.

---

## Extension design as implemented (`extensions/index.ts`)

Two modes, decided by one question: did the caller pin an interpreter? `$DRAFTSIGHT_COM_PYTHON`,
else `$PI_PYTHON` — when either is set it is the **sole** target, nothing else is probed or
installed into, because a fallback that already holds the package hides the kernel that does not.
Otherwise discovery: `$PI_CODING_AGENT_DIR/../.venv` → `~/.pi/.venv` → `$VIRTUAL_ENV`, deduped,
existing-only, first `ready` wins. `report()` prints `pinned:` as either the path or `(discovered)`,
so a log line always says which mode answered.

On `session_start`: probe each candidate with `importlib.util.find_spec` (imports nothing from the
target, so nothing of ours runs inside it), stop at the first `ready`; otherwise remember the first
target in `{missing, incomplete, broken}` and `uv pip install --python <exe> --editable
<PKG_DIR>` (900 s), falling back to `<exe> -m pip install -e` only when `where uv` fails, then
re-probe to confirm. Memoised in `settled`; `/draftsight-com install|repair|force` bypasses it.
Non-Windows short-circuits to `unsupported` with a note that the reference docs still read fine.

Failures surface, they do not hide: `hasUI` gets two notifications, print/RPC mode gets one
stderr line with ` | ` separators, and the report always carries the manual `uv pip install`
command with the resolved `PKG_DIR`.

Constraints held on purpose: never install into an unresolved interpreter, never reinstall a
healthy one, never import the target package during a probe.

---

## Isolated test recipe (nothing of yours is touched)

```powershell
$T = "$env:TEMP\pi-pkg-test"
$env:PI_CODING_AGENT_DIR = "$T\agent"; $env:PI_OFFLINE = "1"; $env:PI_SKIP_VERSION_CHECK = "1"
$env:DRAFTSIGHT_COM_PYTHON = "$T\venv\Scripts\python.exe"
copy ~\.pi\agent\{models.json,auth.json,models-store.json} "$T\agent\"
uv venv --python 3.12 "$T\venv"          # empty venv: the interpreter under test
pi install "<repo>"; pi --print --no-session --offline --model dpi1/coder1 "say ok"
pi --print --no-session --offline --model dpi1/coder1 "/draftsight-com"   # this is the real test
```

The install branch needs a run with **no healthy fallback**, otherwise the kernel venv answers
`ready` and no install is attempted (correct behaviour, but it tests nothing). Since the pin became
exclusive, **the pin *is* the no-fallback switch** — `DRAFTSIGHT_COM_PYTHON` alone is enough, no
fake home required:

```powershell
$env:DRAFTSIGHT_COM_PYTHON = "$T\venv\Scripts\python.exe"   # sole target; nothing else is probed
```

The earlier configuration (`$env:HOME = "$T\fakehome"; $env:USERPROFILE = "$T\fakehome"`, agent dir
inside it) is still the way to test *discovery* with no kernel venv in reach; it is what produced
`install exit=0 tail=+ pywin32==312`.

---

## Next

All three original steps are done: pushed (`082f7eb`, `5458e9c`, then the pin rework `6d5f3c8`),
the `git:` install passes from a fresh isolated agent dir, and the live `settings.json` still ends
at `npm:pi-web-access` with no draftsight entry. What is actually left:

1. Nothing blocks the package. Test it on a *second* machine, where no `~/.pi/.venv` fallback
   exists, because that is the only shape still unexercised end to end (the pin stands in for it
   here, and it did install).
2. Decide whether `installed` should survive a forced re-probe (see the nuance row above). Current
   behaviour is truthful about state; it just loses the word that a repair happened.
3. Consider refusing to probe a pinned path that is not an executable Python-looking name: pinning
   `notepad.exe` launched Notepad before `uv` rejected it (see facts above).

## Open decisions

1. Keep the bootstrap extension at all? It is now proven to work, but "tell the user to run one
   `uv pip install` line" remains simpler and has no silent-failure surface.
2. `package.json` name `pi-draftsight-com` vs Python dist `prime-agent-skill-draftsight-com`: two
   names for one artifact, kept apart only because renaming the dist orphans installed dist-info.
3. Should `pi update --extensions` re-run the Python install? Today it does not.
4. No LICENSE file exists, so `package.json` declares none.
