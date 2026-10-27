/**
 * draftsight-com bootstrap.
 *
 * One repo ships two artifacts: the pi skill (SKILL.md, references/) and the
 * Python package (src/draftsight_com). `pi install git:` only runs npm
 * installs, so the Python half has to land in the interpreter the agent
 * actually executes against. That interpreter is neither `python` on PATH nor
 * `py -3`, and being uv-created it usually has no `pip` of its own.
 *
 * So: resolve candidate interpreters deliberately, probe each one, install
 * with `uv pip --python <exe>` into the first candidate that is a real Python
 * but lacks the package, and never touch an interpreter we did not choose.
 */

import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const PKG_DIR = dirname(dirname(fileURLToPath(import.meta.url)));

/**
 * Import-time probe. `str()` around the spec is load-bearing: find_spec
 * returns None or a spec, and `'x' + None` raises TypeError, which would make
 * every healthy interpreter look broken.
 */
const PROBE = [
  "import importlib.util as u",
  "print('IMPORT ' + str(u.find_spec('draftsight_com') is not None))",
  "print('WIN32 ' + str(u.find_spec('win32com') is not None))",
].join("; ");

type Action =
  | "ready" | "installed" | "missing" | "incomplete" | "broken"
  | "absent" | "no-python" | "install-failed" | "unsupported";

type Setup = { ok: boolean; python?: string; action: Action; detail?: string; note?: string };

/** Actions that mean "this is a working Python we are allowed to install into". */
const INSTALLABLE: Action[] = ["missing", "incomplete", "broken"];

let settled: Setup | undefined;
const isWin = process.platform === "win32";
const venvPython = (venv: string) =>
  isWin ? join(venv, "Scripts", "python.exe") : join(venv, "bin", "python");

/**
 * An explicit pin is an instruction, not a hint. If the caller told us which
 * interpreter runs the session, that interpreter is the only target: a
 * fallback that already has the package does not help a kernel that does not,
 * which is precisely the failure this bootstrap exists to prevent.
 */
function pinned(): string | undefined {
  for (const name of ["DRAFTSIGHT_COM_PYTHON", "PI_PYTHON"]) {
    const v = process.env[name];
    if (v) return v;
  }
  return undefined;
}

/**
 * Discovered candidates, most likely to be the agent kernel first.
 * The pi kernel venv is a sibling of the agent config dir: PI_CODING_AGENT_DIR
 * defaults to ~/.pi/agent, so the venv is ~/.pi/.venv.
 */
function discovered(): string[] {
  const out: string[] = [];
  const agent = process.env.PI_CODING_AGENT_DIR || join(homedir(), ".pi", "agent");
  out.push(venvPython(join(agent, "..", ".venv")));
  out.push(venvPython(join(homedir(), ".pi", ".venv")));
  if (process.env.VIRTUAL_ENV) out.push(venvPython(process.env.VIRTUAL_ENV));
  return [...new Set(out.filter(Boolean))].filter(existsSync);
}

function lastLine(s: string): string {
  const keep = s.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  return keep.length ? keep[keep.length - 1] : "";
}

async function probe(pi: ExtensionAPI, python: string): Promise<Setup> {
  let r;
  try {
    r = await pi.exec(python, ["-c", PROBE], { timeout: 60_000 });
  } catch (e) {
    return { ok: false, python, action: "broken", detail: `spawn failed: ${e}` };
  }
  const out = `${r.stdout}\n${r.stderr}`;
  if (!out.includes("IMPORT ")) return { ok: false, python, action: "broken", detail: lastLine(out) || `exit ${r.code}` };
  if (out.includes("IMPORT True")) {
    if (!out.includes("WIN32 True")) {
      return { ok: false, python, action: "incomplete", detail: "draftsight_com imports but pywin32 is missing" };
    }
    return { ok: true, python, action: "ready" };
  }
  return { ok: false, python, action: "missing" };
}

async function ensure(pi: ExtensionAPI, force: boolean): Promise<Setup> {
  if (settled && !force) return settled;
  if (!isWin) {
    settled = { ok: false, action: "unsupported", detail: "COM automation is Windows only",
                note: "the reference docs are still readable" };
    return settled;
  }

  const pin = pinned();
  const pool = pin ? [pin] : discovered();
  let installable: string | undefined;
  for (const python of pool) {
    const st = await probe(pi, python);
    if (st.ok) { settled = st; return settled; }
    if (!installable && INSTALLABLE.includes(st.action)) installable = python;
  }

  if (!installable) {
    settled = { ok: false, action: "no-python",
                detail: pin
                  ? `pinned interpreter is not a usable Python: ${pin}`
                  : pool.length ? `no usable Python among: ${pool.join(" | ")}` : "no candidate interpreter exists on this machine" };
    return settled;
  }

  // uv owns most agent venvs and those venvs have no pip; fall back to -m pip only when uv is absent.
  const where = await pi.exec(isWin ? "where" : "which", ["uv"], { timeout: 30_000 }).catch(() => undefined);
  const hasUv = !!where && where.code === 0 && where.stdout.trim().length > 0;
  const run = hasUv
    ? await pi.exec("uv", ["pip", "install", "--python", installable, "--editable", PKG_DIR], { timeout: 900_000 })
    : await pi.exec(installable, ["-m", "pip", "install", "--editable", PKG_DIR], { timeout: 900_000 });
  const out = `${run.stdout}\n${run.stderr}`;
  if (run.code !== 0) {
    settled = { ok: false, python: installable, action: "install-failed", detail: lastLine(out) || `exit ${run.code}` };
    return settled;
  }
  const after = await probe(pi, installable);
  settled = after.ok
    ? { ok: true, python: installable, action: "installed", note: hasUv ? "via uv pip" : "via python -m pip" }
    : { ...after, detail: `install exited 0 but the import still fails: ${lastLine(out)}` };
  return settled;
}

function report(setup: Setup): string {
  const pin = pinned();
  const lines = [`action:  ${setup.action}`, `python:  ${setup.python ?? "(none resolved)"}`,
                 `pinned:  ${pin ?? "(discovered)"}`, `windows: ${isWin}`];
  if (setup.detail) lines.push(`detail:  ${setup.detail}`);
  if (setup.note) lines.push(`note:    ${setup.note}`);
  if (!setup.ok && setup.action !== "unsupported") {
    lines.push("", `manual:  uv pip install --python <python> --editable "${PKG_DIR}"`,
               "         (set DRAFTSIGHT_COM_PYTHON to pin the interpreter the agent runs)");
  }
  return lines.join("\n");
}

export default function (pi: ExtensionAPI) {
  pi.on("session_start", async (_event, ctx) => {
    const setup = await ensure(pi, false);
    if (setup.ok || setup.action === "unsupported") return;
    const text = report(setup);
    if (ctx.hasUI) {
      ctx.ui.notify(`draftsight-com: ${setup.action}`, "warning");
      ctx.ui.notify(text, "info");
    } else {
      // Print/RPC mode has no dialog surface; stderr is the only channel that survives.
      process.stderr.write(`draftsight-com setup: ${text.replace(/\n/g, " | ")}\n`);
    }
  });

  pi.registerCommand("draftsight-com", {
    description: "Show or repair the Python install behind the DraftSight skill",
    handler: async (args, ctx) => {
      const force = /(^|\s)(install|repair|force)(\s|$)/i.test(args ?? "");
      const setup = await ensure(pi, force);
      const text = report(setup);
      if (ctx.hasUI) ctx.ui.notify(text, setup.ok ? "info" : "warning");
      else process.stdout.write(`${text}\n`);
    },
  });
}
