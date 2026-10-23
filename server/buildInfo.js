import { execSync } from "child_process";
import { readFileSync, existsSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, "..");
const STAMP_FILE = join(here, "buildInfo.json");

function git(args) {
  try {
    return execSync(`git ${args}`, { cwd: repoRoot, stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
  } catch {
    return "";
  }
}

// Describes the code being built/run right now: package version, git commit, whether the
// working tree has uncommitted changes, and the build time.
export function computeBuildInfo() {
  let version = "0.0.0";
  try {
    version = JSON.parse(readFileSync(join(repoRoot, "package.json"), "utf8")).version || version;
  } catch {
    // keep default
  }
  return {
    version,
    commit: git("rev-parse --short HEAD") || "unknown",
    dirty: git("status --porcelain").length > 0,
    builtAt: new Date().toISOString(),
  };
}

// Deployed (Lambda) packages carry a stamp written by build-lambda-zip.cjs; local dev
// servers compute it from git at startup.
let cached = null;
export function getServerBuildInfo() {
  if (cached) return cached;
  if (existsSync(STAMP_FILE)) {
    try {
      cached = JSON.parse(readFileSync(STAMP_FILE, "utf8"));
      return cached;
    } catch {
      // fall through
    }
  }
  cached = { ...computeBuildInfo(), local: true };
  return cached;
}

export const BUILD_INFO_STAMP_FILE = STAMP_FILE;
