import { copyFileSync, existsSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..", "..");

function candidatePaths() {
  const envPath = process.env.OFFICEMD_JS_ADDON;
  const candidates = [
    envPath,
    resolve(repoRoot, "target", "debug", "deps", "libofficemd_js.dylib"),
    resolve(repoRoot, "target", "release", "deps", "libofficemd_js.dylib"),
    resolve(repoRoot, "target", "debug", "deps", "libofficemd_js.so"),
    resolve(repoRoot, "target", "release", "deps", "libofficemd_js.so"),
    resolve(repoRoot, "target", "debug", "deps", "officemd_js.dll"),
    resolve(repoRoot, "target", "release", "deps", "officemd_js.dll"),
    resolve(repoRoot, "target", "debug", "deps", "officemd_js.node"),
    resolve(repoRoot, "target", "release", "deps", "officemd_js.node"),
  ];

  return candidates.filter((p) => typeof p === "string" && p.length > 0);
}

function firstExistingPath(paths) {
  for (const p of paths) {
    if (existsSync(p)) {
      return p;
    }
  }
  return null;
}

function ensureNodeBinary(path) {
  if (path.endsWith(".node")) {
    return path;
  }

  const outputPath = resolve(dirname(path), "officemd_js.node");
  copyFileSync(path, outputPath);
  return outputPath;
}

export function loadNativeBinding() {
  const candidates = candidatePaths();
  const found = firstExistingPath(candidates);

  if (!found) {
    const searchList = candidates.map((p) => `  - ${p}`).join("\n");
    throw new Error(
      "Could not find the OfficeMD native addon.\n" +
        "Build it first with: cargo build -p officemd_js\n" +
        "Searched:\n" +
        searchList,
    );
  }

  const addonPath = ensureNodeBinary(found);
  return require(addonPath);
}
