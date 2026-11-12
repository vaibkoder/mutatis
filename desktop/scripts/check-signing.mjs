/**
 * THE SIGNING AND HARDENING CHECK — phase 8.
 *
 * Everything in phase 8 is a claim about the artefact, not about the source, so
 * this reads the artefact: the .app bundle on disk, its Info.plist, its code
 * signature if it has one, and its entitlements if it is signed.
 *
 * It PASSES on an unsigned build. That is the point — an unsigned build is the
 * honest default, and the check's job is to describe the artefact accurately,
 * not to fail until someone buys a certificate. What it will not tolerate is a
 * MISDESCRIBED build: entitlements broader than declared, a notarisation claim
 * with no ticket, an update feed baked in by accident, or update machinery
 * wired to startup.
 *
 * Run:  node scripts/check-signing.mjs
 */
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { findPackagedApp } from './cdp.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const results = [];
let failures = 0;

function check(name, ok, detail = '') {
  results.push({ name, ok });
  if (!ok) failures += 1;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`);
}

function note(text) {
  console.log(`  ····  ${text}`);
}

/** Runs a tool and returns its combined output, or null if it failed. */
function run(file, args) {
  try {
    return execFileSync(file, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  } catch (error) {
    const out = `${error.stdout ?? ''}${error.stderr ?? ''}`;
    return out.length > 0 ? out : null;
  }
}

const APP = findPackagedApp(ROOT);
if (!APP) {
  console.error('No packaged app found. Run `npm run dist` first.');
  process.exit(2);
}
console.log(`Checking packaged build: ${path.relative(ROOT, APP)}\n`);

// ------------------------------------------------------------- the bundle ---

console.log('The bundle');

const plistPath = path.join(APP, 'Contents/Info.plist');
check('Info.plist exists', existsSync(plistPath));

function plistValue(key) {
  const out = run('/usr/libexec/PlistBuddy', ['-c', `Print :${key}`, plistPath]);
  // PlistBuddy reports a missing key on stderr with exit 1, and `run` hands back
  // that text rather than null. Treat it as absent, or every "this key is not
  // set" assertion passes its own error message off as a value.
  if (out === null || /Does Not Exist/.test(out)) return null;
  return out.trim();
}

check('the bundle identifier is ours', plistValue('CFBundleIdentifier') === 'com.mutatis.desktop',
  plistValue('CFBundleIdentifier') ?? 'missing');
check('the minimum system version is declared, not inherited',
  plistValue('LSMinimumSystemVersion') === '11.0', plistValue('LSMinimumSystemVersion') ?? 'missing');
// A packaged app that still declares a dev-server URL scheme, or that registers
// itself for arbitrary URLs, is a deep-link attack surface nobody asked for.
check('the app registers no URL scheme', plistValue('CFBundleURLTypes') === null,
  plistValue('CFBundleURLTypes') ?? 'none');

// ------------------------------------------------------------- entitlements --

console.log('\nEntitlements (what we asked to be granted)');

const ENTITLEMENT_FILES = [
  'build/entitlements.mac.plist',
  'build/entitlements.mac.inherit.plist',
];

// These are the ones @electron/osx-sign's defaults would hand out, or that are
// commonly copy-pasted in. Each is a real reduction in the process's hardening,
// and this app needs none of them.
const FORBIDDEN = [
  'com.apple.security.cs.allow-unsigned-executable-memory',
  'com.apple.security.cs.disable-library-validation',
  'com.apple.security.cs.allow-dyld-environment-variables',
  'com.apple.security.cs.debugger',
  'com.apple.security.device.audio-input',
  'com.apple.security.device.camera',
  'com.apple.security.device.bluetooth',
  'com.apple.security.device.usb',
  'com.apple.security.personal-information.location',
];

for (const file of ENTITLEMENT_FILES) {
  const full = path.join(ROOT, file);
  check(`${file} exists`, existsSync(full));
  if (!existsSync(full)) continue;
  // Strip XML comments first. Both plists NAME the entitlements they refuse to
  // grant and say why; a plain substring scan reads those explanations as
  // grants and fails the file for documenting itself.
  const text = readFileSync(full, 'utf8').replace(/<!--[\s\S]*?-->/g, '');
  check(`${file} grants allow-jit (V8 needs it)`,
    text.includes('<key>com.apple.security.cs.allow-jit</key>'));
  const granted = FORBIDDEN.filter((entitlement) => text.includes(`<key>${entitlement}</key>`));
  check(`${file} grants nothing this app does not need`,
    granted.length === 0, granted.join(', ') || 'none');
}

// --------------------------------------------------------- the signature ----

console.log('\nThe code signature');

const signaturePath = path.join(APP, 'Contents/_CodeSignature/CodeResources');
const signed = existsSync(signaturePath);
const codesign = run('codesign', ['-dvvv', '--entitlements', ':-', APP]) ?? '';

if (!signed) {
  check('unsigned build reports itself as unsigned', /code object is not signed/i.test(codesign)
    || codesign === '', codesign.split('\n')[0] ?? '');
  note('This build is UNSIGNED. Gatekeeper will refuse it on another Mac until the');
  note('user opens System Settings > Privacy & Security > Open Anyway. It also cannot');
  note('auto-update: Squirrel.Mac will not replace a bundle it cannot verify.');
  note('See docs/signing.md for what it would take to change that.');
} else {
  // Hardened Runtime shows up as the `runtime` flag in the CodeDirectory line.
  check('the hardened runtime is on', /flags=.*runtime/.test(codesign),
    (codesign.match(/flags=\S+/) ?? ['no flags line'])[0]);
  check('the signature carries a Team ID', /TeamIdentifier=(?!not set)\S+/.test(codesign),
    (codesign.match(/TeamIdentifier=\S+/) ?? ['TeamIdentifier=not set'])[0]);
  check('the signed entitlements grant nothing forbidden',
    !FORBIDDEN.some((entitlement) => codesign.includes(entitlement)));
  check('the signature verifies', run('codesign', ['--verify', '--deep', '--strict', APP]) !== null);

  // A notarised app carries a stapled ticket. Without one the app still needs a
  // network round trip to Apple on first launch, and fails closed if offline.
  const stapled = run('xcrun', ['stapler', 'validate', APP]);
  check('a notarisation ticket is stapled',
    stapled !== null && /validate action worked/i.test(stapled),
    stapled?.trim().split('\n').pop() ?? 'no ticket');
}

// ------------------------------------------------------------ the updater ---

console.log('\nThe updater');

const feed = path.join(APP, 'Contents/Resources/app-update.yml');
if (existsSync(feed)) {
  const text = readFileSync(feed, 'utf8');
  check('the update feed is https', /url:\s*https:\/\//.test(text), text.trim().replace(/\n/g, ' '));
} else {
  check('no update feed was baked in by accident', true, 'app-update.yml absent');
}

// The static half of the promise made in src/main/updater.ts. `main` is bundled
// by electron-vite, so this reads what actually shipped rather than the source.
const asarList = run(process.execPath, [
  path.join(ROOT, 'node_modules/@electron/asar/bin/asar.js'),
  'list',
  path.join(APP, 'Contents/Resources/app.asar'),
]) ?? '';
check('the packaged app contains a main bundle', asarList.includes('/out/main/index.js'));

// Read the SHIPPED main bundle, not the source. electron-vite may inline or
// rename anything, so a claim about "what main does at startup" has to be made
// against the bundle that is actually inside the asar.
const asarDir = mkdtempSync(path.join(tmpdir(), 'mutatis-sign-'));
let allMain = '';
try {
  execFileSync(
    process.execPath,
    [
      path.join(ROOT, 'node_modules/@electron/asar/bin/asar.js'),
      'extract',
      path.join(APP, 'Contents/Resources/app.asar'),
      asarDir,
    ],
    { stdio: 'ignore' },
  );
  const mainDir = path.join(asarDir, 'out/main');
  for (const entry of readdirSync(mainDir)) {
    if (entry.endsWith('.js')) allMain += readFileSync(path.join(mainDir, entry), 'utf8');
  }
} finally {
  /* cleaned at the end */
}

check('nothing calls checkForUpdatesAndNotify (that is the startup beacon)',
  !allMain.includes('checkForUpdatesAndNotify'));
check('automatic download is disabled', /autoDownload\s*=\s*false/.test(allMain));
check('automatic install on quit is disabled', /autoInstallOnAppQuit\s*=\s*false/.test(allMain));
check('electron-updater is imported lazily, not at startup',
  !/^import[^;\n]*['"]electron-updater['"]/m.test(allMain)
    && /\bimport\(\s*['"]electron-updater['"]\s*\)/.test(allMain));

rmSync(asarDir, { recursive: true, force: true });

// ------------------------------------------------------------------ done ----

console.log(`\n${results.length - failures}/${results.length} checks passed.`);
if (failures > 0) {
  console.log('\nFAILED:');
  for (const result of results) if (!result.ok) console.log(`  - ${result.name}`);
}
process.exit(failures > 0 ? 1 : 0);
