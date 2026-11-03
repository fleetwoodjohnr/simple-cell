# Build and release

How to build, test and release Simple Cell.
- **Linux** (Fedora 44, x86-64): steps marked **verified** were run there.
- **Windows and macOS:** the CI workflow ([.github/workflows/ci.yml](../.github/workflows/ci.yml)) is written to build and test them on every push. It **has not run yet**, because the repository has no GitHub remote.

## Requirements

- **Node 24** and npm. The Electron 45 runtime bundles its own Node.
- **For the tests (optional):**
  - poppler (`pdftotext`, `pdfinfo`) for the PDF tests.
  - LibreOffice (`soffice`) for the LibreOffice re-read and cross-check tests, and for .xls/.xlsb files.
  - xvfb on headless Linux.

  Tests that need a missing tool skip themselves.

## Develop and test (verified on Linux)

```bash
npm ci
npm run dev                # the app with hot reload
npm run typecheck          # strict TypeScript, main process and renderer
npm test                   # unit tests (formulas, engine, file formats, security)
npm run e2e:headless       # build, then Playwright on Electron under xvfb; every test also fails on a CSP violation
npm run crosscheck         # every function-test formula also calculated by LibreOffice, compared
npm run bench              # performance benchmarks → test-results/bench.md (see PERFORMANCE_BENCHMARKS.md)
SC_PERF=1 npm run e2e:headless -- perf   # scrolling frame times in the app
```

After changing commands, shortcuts or functions, regenerate the generated docs and commit them:

```bash
UPDATE_DOCS=1 npx vitest run tests/unit/docs.test.ts
```

## Build

```bash
npm run dist                       # build, then electron-builder for the current OS (targets in electron-builder.yml)
npx electron-builder --linux dir   # unpacked app only: the quickest for testing
```

Output goes to `release/`.

| OS | Installers |
|---|---|
| Linux | `Simple Cell-<v>.AppImage`, `simple-cell_<v>_amd64.deb`, `simple-cell-<v>.x86_64.rpm` |
| Windows | `Simple Cell Setup <v>.exe` (NSIS, x64) |
| macOS | `.dmg` and `.zip` for arm64 and x64 |

**Signing.** Builds are unsigned, so Windows SmartScreen and macOS Gatekeeper warn on first launch.
- macOS builds are signed ad hoc (`mac.identity: "-"`), because Apple Silicon won't run unsigned code.
- To sign later, add the certificates as repository secrets (`CSC_LINK`, `CSC_KEY_PASSWORD`, and Apple's notarization variables), then remove the ad-hoc lines.

**deb and rpm on a host without fpm's libraries.** On Fedora, electron-builder's bundled fpm needs `libcrypt.so.1`, and rpm needs `rpmbuild`. Instead of installing them, package the unpacked app in a throwaway container:

```bash
npx electron-builder --linux dir
podman run --rm --security-opt label=disable -v "$PWD":/src -v "$HOME/.cache/electron-builder":/root/.cache/electron-builder -w /src \
  registry.fedoraproject.org/fedora:44 bash -c \
  "dnf install -y -q nodejs24 libxcrypt-compat rpm-build binutils && node-24 node_modules/.bin/electron-builder --linux deb rpm --prepackaged release/linux-unpacked"
```

## Test the packaged app

```bash
# the whole E2E suite against the packaged binary (also runs the updater test, which needs a packaged build)
SC_PACKAGED=$PWD/release/linux-unpacked/simple-cell xvfb-run -a npx playwright test
```

```bash
# the AppImage itself (no FUSE needed)
APPIMAGE_EXTRACT_AND_RUN=1 SC_PACKAGED="$PWD/release/Simple Cell-<version>.AppImage" xvfb-run -a npx playwright test smoke files macros print update
# deb / rpm: install into clean containers and check that every shared library resolves
podman run --rm -v "$PWD/release":/rel:ro docker.io/library/ubuntu:24.04 bash -c \
  'apt-get update -qq && apt-get install -y /rel/simple-cell_<version>_amd64.deb && ldd "/opt/Simple Cell/simple-cell" | grep "not found"; true'
podman run --rm -v "$PWD/release":/rel:ro registry.fedoraproject.org/fedora:44 bash -c \
  'dnf install -y /rel/simple-cell-<version>.x86_64.rpm /usr/bin/ldd && ldd "/opt/Simple Cell/simple-cell" | grep "not found"; true'
```

**Results on 2026-11-03 (verified on Linux), version 0.1.0:**
- The AppImage (131 MB), deb (105 MB) and rpm (94 MB) all built.
- **`linux-unpacked`:** all 24 E2E tests pass, including the updater. The scroll-timing test is skipped unless `SC_PERF` is set.
- **AppImage:** the smoke, files, macros, print and updater tests pass (6 tests).
- **deb:** installs on clean Ubuntu 22.04 and 24.04 with every library resolved. Ubuntu 26.04 is in CI only.
- **rpm:** installs on clean Fedora 44 with every library resolved.
- **Not yet started from a desktop menu:** the installed deb and rpm builds. File associations have not been tried either.

## Updates

Installed copies check the newest **published** GitHub release of `fleetwoodjohnr/simple-cell` (the `publish` block in `electron-builder.yml`). They check at most once a day, and never if the user turns it off in Settings → Updates.
- **Asks first.** Nothing is downloaded until the user clicks Install. The download is checked against the sha512 in the release's `latest*.yml`.
- **macOS** opens the release page instead, because only apps with a Developer ID signature can update themselves.
- **Test:** `tests/e2e/update.spec.ts`, against the packaged app. It serves a 99.0.0 release from a local feed and checks three things: the prompt, the download (verified against the sha512), and the restart prompt.
  - On Linux the test sets `APPIMAGE` to a stand-in file, because electron-updater's AppImage updater (the default for an unpacked build) is inactive otherwise.
  - The final install step (installer, `pkexec`) is not exercised. Check it by hand when the second release ships.
- **No published release exists yet,** so no installed copy has updated for real.

## Release

1. Bump `version` in `package.json`, turn the top `CHANGELOG.md` section into that version, and update the date line in `IMPLEMENTATION_STATUS.md`.
2. Push a tag `v<version>`. CI builds and tests on all three OSes and installs the deb and rpm into clean Ubuntu and Fedora containers. If that passes, it creates a **draft** release with every installer and the changelog section as notes.
3. Check by hand before publishing:
   - Open files from Excel and LibreOffice, save them, and reopen the results in both.
   - Print a page and export a PDF.
   - Run a macro.
   - Update from the previous release.
   - With the update check off, confirm the app makes no network traffic.
4. Publish the draft. Publishing is what ships the update to installed copies.

None of this has been done yet: there is no remote, and nothing has been pushed.
