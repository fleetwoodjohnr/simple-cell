# Build and release

How to build, test and release Simple Cell.
- **Linux** (Fedora 44, x86-64): steps marked **verified** were run there.
- **Windows and macOS:** the CI workflow ([.github/workflows/ci.yml](../.github/workflows/ci.yml)) builds and tests them on GitHub on every push. On CI a failed end-to-end test is retried twice; one that passes on a retry is reported as flaky.

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
npm run build && SC_SCREENSHOTS=1 xvfb-run -a -s "-screen 0 1920x1080x24" npx playwright test screenshots   # README screenshots → docs/screenshots
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

1. Set the version with `npm version <version> --no-git-tag-version` (this keeps `package.json` and `package-lock.json` in sync), turn the top `CHANGELOG.md` section into that version, and update the date line in `IMPLEMENTATION_STATUS.md`. Use a stable version such as `0.1.0`; the workflow checks the `latest*.yml` update feeds.
2. Commit the release changes, then create and push its tag:
   ```bash
   git tag v<version>
   git push origin HEAD
   git push origin v<version>
   ```
   The tag must already exist and match `package.json`; the preflight job rejects mismatches before building. All jobs use the same resolved commit. CI builds and tests on all three OSes and installs the deb and rpm into clean Ubuntu and Fedora containers.
3. After the checks pass, CI validates **all eight installers**: Windows x64 NSIS; macOS Intel and Apple Silicon DMG and ZIP; Linux x64 AppImage, deb and rpm. It checks the generated blockmaps and all three updater manifests, including every referenced file's version, size and SHA-512 hash. Filenames use dashes instead of spaces, matching the manifests and README. Only these validated assets are uploaded to a **draft** release, with the newest changelog section as notes. The workflow summary links to the draft and lists its assets.
4. Check by hand before publishing:
   - Open files from Excel and LibreOffice, save them, and reopen the results in both.
   - Print a page and export a PDF.
   - Run a macro.
   - Update from the previous release.
   - With the update check off, confirm the app makes no network traffic.
5. Wait for the last workflow run for the tag to finish, then publish the draft in GitHub. **Draft downloads are not public.** Publishing makes the README's latest-release link work and ships the update to installed copies. Do not publish while an upload or retry is running.

### Rebuild or retry a draft

Rerun a failed tag workflow, or use **Actions → CI → Run workflow** with the existing version in the **tag** input. The CLI equivalent is:

```bash
gh workflow run ci.yml --ref main -f tag=v<version>
```

A blank tag input runs ordinary CI without creating a release (unless the selected workflow ref is itself a version tag). Manual rebuilds check out the supplied tag, not the selected branch. Tag pushes and manual rebuilds for the same tag share a concurrency group.

Retries replace matching assets and refresh the notes on an existing draft, so an interrupted upload can be resumed. The workflow refuses to modify a published release and rechecks draft status before uploads and before editing notes. Keep publishing and retries separate: publish only after the last run completes. Never move or reuse a version tag; the workflow checks the remote tag against the built commit again before drafting. Release a new version to change public downloads. Builds stay unsigned on Windows and signed ad hoc on macOS; publishing a draft does not add a certificate or notarization.

No release has been published yet.
