# Publishing Office Forge

Office Forge is a standard DSH bundle package. The manifest already declares:

- `dsh.bundle.patch: "./cordis.patch.yml"`
- `dsh.client: { "platform": "web" }`
- `exports["./client"]: "./lib/client.js"`
- `engines.dsh: "^0.2.2"`

## 1. npm

From a clean checkout:

```powershell
cd C:\022工作日志1\office-forge
npm pack --json
npm publish .\dsh-office-forge-<version>.tgz --access public
```

`npm publish` needs an npm token with publish rights for the `dsh-office-forge` package. The registry currently has `0.14.0`; the local workspace is at `3.5.9` and is local-only until this command is run. The same publish step runs in `.github/workflows/release.yml` when an `NPM_TOKEN` repository secret is configured.

## 2. GitHub release

1. Commit and tag the version (`git tag v<version>`).
2. Push the tag and let `.github/workflows/release.yml` build the tarball, or attach it manually to the GitHub Release.
3. The release asset must be named `dsh-office-forge-<version>.tgz`.

## 3. DSH plugin catalog

The DSH marketplace reads the curated list at `awesome-dsh-plugin/awesome-dsh-plugin`; it is not a directory to which plugins push directly.

1. Add `data/plugins/<owner>__dsh-office-forge.yml` with `name`, `owner`, `url`, `npm`, `version`, `description`, `category` and optional `icon`/`screenshots`.
2. Open a pull request against `awesome-dsh-plugin/awesome-dsh-plugin`.
3. The site and `dsh-plugin-catalog` index update after the repository's scheduled CI run.

The catalog package is also consumable directly:

```text
https://cdn.jsdelivr.net/npm/dsh-plugin-catalog/index.json
```

## 4. Offline / cross-machine package

```powershell
python scripts/build_offline.py --tarball .\dsh-office-forge-<version>.tgz --output .\dist\dsh-office-forge-<version>-offline.zip
```

The zip contains the npm tarball, `install.ps1`, `check.ps1`, `SHA256SUMS.txt` and a small README. It is intentionally dependency-honest: Python packages are bundled in the DSH runtime; OCR and LibreOffice remain optional host dependencies reported by `officeforge_doctor`.

## 5. Release checklist

- [ ] `tests/run_regression.py` PASS
- [ ] `tests/test_dsh_profile.mjs` PASS
- [ ] `npm pack` six metrics recorded in `OfficeForge_交付说明.md`
- [ ] `SHA256SUMS.txt` updated and `sha256sum -c` OK
- [ ] `cordis.patch.yml` contains no `officeforge` entry
- [ ] `@adwmc/helm-d` dependency tree intact
