# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.5.2] - 2026-10-02

### Added
- `requirements.txt` for `devices.py`: Pillow (JPG output) and selenium (needed by devon.py, which `devices.py` runs
  with the same Python).

## [0.5.1] - 2026-10-02

### Changed
- README now explains how to clone devon.py alongside the `devices` directory, install its requirements and check its
  dependencies, as it does for the visio-stencils repository.

## [0.5.0] - 2026-10-02

### Added
- `devices.py` now searches the `visio-stencils` directory layout (`<first letter>/<vendor>/<stencil>`) for a model's
  stencil when the built-in vendor/model rules have no master for it (unknown vendors and models, or models missing from
  the rule's stencil, e.g. HPE `DL380 Gen10`, Dell `R650`). It lists the masters in each of the vendor's stencils
  (cached in `svg-cache/_index`), picks the best front and rear masters, and extracts that stencil to SVGs on demand.
- `-nodiscover` switch to turn the search off, and `-maxscan N` to limit how many stencils are searched per model.
- A view with no master (for example a model that only has a front) is drawn as a blank plate stretched to the
  other view's size.

## [0.4.0] - 2026-10-02

### Added
- `devices.py`, a Python version of `devices.ps1` that draws the front and rear rack elevations from the CSV using
  SVGs extracted from the Visio stencils by devon.py, and writes SVG, PNG, JPG or PDF. Needs neither Windows nor Visio.
  Supports `-longracknames`, `-showlabels`, `-rackperfile` and `-pagelabels`, with the same vendor/model dispatch as
  the PowerShell script.
- `svg-cache/` (git ignored) holds the per-master SVGs `devices.py` extracts from each stencil on first use.
- README documentation for `devices.py`: requirements, usage, switches, output formats and known limitations.

### Changed
- `devices.py` shares the project version with `devices.ps1`, so both scripts and the README now report 0.4.0.

## [0.3.6] - 2026-10-02

### Fixed
- Older Dell rack servers (e.g. R820, R730, R630, C6320) are no longer in `Dell-PowerEdge-RackServers.vss`; they were
  moved to `Dell-PowerEdge-RackServers-Classic.vss`, so rows for them could not find a master. Both stencils are now
  registered as needed, with the current-generation models selected by `$dell_current_server_models` and everything
  else `^R|^C` going to the Classic stencil. The registration gates and shape dispatch use the same condition.
- The Dell `^R|^C` dispatch case no longer also matches `CX4` models (which have their own case).

## [0.3.5] - 2026-10-02

### Changed
- Stencil directory renamed from `stencils` to `visio-stencils`, so the
  https://github.com/lateralblast/visio-stencils repository can be cloned straight into the script directory.
- Stencil file paths now follow the `visio-stencils` repository layout (`<first letter>/<vendor>/<file>`, e.g.
  `visio-stencils\d\dell\Dell-Racks.vss`).

### Fixed
- `unzip_stencil` now also looks for a zip named without the stencil extension (the repository only has
  `Dell-PowerEdge-RackServers.zip` for the classic `.vss`), and skips the `__MACOSX` folder some zips contain.

## [0.3.4] - 2026-09-28

### Fixed
- Oracle/Sun stencil registration was gated on `Model` prefix regexes (e.g. `^T[0-9]|^M[0-9]|^E[0-9]` for SPARC),
  but the per-row dispatch switch that actually picks the master shape decides SPARC vs. Intel from the
  `Architecture` column instead. A row with a SPARC `Architecture` but a `Model` that didn't match those prefixes
  would never get its SPARC stencil registered, then fail at dispatch time trying to use it. The registration
  gates now mirror the dispatch switch's logic exactly (`Model` contains `Blade` -> blade, else `Architecture`
  SPARC/sparc -> SPARC, else -> Intel).

## [0.3.3] - 2026-09-27

### Changed
- Factored the repeated shape text/formatting block (label text, `Char.Color`, `Char.Size`, `TxtLocPinX/Y`,
  `TextBkgnd`, `HideText`), previously duplicated four times across the rack front/back and device front/back
  placement code, into a single `Set-ShapeDefaults` helper function.
- Commented out the unused `$script_path` and `$data_dir` variable assignments.

## [0.3.2] - 2026-09-27

### Fixed
- Rack names taken directly from the CSV `Rack` column could contain characters invalid in Windows filenames
  (`/ \ : * ? " < > |`); `-rackperfile` now sanitizes them before building the output filename (the Visio page
  name itself is left unsanitized).
- `-outputfile` path detection used to check for a literal `:` to decide whether a path was absolute, which
  misclassified UNC paths (`\\server\share\...`) as relative. Now uses `[System.IO.Path]::IsPathRooted`.
- Missing vendor stencil files (and no `.zip` to extract them from) previously failed with a raw COM error from
  deep inside `unzip_stencil`. Now fails with a clear message pointing at the README's stencil-sourcing docs.

## [0.3.1] - 2026-09-27

### Fixed
- Vendor stencils were only registered based on the first rack's rows in single-document mode, so a later rack
  introducing a vendor/model not present in the first rack would fail to place its shape. Stencil registration is
  now tracked per stencil instead of gated by a single one-shot flag.
- The Dell model dispatch had no fallback case, so a Dell row with an unrecognized model prefix silently reused
  whatever shape had last been registered under the same name (potentially from a different vendor's row) instead
  of falling back to a blank plate.
- `get_file_type` computed the MIME type of the running script instead of the input file, so Excel-format detection
  never reflected the actual input file.

## [0.3.0] - 2018-02-11

### Changed
- Improved page label handling.

## [0.2.9] - 2018-02-11

### Added
- `-pagelabels` switch to put rack label on top of page.

## [0.2.8] - 2018-02-10

### Added
- Additional examples and code to output a file per rack.

## [0.2.7] - 2018-02-10

### Fixed
- Ensure server stencils are on top of rack stencils.

## [0.2.6] - 2018-02-10

### Added
- Code to show labels on stencils.

## [0.2.5] - 2018-02-09

### Fixed
- Calls to code to unzip stencils.

## [0.2.4] - 2018-02-09

### Added
- `-longrackname` switch option to include hostnames in page name and rack name.

## [0.2.3] - 2018-02-09

### Fixed
- Shape labels and added additional shape data.

## [0.2.2] - 2018-02-08

### Changed
- Example and shape data updates.

## [0.2.1] - 2018-02-08

### Added
- Pure FlashBlade support, updated example.

## [0.2.0] - 2018-02-07

### Changed
- Updated code to support modifying shape data.

## [0.1.9] - 2018-02-05

### Changed
- Improved outputfile handling.

## [0.1.8] - 2018-02-05

### Changed
- Rewrite to use VisioBot3000 PowerShell module.

## [0.1.7] - 2016-10-31

### Fixed
- More bug fixes.

## [0.1.6] - 2016-10-30

### Fixed
- Bug with processing racks with one item.

## [0.1.5] - 2016-10-30

### Fixed
- Various bugs.

## [0.1.4] - 2016-10-30

### Added
- Code to unzip stencils file.

## [0.1.3] - 2016-10-29

### Added
- Support for NetApp storage.

## [0.1.2] - 2016-10-29

### Added
- Support for IBM servers.

## [0.1.1] - 2016-10-29

### Added
- Support for Dell storage.

## [0.1.0] - 2016-10-27

### Added
- Support for Dell rackmount servers.

## [0.0.9] - 2016-10-27

### Added
- Code to create an output file for each rack within a list.

## [0.0.8] - 2016-10-27

### Fixed
- Blanking plate placement.

## [0.0.7] - 2016-10-27

### Added
- Code to put in blanking plates if name can't be resolved.

## [0.0.6] - 2016-10-27

### Fixed
- Numerous fixes.

## [0.0.5]

### Added
- Code to determine location on page.

## [0.0.4]

### Added
- Code to load CSV file.

## [0.0.3]

### Added
- Code to test type of files.

## [0.0.2]

### Added
- Code to import stencils.
- Code to select server stencils.

## [0.0.1]

### Added
- Initial version.
