# Performance benchmarks

Measured on 2026-11-03.
- **Machine:** AMD Ryzen 3 3200U (2 cores, 4 threads, 2019 laptop CPU), 13 GB RAM, Fedora Linux, Node 24 / Electron 45.
- **Repeat:** `npm run bench` for the core, and `SC_PERF=1 npm run e2e:headless -- perf` for scrolling.

The core figures time the calculation and file code directly. Electron is involved only in the scroll test.

| Measure | Result | Target | |
|---|---|---|---|
| Open a 100,000-row × 4-column XLSX (numbers, text, dates) | 1,288 ms | < 3,000 ms | ✅ |
| Save it as XLSX | 1,430 ms | — | |
| Full recalculation of a 100,000-cell dependency chain (`A2=A1+1` … `A100000`) | 422–482 ms (3 runs) | < 500 ms | ✅ |
| Edit the head of that chain (all 100,000 dependents recalculate) | 407–440 ms (3 runs) | < 500 ms | ✅ |
| Type into a cell that two whole-column totals (`SUM(A:A)`, `AVERAGE(A1:A100000)`) depend on, core time to the screen patch (median of 50) | 33 ms | < 50 ms | ✅ |
| Scroll a filled sheet (20,000 × 8) with the mouse wheel: frame time | median 16.7 ms, p95 16.7 ms, max 33.3 ms | frames at the display rate | ✅ at least 95% of frames on time; the slowest took two frames |
| Heap for 1,000,000 numeric cells (100,000 × 10) | 60 MB | < 500 MB | ✅ |

## Notes

- **Recalculation was quadratic until this round.** The same chain took 35 s before the fix: finding the dirty cells a formula reads scanned every dirty row in the column for each formula. It now looks up the formula's own rows when there are fewer of them. Doubling the chain now roughly doubles the time:
  - 5,000 cells: 47 ms
  - 10,000 cells: 57 ms
  - 20,000 cells: 132 ms
- **The chain meets its target since 2026-11-04** (it was ~10% over). Each formula's evaluation no longer allocates closures or reference objects: the evaluation context shares its callbacks, the dependency ordering keeps its stack in flat arrays, dirty propagation reuses one queue, and arithmetic on a single cell reads it directly. The margin is small (up to 482 ms on a cold first run), so the result varies with the machine's load.
- **Scrolling** was measured under xvfb with software rendering, with frames capped at 60 Hz. 16.7 ms is one frame at 60 Hz, so at least 95% of frames were on time. The test doesn't measure a GPU or a 120 Hz display.
- **Not measured yet:**
  - Opening and calculating real-world workbooks with heavy formulas. No such fixture is bundled, and none will be made from users' files.
  - Paint time for the *edit → frame* path in the app. The figure above stops at the screen patch the core sends.
  - Memory in the app as a whole, beyond the core's heap.
  - Windows and macOS.
