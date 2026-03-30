# Performance Log

Test document:
`/home/yuche/Downloads/The Routledge Handbook of Translation and Philosophy.docx`

Primary browser:
`/home/yuche/.nix-profile/bin/google-chrome-stable`

Metrics captured per run:
- `parseMs`: `docx.parseAsync(...)`
- `transferMs`: dedicated worker-to-main snapshot transfer time
- `renderMs`: `docx.renderDocument(...)`
- `totalMs`: parse + render + 2 animation frames
- `elements`: rendered element count inside benchmark root
- `textNodes`: rendered text node count inside benchmark root
- `pages`: rendered page count
- `visiblePages`: visible page count after first paint
- `longTasks`: count / sum / max from Long Tasks API
- `scrollPassMs`: time for one automated scroll pass
- `scrollPeakElements`: peak rendered element count during scroll

## Baseline

Date:
`2026-03-26`

Scenario:
`npm run perf:run`

Result:
- `parseMs`: `1715.0`
- `renderMs`: `1794.1`
- `totalMs`: `3541.1`
- `elements`: `471931`
- `textNodes`: `455978`
- `pages`: `520`
- `visiblePages`: `1`
- `longTasks.count`: `4`
- `longTasks.sumMs`: `3283`
- `longTasks.maxMs`: `1480`
- `scrollPassMs`: `40078.7`
- `scrollPeakElements`: `471931`

Notes:
- Baseline confirms the bottleneck is still main-thread parse + render + style/layout on a very large DOM.
- The initial DOM size is already at the same order of magnitude as the trace numbers you provided.

## Virtualized Pages

Date:
`2026-03-26`

Scenario:
`node scripts/run-benchmark.mjs virtualized`

Result:
- `parseMs`: `1680.2`
- `renderMs`: `62.4`
- `totalMs`: `1770.8`
- `elements`: `677`
- `textNodes`: `575`
- `pages`: `520`
- `mountedPages`: `3`
- `visiblePages`: `1`
- `longTasks.count`: `4`
- `longTasks.sumMs`: `1684`
- `longTasks.maxMs`: `1447`
- `scrollPassMs`: `37548.3`
- `scrollPeakElements`: `6832`

Delta vs baseline:
- `renderMs`: `-96.5%`
- `totalMs`: `-50.0%`
- `elements`: `-99.86%`
- `scrollPeakElements`: `-98.55%`

Notes:
- The parse cost is still on the main thread, so the largest remaining long task is still dominated by parsing.
- Page virtualization removed the DOM/layout explosion from initial render without changing the mounted page markup.

## Worker Parser

Date:
`2026-03-26`

Scenario:
`node scripts/run-benchmark.mjs worker`

Result:
- `parseMs`: `2154.4`
- `renderMs`: `384.0`
- `totalMs`: `3991.7`
- `elements`: `471931`
- `textNodes`: `455978`
- `pages`: `520`
- `mountedPages`: `520`
- `visiblePages`: `1`
- `longTasks.count`: `2`
- `longTasks.sumMs`: `2100`
- `longTasks.maxMs`: `1445`
- `scrollPassMs`: `41605.7`
- `scrollPeakElements`: `471931`

Delta vs baseline:
- `longTasks.sumMs`: `-36.0%`
- `longTasks.count`: `-50.0%`
- `renderMs`: `-78.6%`
- `totalMs`: `+12.7%`

Notes:
- Worker parsing reduced main-thread long tasks, but by itself it did not improve total time because the document still rendered into the full 520-page DOM.
- The main benefit here is responsiveness, not end-to-end completion time.

## Virtualized + Worker

Date:
`2026-03-26`

Scenario:
`node scripts/run-benchmark.mjs virtualized-worker`

Result:
- `parseMs`: `1978.8`
- `renderMs`: `48.3`
- `totalMs`: `2030.4`
- `elements`: `677`
- `textNodes`: `575`
- `pages`: `520`
- `mountedPages`: `3`
- `visiblePages`: `1`
- `longTasks.count`: `2`
- `longTasks.sumMs`: `381`
- `longTasks.maxMs`: `320`
- `scrollPassMs`: `37498.1`
- `scrollPeakElements`: `6832`

Delta vs baseline:
- `totalMs`: `-42.7%`
- `renderMs`: `-97.3%`
- `elements`: `-99.86%`
- `scrollPeakElements`: `-98.55%`
- `longTasks.sumMs`: `-88.4%`

Notes:
- This is the first combination that meaningfully reduces both end-to-end time and main-thread blocking.
- At this point the main remaining cost is worker parse + transfer, not layout.

## Optimized All

Date:
`2026-03-26`

Scenario:
`node scripts/run-benchmark.mjs optimized-all`

Result:
- `parseMs`: `2121.4`
- `renderMs`: `48.4`
- `totalMs`: `2172.6`
- `elements`: `119`
- `textNodes`: `575`
- `pages`: `520`
- `mountedPages`: `3`
- `visiblePages`: `1`
- `longTasks.count`: `2`
- `longTasks.sumMs`: `384`
- `longTasks.maxMs`: `321`
- `scrollPassMs`: `37339.8`
- `scrollPeakElements`: `741`

Delta vs virtualized + worker:
- `elements`: `-82.4%`
- `scrollPeakElements`: `-89.2%`
- `longTasks.sumMs`: `+0.8%`
- `renderMs`: `+0.2%`
- `totalMs`: `+7.0%`

Notes:
- The previous lower DOM counts came from an over-aggressive run inlining path that broke font styling and drawing layout. That path has been removed.
- The current safe run compaction still reduces mounted DOM materially, but it is no longer a first-order performance win compared with virtualization + worker parsing.

## Worker Handoff

Date:
`2026-03-27`

Scenario:
`node scripts/run-benchmark.mjs worker-parse-only`

Before:
- `parseMs`: `2246.4`
- `totalMs`: `2256.9`
- `parts`: `505`
- `rels`: `4`
- `longTasks.count`: `1`
- `longTasks.sumMs`: `265`
- `longTasks.maxMs`: `265`

After:
- `parseMs`: `1997.2`
- `totalMs`: `2005.3`
- `parts`: `505`
- `rels`: `4`
- `longTasks.count`: `1`
- `longTasks.sumMs`: `216`
- `longTasks.maxMs`: `216`

Delta:
- `parseMs`: `-11.1%`
- `totalMs`: `-11.1%`
- `longTasks.sumMs`: `-18.5%`
- `longTasks.maxMs`: `-18.5%`

Notes:
- This benchmark isolates the worker parse handoff path by running `parseAsync()` only with `useWorkerParser: true`, on the 520-page Routledge DOCX.
- The main improvement comes from keeping the unzip/file table inside a long-lived parser worker session instead of posting the full package back to the main thread.
- The worker-backed package now fetches images, fonts, and altChunks on demand, so the initial `postMessage` payload contains only serialized document structure.

## Snapshot Worker

Date:
`2026-03-30`

Scenarios:
- `node scripts/run-benchmark.mjs virtualized-worker`
- `node scripts/run-benchmark.mjs snapshot-worker`

Internal worker result:
- `parseMs`: `1719.1`
- `renderMs`: `45.3`
- `totalMs`: `1765.9`
- `elements`: `686`
- `pages`: `519`
- `mountedPages`: `3`
- `longTasks.sumMs`: `320`
- `longTasks.maxMs`: `259`

Snapshot worker result:
- `parseMs`: `1360.4`
- `parseRoundtripMs`: `1370.2`
- `transferMs`: `440.1`
- `workerRoundtripMs`: `1810.3`
- `renderMs`: `506.9`
- `totalMs`: `2322.4`
- `elements`: `686`
- `pages`: `519`
- `mountedPages`: `3`
- `snapshotResourceFiles`: `129`
- `longTasks.sumMs`: `793`
- `longTasks.maxMs`: `732`

Delta vs internal worker:
- `parseMs`: `-20.9%`
- `workerRoundtripMs` vs internal `parseMs`: `+5.3%`
- `renderMs`: `+1019.2%`
- `totalMs`: `+31.5%`
- `longTasks.sumMs`: `+147.8%`

Notes:
- The two-step snapshot benchmark measures `parseMs` entirely inside the caller-managed worker, then measures `transferMs` as a separate `postMessage(snapshot, transferables)` phase.
- `parseToSnapshot()` itself is faster than the existing internal worker parse path because unzip, XML parse, and pagination are completed once inside the worker.
- End-to-end time is currently worse because the main thread still pays a large snapshot handoff and rehydrate cost before DOM rendering starts.
- `renderMs` for the snapshot path includes `WordDocument.fromSnapshot(...)` rehydration plus the normal `renderSnapshot(...)` DOM work.
- This benchmark confirms the new API is functionally correct, but it also shows the next optimization target clearly: reduce snapshot payload size for `pages/parts` and cut main-thread rehydration work.

## Snapshot Rehydrate Optimization

Date:
`2026-03-30`

Scenario:
- `node scripts/run-benchmark.mjs snapshot-worker`
- custom breakdown on the same 520-page DOCX with `virtualizePages: true`

Before:
- `snapshot-worker.renderMs`: `517.9`
- `snapshot-worker.totalMs`: `2647.0`
- `snapshot-worker.longTasks.sumMs`: `825`
- `snapshot-worker.longTasks.maxMs`: `744`
- `rehydrateMs`: `502.7`
- `domRenderMs`: `119.9`

After:
- `snapshot-worker.renderMs`: `26.7`
- `snapshot-worker.totalMs`: `2146.9`
- `snapshot-worker.longTasks.sumMs`: `353`
- `snapshot-worker.longTasks.maxMs`: `271`
- `rehydrateMs`: `1.6`
- `domRenderMs`: `3.9`

Delta:
- `snapshot-worker.renderMs`: `-94.8%`
- `snapshot-worker.totalMs`: `-18.9%`
- `snapshot-worker.longTasks.sumMs`: `-57.2%`
- `snapshot-worker.longTasks.maxMs`: `-63.6%`
- `rehydrateMs`: `-99.7%`
- `domRenderMs`: `-96.7%`

Notes:
- `WordDocument.fromSnapshot()` no longer deep-clones the entire `pages` tree up front.
- Snapshot-backed documents now keep authoritative page metadata and clone each page only when it is actually rendered.
- Snapshot part materialization is now selective: `styles`, `header/footer`, and `footnotes/endnotes` stay isolated for correctness, while read-mostly parts reuse the structured-cloned payload directly.
- This keeps snapshot reusability intact while removing the biggest main-thread rehydrate hotspot for virtualized rendering.

## Virtual Scroll Optimization

Date:
`2026-03-30`

Document:
`/home/yuche/Documents/large-docx/big-docx-moby-dick.docx`

Scenario:
- `TEST_DOCX_PATH=/home/yuche/Documents/large-docx/big-docx-moby-dick.docx node scripts/run-benchmark.mjs virtualized-worker`

Before:
- `parseMs`: `36.9`
- `renderMs`: `102.7`
- `totalMs`: `227.8`
- `elements`: `8398`
- `mountedPages`: `3`
- `longTasks.count`: `8`
- `longTasks.sumMs`: `726`
- `longTasks.maxMs`: `106`
- `scrollPassMs`: `24223.3`

After:
- `parseMs`: `31.7`
- `renderMs`: `12.6`
- `totalMs`: `151.5`
- `elements`: `8397`
- `mountedPages`: `3`
- `longTasks.count`: `1`
- `longTasks.sumMs`: `107`
- `longTasks.maxMs`: `107`
- `scrollPassMs`: `23229.6`

Delta:
- `parseMs`: `-14.1%`
- `renderMs`: `-87.7%`
- `totalMs`: `-33.5%`
- `longTasks.count`: `-87.5%`
- `longTasks.sumMs`: `-85.3%`
- `scrollPassMs`: `-4.1%`

Notes:
- `VirtualizedRenderer` no longer does `replaceChildren()` on every window sync.
- Mounted pages are kept in a stable absolute-positioned inner container and updated incrementally.
- Page measurement is frozen during scroll and resumed on idle, using direct `resizeItem()` updates instead of eager `measureElement()` calls.
- A new mounted page window callback is available so hosts can stop watching the subtree with `MutationObserver`.
