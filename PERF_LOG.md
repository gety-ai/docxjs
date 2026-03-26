# Performance Log

Test document:
`/home/yuche/Downloads/The Routledge Handbook of Translation and Philosophy.docx`

Primary browser:
`/home/yuche/.nix-profile/bin/google-chrome-stable`

Metrics captured per run:
- `parseMs`: `docx.parseAsync(...)`
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
