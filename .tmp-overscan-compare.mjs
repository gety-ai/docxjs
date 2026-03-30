import { rm } from 'node:fs/promises';
import puppeteer from 'puppeteer-core';
import { startPerfServer } from './scripts/perf-server.mjs';

const chromePath = process.env.CHROME_BIN || '/home/yuche/.nix-profile/bin/google-chrome-stable';
const scenarios = [
  {
    name: 'virtualized-worker-overscan-2',
    options: {
      label: 'virtualized-worker-overscan-2',
      docxOptions: {
        debug: false,
        virtualizePages: true,
        virtualizePagesOverscan: 2,
        useWorkerParser: true,
      },
    },
  },
  {
    name: 'virtualized-worker-overscan-3',
    options: {
      label: 'virtualized-worker-overscan-3',
      docxOptions: {
        debug: false,
        virtualizePages: true,
        virtualizePagesOverscan: 3,
        useWorkerParser: true,
      },
    },
  },
  {
    name: 'snapshot-worker-overscan-2',
    options: {
      label: 'snapshot-worker-overscan-2',
      mode: 'snapshot-worker',
      parseOptions: { debug: false },
      renderOptions: {
        debug: false,
        virtualizePages: true,
        virtualizePagesOverscan: 2,
      },
      docxOptions: { className: 'docx' },
    },
  },
  {
    name: 'snapshot-worker-overscan-3',
    options: {
      label: 'snapshot-worker-overscan-3',
      mode: 'snapshot-worker',
      parseOptions: { debug: false },
      renderOptions: {
        debug: false,
        virtualizePages: true,
        virtualizePagesOverscan: 3,
      },
      docxOptions: { className: 'docx' },
    },
  },
];

const server = await startPerfServer();
const browser = await puppeteer.launch({
  executablePath: chromePath,
  headless: 'new',
  args: ['--disable-dev-shm-usage', '--disable-gpu', '--no-first-run', '--no-default-browser-check'],
});

try {
  const page = await browser.newPage();
  await page.goto(server.url + '/perf/index.html', { waitUntil: 'networkidle0' });
  await page.waitForFunction(() => typeof window.__runDocxBench === 'function');

  const results = [];
  for (const scenario of scenarios) {
    const result = await page.evaluate(async options => await window.__runDocxBench(options), scenario.options);
    results.push({ name: scenario.name, result });
  }

  console.log(JSON.stringify(results, null, 2));
} finally {
  await browser.close();
  await server.close();
  await rm('.tmp-overscan-compare.mjs', { force: true });
}
