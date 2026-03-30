import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

import puppeteer from "puppeteer-core";

import { startPerfServer } from "./perf-server.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");
const chromePath = process.env.CHROME_BIN || "/home/yuche/.nix-profile/bin/google-chrome-stable";
const scenarioName = process.argv[2] || "baseline";
const scenarios = {
  baseline: {
    label: "baseline",
    docxOptions: {
      debug: false
    }
  },
  worker: {
    label: "worker-parser",
    docxOptions: {
      debug: false,
      useWorkerParser: true
    }
  },
  "worker-parse-only": {
    label: "worker-parse-only",
    mode: "parse-only",
    docxOptions: {
      debug: false,
      useWorkerParser: true
    }
  },
  virtualized: {
    label: "virtualized-pages",
    docxOptions: {
      debug: false,
      virtualizePages: true
    }
  },
  "virtualized-worker": {
    label: "virtualized-worker",
    docxOptions: {
      debug: false,
      virtualizePages: true,
      useWorkerParser: true
    }
  },
  "snapshot-worker": {
    label: "snapshot-worker",
    mode: "snapshot-worker",
    parseOptions: {
      debug: false
    },
    renderOptions: {
      debug: false,
      virtualizePages: true
    },
    docxOptions: {
      className: "docx"
    }
  },
  "optimized-all": {
    label: "optimized-all",
    docxOptions: {
      debug: false,
      virtualizePages: true,
      useWorkerParser: true,
      mergeAdjacent: true
    }
  }
};

function runCommand(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: projectRoot,
      stdio: "inherit"
    });

    child.once("exit", code => {
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`${command} ${args.join(" ")} exited with code ${code}`));
      }
    });

    child.once("error", reject);
  });
}

async function main() {
  const scenario = scenarios[scenarioName];

  if (!scenario) {
    throw new Error(`Unknown scenario: ${scenarioName}`);
  }

  await runCommand("npm", ["run", "build"]);

  const server = await startPerfServer();
  const browser = await puppeteer.launch({
    executablePath: chromePath,
    headless: "new",
    args: [
      "--disable-dev-shm-usage",
      "--disable-gpu",
      "--no-first-run",
      "--no-default-browser-check"
    ]
  });

  try {
    const page = await browser.newPage();
    await page.goto(`${server.url}/perf/index.html`, { waitUntil: "networkidle0" });
    await page.waitForFunction(() => typeof window.__runDocxBench === "function");

    const result = await page.evaluate(async options => {
      return await window.__runDocxBench(options);
    }, scenario);

    const outputDir = path.join(projectRoot, "perf-results");
    const outputPath = path.join(outputDir, `${scenario.label}.json`);

    await mkdir(outputDir, { recursive: true });
    await writeFile(outputPath, `${JSON.stringify(result, null, 2)}\n`, "utf8");

    console.log(JSON.stringify({
      outputPath,
      result
    }, null, 2));
  } finally {
    await browser.close();
    await server.close();
  }
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
