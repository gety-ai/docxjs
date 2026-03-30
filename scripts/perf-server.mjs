import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const DEFAULT_TEST_DOCX_PATH = "/home/yuche/Downloads/The Routledge Handbook of Translation and Philosophy.docx";
const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8"
};

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");

function resolveContentType(filePath) {
  return MIME_TYPES[path.extname(filePath)] || "application/octet-stream";
}

function normalizeRequestPath(urlPathname) {
  const decoded = decodeURIComponent(urlPathname);
  const relativePath = decoded === "/" ? "/perf/index.html" : decoded;
  const absolutePath = path.resolve(projectRoot, `.${relativePath}`);

  if (!absolutePath.startsWith(projectRoot)) {
    return null;
  }

  return absolutePath;
}

export async function startPerfServer(port = 4173, options = {}) {
  const testDocxPath = options.testDocxPath || process.env.TEST_DOCX_PATH || DEFAULT_TEST_DOCX_PATH;
  const server = createServer(async (request, response) => {
    try {
      const requestUrl = new URL(request.url, `http://${request.headers.host}`);

      if (requestUrl.pathname === "/api/test-docx") {
        const file = await readFile(testDocxPath);
        response.writeHead(200, {
          "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "Content-Length": file.byteLength
        });
        response.end(file);
        return;
      }

      if (requestUrl.pathname === "/api/health") {
        response.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        response.end(JSON.stringify({ ok: true }));
        return;
      }

      const filePath = normalizeRequestPath(requestUrl.pathname);

      if (!filePath) {
        response.writeHead(403);
        response.end("Forbidden");
        return;
      }

      const stats = statSync(filePath, { throwIfNoEntry: false });

      if (!stats || !stats.isFile()) {
        response.writeHead(404);
        response.end("Not found");
        return;
      }

      const file = await readFile(filePath);
      response.writeHead(200, {
        "Content-Type": resolveContentType(filePath),
        "Content-Length": file.byteLength
      });
      response.end(file);
    } catch (error) {
      response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
      response.end(error?.stack || String(error));
    }
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", () => resolve());
  });

  return {
    close: () => new Promise(resolve => server.close(() => resolve())),
    url: `http://127.0.0.1:${port}`
  };
}
