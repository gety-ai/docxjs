(function () {
  const statusEl = document.getElementById("status");
  const viewerEl = document.getElementById("viewer");
  const rootEl = document.getElementById("root");
  const stylesEl = document.getElementById("styles");

  function setStatus(text) {
    statusEl.textContent = text;
  }

  function nextFrame() {
    return new Promise(resolve => requestAnimationFrame(() => resolve()));
  }

  function countTextNodes(root) {
    let count = 0;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);

    while (walker.nextNode()) {
      count += 1;
    }

    return count;
  }

  function countVisiblePages(className) {
    const viewerRect = viewerEl.getBoundingClientRect();
    let visible = 0;

    for (const page of rootEl.querySelectorAll(`section.${className}`)) {
      const rect = page.getBoundingClientRect();
      const intersects = rect.bottom > viewerRect.top && rect.top < viewerRect.bottom;

      if (intersects) {
        visible += 1;
      }
    }

    return visible;
  }

  function countTotalPages(className) {
    const virtualHost = rootEl.querySelector("[data-docx-page-count]");
    const totalPages = Number(virtualHost?.dataset?.docxPageCount || 0);

    return totalPages || rootEl.querySelectorAll(`section.${className}`).length;
  }

  async function runScrollPass(className) {
    const maxScrollTop = Math.max(0, viewerEl.scrollHeight - viewerEl.clientHeight);
    const step = Math.max(1, Math.floor(viewerEl.clientHeight * 0.8));
    let peakElements = rootEl.querySelectorAll("*").length;

    for (let scrollTop = 0; scrollTop <= maxScrollTop; scrollTop += step) {
      viewerEl.scrollTop = scrollTop;
      await nextFrame();
      await nextFrame();
      peakElements = Math.max(peakElements, rootEl.querySelectorAll("*").length);
    }

    viewerEl.scrollTop = maxScrollTop;
    await nextFrame();
    await nextFrame();
    peakElements = Math.max(peakElements, rootEl.querySelectorAll("*").length);

    viewerEl.scrollTop = 0;
    await nextFrame();
    await nextFrame();

    return {
      peakElements,
      visiblePagesAfterScroll: countVisiblePages(className)
    };
  }

  async function runBench(options) {
    const docxOptions = options.docxOptions || {};
    const className = docxOptions.className || "docx";
    const longTasks = [];
    const observer = "PerformanceObserver" in window
      ? new PerformanceObserver(list => {
          for (const entry of list.getEntries()) {
            longTasks.push(entry.duration);
          }
        })
      : null;

    rootEl.innerHTML = "";
    stylesEl.innerHTML = "";
    viewerEl.scrollTop = 0;

    if (observer) {
      observer.observe({ type: "longtask", buffered: true });
    }

    setStatus("Fetching docx...");

    const response = await fetch("/api/test-docx");
    const blob = await response.blob();

    setStatus("Parsing...");
    const totalStart = performance.now();
    const parseStart = performance.now();
    const doc = await window.docx.parseAsync(blob, docxOptions);
    const parseMs = performance.now() - parseStart;

    setStatus("Rendering...");
    const renderStart = performance.now();
    await window.docx.renderDocument(doc, rootEl, stylesEl, docxOptions);
    const renderMs = performance.now() - renderStart;

    await nextFrame();
    await nextFrame();

    const totalMs = performance.now() - totalStart;
    const elements = rootEl.querySelectorAll("*").length;
    const mountedPages = rootEl.querySelectorAll(`section.${className}`).length;
    const pages = countTotalPages(className);
    const textNodes = countTextNodes(rootEl);
    const visiblePages = countVisiblePages(className);

    setStatus("Auto-scrolling...");
    const scrollStart = performance.now();
    const scrollPass = await runScrollPass(className);
    const scrollPassMs = performance.now() - scrollStart;

    observer?.disconnect();

    const memory = typeof performance.measureUserAgentSpecificMemory === "function"
      ? await performance.measureUserAgentSpecificMemory().catch(() => null)
      : null;

    const result = {
      label: options.label || "baseline",
      options: docxOptions,
      parseMs,
      renderMs,
      totalMs,
      elements,
      textNodes,
      pages,
      mountedPages,
      visiblePages,
      longTasks: {
        count: longTasks.length,
        sumMs: longTasks.reduce((sum, value) => sum + value, 0),
        maxMs: longTasks.length ? Math.max.apply(null, longTasks) : 0
      },
      scrollPassMs,
      scrollPeakElements: scrollPass.peakElements,
      visiblePagesAfterScroll: scrollPass.visiblePagesAfterScroll,
      memoryBytes: memory?.bytes ?? null
    };

    setStatus(JSON.stringify(result, null, 2));
    window.__DOCX_BENCH_RESULT__ = result;
    return result;
  }

  window.__runDocxBench = async function (options) {
    try {
      return await runBench(options || {});
    } catch (error) {
      window.__DOCX_BENCH_ERROR__ = {
        message: error?.message || String(error),
        stack: error?.stack || ""
      };
      setStatus(JSON.stringify(window.__DOCX_BENCH_ERROR__, null, 2));
      throw error;
    }
  };
})();
