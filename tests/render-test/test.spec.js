describe("Render document", function () {
  const tests = [
    'text',
    'underlines',
    'text-break',
    'table',
    'page-layout',
    'revision',
    'numbering',
    'line-spacing',
    'header-footer',
    'footnote',
    'equation'
  ];

  for (let path of tests) {
    it(`from ${path} should be correct`, async () => {

      const docBlob = await fetch(`/base/tests/render-test/${path}/document.docx`).then(r => r.blob());
      const resultText = await fetch(`/base/tests/render-test/${path}/result.html`).then(r => r.text());

      const div = document.createElement("div");

      document.body.appendChild(div);

      await docx.renderAsync(docBlob, div);
      
      const actual = formatHTML(div.innerHTML);
      const expected = formatHTML(resultText);

      expect(actual).toBe(expected);

      if(actual != expected) {
        const diffs = Diff.diffLines(expected, actual);

        for(const diff of diffs) {
          if(diff.added)
            console.log('[+] ' + diff.value);

          if(diff.removed)
            console.log('[-] ' + diff.value);
        }
      }

      div.remove();
    });
  }

  it("virtualized multi-page rendering should preserve mounted page markup", async () => {
    const path = 'header-footer';
    const docBlob = await fetch(`/base/tests/render-test/${path}/document.docx`).then(r => r.blob());
    const resultText = await fetch(`/base/tests/render-test/${path}/result.html`).then(r => r.text());
    const expectedRoot = new DOMParser().parseFromString(resultText, "text/html");
    const expectedSections = Array.from(expectedRoot.querySelectorAll("section.docx")).map(x => normalizeSectionHTML(x));
    const scrollHost = document.createElement("div");
    const div = document.createElement("div");

    scrollHost.style.height = "900px";
    scrollHost.style.overflow = "auto";
    scrollHost.appendChild(div);
    document.body.appendChild(scrollHost);

    await docx.renderAsync(docBlob, div, null, {
      virtualizePages: true,
      virtualizePagesOverscan: 0
    });

    const virtualHost = div.querySelector("[data-docx-page-count]");
    const initialSections = Array.from(div.querySelectorAll("section.docx")).map(x => normalizeSectionHTML(x));

    expect(virtualHost.dataset.docxPageCount).toBe(`${expectedSections.length}`);
    expect(initialSections[0]).toBe(expectedSections[0]);

    scrollHost.scrollTop = scrollHost.scrollHeight;
    await waitFrames(6);

    const scrolledSections = Array.from(div.querySelectorAll("section.docx")).map(x => normalizeSectionHTML(x));
    expect(scrolledSections[scrolledSections.length - 1]).toBe(expectedSections[expectedSections.length - 1]);

    scrollHost.remove();
  });

  it("worker parser should preserve multi-page rendering output", async () => {
    const path = 'header-footer';
    const docBlob = await fetch(`/base/tests/render-test/${path}/document.docx`).then(r => r.blob());
    const resultText = await fetch(`/base/tests/render-test/${path}/result.html`).then(r => r.text());
    const div = document.createElement("div");

    document.body.appendChild(div);

    await docx.renderAsync(docBlob, div, null, {
      useWorkerParser: true,
      workerUrl: "/base/dist/docx-preview-worker.js"
    });

    const actual = formatHTML(div.innerHTML);
    const expected = formatHTML(resultText);

    expect(actual).toBe(expected);

    div.remove();
  });

  it("mergeAdjacent should preserve text content while reducing wrapper spans", async () => {
    const path = 'header-footer';
    const docBlob = await fetch(`/base/tests/render-test/${path}/document.docx`).then(r => r.blob());
    const baselineDiv = document.createElement("div");
    const optimizedDiv = document.createElement("div");

    document.body.appendChild(baselineDiv);
    document.body.appendChild(optimizedDiv);

    await docx.renderAsync(docBlob, baselineDiv);
    await docx.renderAsync(docBlob, optimizedDiv, null, {
      mergeAdjacent: true
    });

    expect(normalizeText(optimizedDiv.textContent)).toBe(normalizeText(baselineDiv.textContent));
    expect(optimizedDiv.querySelectorAll("section.docx").length).toBe(baselineDiv.querySelectorAll("section.docx").length);
    expect(optimizedDiv.querySelectorAll("span").length).toBeLessThanOrEqual(baselineDiv.querySelectorAll("span").length);

    baselineDiv.remove();
    optimizedDiv.remove();
  });
});

function formatHTML(text) {
  return text.replace(/\t+|\s+/ig, ' ').replace(/></ig, '>\n<');
}

function normalizeSectionHTML(section) {
  const clone = section.cloneNode(true);
  clone.removeAttribute("data-index");
  return formatHTML(clone.outerHTML);
}

function waitFrames(count) {
  return new Promise(resolve => {
    const step = () => {
      if (count <= 0) {
        resolve();
        return;
      }

      count -= 1;
      requestAnimationFrame(step);
    };

    step();
  });
}

function normalizeText(text) {
  return (text || "").replace(/\s+/g, " ").trim();
}
