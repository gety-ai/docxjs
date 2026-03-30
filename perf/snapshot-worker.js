self.importScripts("/dist/docx-preview.js");

const snapshots = new Map();

self.onmessage = async event => {
  const message = event.data;

  if (!message || !message.type) {
    return;
  }

  try {
    if (message.type === "snapshot-parse") {
      const parseStart = performance.now();
      const snapshot = await self.docx.parseToSnapshot(message.buffer, message.parseOptions || {});
      const parseMs = performance.now() - parseStart;

      snapshots.set(message.requestId, snapshot);
      self.postMessage({
        type: "snapshot-ready",
        requestId: message.requestId,
        parseMs
      });
      return;
    }

    if (message.type === "snapshot-transfer") {
      const snapshot = snapshots.get(message.requestId);

      if (!snapshot) {
        throw new Error("Snapshot not found for transfer");
      }

      snapshots.delete(message.requestId);
      const transferables = self.docx.collectSnapshotTransferables(snapshot);

      self.postMessage({
        type: "snapshot-parsed",
        requestId: message.requestId,
        snapshot
      }, transferables);
    }
  } catch (error) {
    snapshots.delete(message.requestId);
    self.postMessage({
      type: "snapshot-error",
      requestId: message.requestId,
      error: error?.stack || error?.message || String(error)
    });
  }
};
