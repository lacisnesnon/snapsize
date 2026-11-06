/* global Office, SnapEditor, SnapPng, SnapSettings */
/* SnapSize self-test: checks Outlook, the hosting, every drawing tool, resizing and PNG compression. */
(function (global) {
  "use strict";

  const E = () => global.SnapEditor;

  function inOutlook() {
    try { return typeof Office !== "undefined" && !!(Office.context && Office.context.mailbox); } catch (e) { return false; }
  }

  function call(fn) {
    return new Promise((resolve, reject) =>
      fn((r) => (r.status === "succeeded" ? resolve(r.value) : reject(new Error(r.error && r.error.message)))));
  }

  function whiteCanvas(w, h) {
    const c = E().canvas(w, h), g = c.getContext("2d");
    g.fillStyle = "#ffffff"; g.fillRect(0, 0, w, h);
    return c;
  }

  function changedPixels(c) {
    const d = c.getContext("2d").getImageData(0, 0, c.width, c.height).data;
    let n = 0;
    for (let i = 0; i < d.length; i += 4) if (d[i] !== 255 || d[i + 1] !== 255 || d[i + 2] !== 255) n++;
    return n;
  }

  /** A screenshot-like sample: flat colours, text and a few annotations. */
  function sampleImage() {
    const c = whiteCanvas(480, 270), g = c.getContext("2d");
    g.fillStyle = "#f3f3f3"; g.fillRect(0, 0, 480, 40);
    g.fillStyle = "#0f6cbd"; g.fillRect(16, 60, 200, 120);
    g.fillStyle = "#1f1f1f"; g.font = '600 20px "Segoe UI", sans-serif'; g.fillText("SnapSize test image", 16, 27);
    g.fillStyle = "#ffffff"; g.font = '16px "Segoe UI", sans-serif'; g.fillText("Inserted at true size", 28, 128);
    const { drawShape } = E();
    drawShape(g, { type: "arrow", x1: 330, y1: 200, x2: 230, y2: 140, color: "#e81123", width: 3 }, 1);
    drawShape(g, { type: "step", x1: 350, y1: 210, n: 1, color: "#e81123", fontSize: 16 }, 1);
    drawShape(g, { type: "ellipse", x1: 250, y1: 60, x2: 460, y2: 120, color: "#107c10", width: 3 }, 1);
    return c;
  }

  /**
   * Runs every check. `report(result)` is called as each one finishes.
   * result = { name, status: "pass" | "warn" | "fail", detail }
   */
  async function run(opts, report) {
    const results = [];
    const add = (name, status, detail) => { const r = { name, status, detail }; results.push(r); report && report(r); };
    const test = async (name, fn) => {
      try { const r = await fn(); add(name, r && r.status ? r.status : "pass", r && r.detail ? r.detail : (typeof r === "string" ? r : "")); }
      catch (e) { add(name, "fail", e.message || String(e)); }
    };

    // ---- Outlook & hosting
    await test("Running inside Outlook", () => {
      if (typeof Office === "undefined") throw new Error("Office.js didn't load — check your internet connection.");
      if (!inOutlook()) return { status: "warn", detail: "Opened outside Outlook (fine for testing in a browser)." };
      const d = Office.context.diagnostics || {};
      return `${d.host || "Outlook"} ${d.platform || ""} ${d.version || ""}`.trim();
    });

    await test("Outlook supports the features SnapSize needs", () => {
      if (!inOutlook()) return { status: "warn", detail: "Skipped (not in Outlook)." };
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) throw new Error("This Outlook is too old (needs Mailbox 1.8). Update Outlook.");
      return "Mailbox 1.8+";
    });

    await test("An email you're writing is open, in HTML format", async () => {
      if (!inOutlook()) return { status: "warn", detail: "Skipped (not in Outlook)." };
      const item = Office.context.mailbox.item;
      if (!item || !item.body) throw new Error("Open a new email or reply, then run the test again.");
      const t = await call((cb) => item.body.getTypeAsync(cb));
      if (t !== "html") throw new Error("This email is plain text. Switch it to HTML (Format text → HTML).");
      return "Ready to insert images";
    });

    await test("Hosted files are up to date", async () => {
      const r = await fetch("version.json?t=" + Date.now(), { cache: "no-store" });
      if (!r.ok) throw new Error(`version.json not found on your hosting (HTTP ${r.status}). Upload all files from the new-outlook folder.`);
      const v = (await r.json()).version;
      if (v !== opts.version) return { status: "warn", detail: `Hosting has v${v} but Outlook is showing v${opts.version}. Close and reopen the SnapSize pane.` };
      return "v" + v;
    });

    await test("Settings save and load", async () => {
      const s = SnapSettings.load();
      const roamed = await SnapSettings.save(s);
      const back = SnapSettings.load();
      if (JSON.stringify(back) !== JSON.stringify(s)) throw new Error("Settings didn't read back the same.");
      return roamed ? "Saved to your mailbox (roams to all devices)" : { status: inOutlook() ? "warn" : "pass", detail: "Saved on this device only" };
    });

    // ---- Drawing tools
    const { drawShape, pixelate, resample } = E();
    const shapes = {
      Rectangle: { type: "rect", x1: 10, y1: 10, x2: 90, y2: 70 },
      Ellipse: { type: "ellipse", x1: 10, y1: 10, x2: 90, y2: 70 },
      Line: { type: "line", x1: 10, y1: 10, x2: 90, y2: 70 },
      Arrow: { type: "arrow", x1: 10, y1: 70, x2: 90, y2: 10 },
      Freehand: { type: "free", x1: 10, y1: 10, points: [[10, 10], [40, 50], [90, 20]] },
      Text: { type: "text", x1: 10, y1: 10, text: "Hi", fontSize: 20 },
      "Speech bubble": { type: "bubble", x1: 10, y1: 10, x2: 90, y2: 50, text: "Hi", fontSize: 14 },
      "Step number": { type: "step", x1: 50, y1: 40, n: 3, fontSize: 16 },
    };
    await test("Drawing tools render", () => {
      const bad = [];
      for (const [name, s] of Object.entries(shapes)) {
        const c = whiteCanvas(100, 80);
        drawShape(c.getContext("2d"), { color: "#e81123", width: 3, ...s }, 1);
        if (changedPixels(c) < 20) bad.push(name);
      }
      if (bad.length) throw new Error("Didn't draw: " + bad.join(", "));
      return Object.keys(shapes).length + " tools OK";
    });

    await test("Highlight is see-through", () => {
      const c = whiteCanvas(40, 40), g = c.getContext("2d");
      g.fillStyle = "#000000"; g.fillRect(0, 0, 20, 40);
      drawShape(g, { type: "highlight", x1: 0, y1: 0, x2: 40, y2: 40, color: "#fff200" }, 1);
      const d = g.getImageData(0, 0, 40, 40).data;
      const dark = d[(20 * 40 + 5) * 4], yellow = [d[(20 * 40 + 30) * 4], d[(20 * 40 + 30) * 4 + 2]];
      if (dark > 10) throw new Error("Highlight covered dark content instead of blending with it.");
      if (yellow[0] < 240 || yellow[1] > 20) throw new Error("Highlight colour is wrong.");
      return "Text under highlights stays readable";
    });

    await test("Obfuscate hides detail", () => {
      const c = whiteCanvas(32, 32), g = c.getContext("2d");
      for (let y = 0; y < 32; y++) for (let x = 0; x < 32; x++) if ((x + y) % 2) { g.fillStyle = "#000"; g.fillRect(x, y, 1, 1); }
      pixelate(g, 0, 0, 32, 32, 8);
      const d = g.getImageData(0, 0, 32, 32).data;
      for (let i = 4; i < 8 * 4; i += 4) if (d[i] !== d[0]) throw new Error("Pixels inside a block still differ.");
      return "8 px blocks";
    });

    await test("Resizing gives the exact size", () => {
      const src = whiteCanvas(1920, 1080);
      for (const [w, h] of [[600, 338], [123, 69], [1920, 1080], [2400, 1350]]) {
        const out = resample(src, w, h);
        if (out.width !== w || out.height !== h) throw new Error(`Asked for ${w}×${h}, got ${out.width}×${out.height}`);
      }
      return "Down, up and 100% all exact";
    });

    // ---- Compression
    await test("PNG compression works", async () => {
      if (typeof CompressionStream === "undefined") throw new Error("This Outlook's browser engine lacks CompressionStream — update Outlook/WebView2.");
      const c = sampleImage();
      const img = c.getContext("2d").getImageData(0, 0, c.width, c.height);
      const bytes = await SnapPng.encode(img);
      const bmp = await createImageBitmap(new Blob([bytes], { type: "image/png" }));
      const back = E().canvas(bmp.width, bmp.height);
      back.getContext("2d").drawImage(bmp, 0, 0);
      const d2 = back.getContext("2d").getImageData(0, 0, bmp.width, bmp.height).data;
      if (bmp.width !== c.width || bmp.height !== c.height) throw new Error("Decoded image is the wrong size.");
      let diff = 0;
      for (let i = 0; i < d2.length; i++) if (Math.abs(d2[i] - img.data[i]) > 1) diff++;
      if (diff) throw new Error(`${diff} values changed after compression (should be lossless).`);
      const raw = c.width * c.height * 4;
      return `Lossless · ${c.width}×${c.height} sample = ${Math.round(bytes.length / 1024)} KB (${Math.round((1 - bytes.length / raw) * 100)}% smaller than raw)`;
    });

    await test("Paste from clipboard is available", () => {
      if (!("onpaste" in document)) throw new Error("Paste events aren't supported here.");
      return "Press Ctrl+V in the pane";
    });

    return results;
  }

  global.SnapTest = { run, sampleImage };
})(window);
