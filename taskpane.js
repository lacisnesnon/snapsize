/* global Office, SnapPng */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const MIN = 0.05, MAX = 4, STEP = 1.1;

  let base = null;      // canvas holding the full-resolution (rotated) original
  let scale = 1;
  let origBytes = 0;
  let busy = false;
  let fitWidth = 600;

  Office.onReady(() => {
    try {
      const s = Office.context.roamingSettings && Office.context.roamingSettings.get("fitWidth");
      if (s) fitWidth = s;
    } catch (e) { /* running outside Outlook */ }
    $("fitWidth").value = fitWidth;
    wire();
    $("drop").focus();
  });

  // ---------- wiring ----------

  function wire() {
    document.addEventListener("paste", (e) => {
      const item = [...(e.clipboardData?.items || [])].find((i) => i.type.startsWith("image/"));
      if (item) { e.preventDefault(); load(item.getAsFile()); }
      else if (!base) setStatus("The clipboard doesn't contain an image.", "err");
    });

    const drop = $("drop");
    drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("over"));
    drop.addEventListener("drop", (e) => {
      e.preventDefault(); drop.classList.remove("over");
      const f = [...e.dataTransfer.files].find((x) => x.type.startsWith("image/"));
      if (f) load(f);
    });
    $("file").addEventListener("change", (e) => { if (e.target.files[0]) load(e.target.files[0]); e.target.value = ""; });

    $("fit").onclick = fit;
    $("zoomIn").onclick = () => zoom(STEP);
    $("zoomOut").onclick = () => zoom(1 / STEP);
    $("actual").onclick = () => { scale = 1; render(); };
    $("rotate").onclick = rotate;
    $("border").onchange = render;
    $("done").onclick = done;
    $("discard").onclick = reset;

    $("fitWidth").onchange = (e) => {
      const v = Math.max(100, Math.min(2000, parseInt(e.target.value, 10) || 600));
      fitWidth = v; e.target.value = v;
      try { Office.context.roamingSettings.set("fitWidth", v); Office.context.roamingSettings.saveAsync(); } catch (err) { /* ignore */ }
    };

    document.addEventListener("keydown", (e) => {
      if (!base || busy || e.target.tagName === "INPUT") return;
      const k = e.key;
      if (k === "+" || k === "=") zoom(STEP);
      else if (k === "-" || k === "_") zoom(1 / STEP);
      else if (k === "0") { scale = 1; render(); }
      else if (k === "f" || k === "F") fit();
      else if (k === "r" || k === "R") rotate();
      else if (k === "Enter") done();
      else if (k === "Escape" || k === "Delete") reset();
      else return;
      e.preventDefault();
    });
  }

  // ---------- edit mode ----------

  async function load(blob) {
    try {
      const bmp = await createImageBitmap(blob);
      base = canvas(bmp.width, bmp.height);
      base.getContext("2d").drawImage(bmp, 0, 0);
      bmp.close && bmp.close();
      origBytes = blob.size;
      scale = Math.min(1, fitWidth / base.width);
      $("border").checked = false;
      $("empty").hidden = true;
      $("editor").hidden = false;
      render();
      $("done").focus();
    } catch (e) {
      setStatus("Couldn't read that image.", "err");
    }
  }

  function target() {
    return { w: Math.max(1, Math.round(base.width * scale)), h: Math.max(1, Math.round(base.height * scale)) };
  }

  function zoom(f) {
    let s = Math.max(MIN, Math.min(MAX, scale * f));
    if (Math.abs(s - 1) < 0.03) s = 1;
    scale = s; render();
  }

  function fit() { scale = Math.max(MIN, Math.min(MAX, fitWidth / base.width)); render(); }

  function rotate() {
    const r = canvas(base.height, base.width);
    const g = r.getContext("2d");
    g.translate(r.width, 0);
    g.rotate(Math.PI / 2);
    g.drawImage(base, 0, 0);
    base = r;
    render();
  }

  /** Preview at true email size (CSS px), drawn at device resolution so it's sharp. */
  function render() {
    const { w, h } = target();
    const dpr = window.devicePixelRatio || 1;
    const c = $("preview");
    c.width = Math.round(w * dpr);
    c.height = Math.round(h * dpr);
    c.style.width = w + "px";
    c.style.height = h + "px";
    const g = c.getContext("2d");
    g.imageSmoothingEnabled = true;
    g.imageSmoothingQuality = "high";
    g.drawImage(base, 0, 0, c.width, c.height);
    if ($("border").checked) {
      g.strokeStyle = "#aaaaaa";
      g.lineWidth = dpr;
      g.strokeRect(dpr / 2, dpr / 2, c.width - dpr, c.height - dpr);
    }
    $("info").textContent = `${w} × ${h} px · ${Math.round(scale * 100)}%`;
  }

  // ---------- Done: resample, compress, insert ----------

  async function done() {
    if (!base || busy) return;
    busy = true;
    setButtons(false);
    try {
      const { w, h } = target();
      const out = resample(base, w, h);
      if ($("border").checked) {
        const g = out.getContext("2d");
        g.strokeStyle = "#aaaaaa";
        g.lineWidth = 1;
        g.strokeRect(0.5, 0.5, w - 1, h - 1);
      }
      const data = out.getContext("2d").getImageData(0, 0, w, h);
      let bytes = await SnapPng.encode(data);

      // Keep whichever is smaller: our encoder or the browser's.
      const native = await new Promise((r) => out.toBlob(r, "image/png"));
      if (native && native.size < bytes.length) bytes = new Uint8Array(await native.arrayBuffer());

      await insert(toBase64(bytes), w, h);
      reset();
      setStatus(`Inserted ${w} × ${h} px · ${kb(bytes.length)}` + (origBytes ? ` (pasted image was ${kb(origBytes)})` : ""), "ok");
    } catch (e) {
      setStatus(e.message || String(e), "err");
    } finally {
      busy = false;
      setButtons(true);
    }
  }

  /** High-quality downscale: halve repeatedly, then one final step to the exact size. */
  function resample(src, w, h) {
    let cur = src;
    if (w === src.width && h === src.height) {
      const c = canvas(w, h); c.getContext("2d").drawImage(src, 0, 0); return c;
    }
    while (cur.width / 2 >= w && cur.height / 2 >= h) {
      const c = canvas(Math.floor(cur.width / 2), Math.floor(cur.height / 2));
      const g = c.getContext("2d");
      g.imageSmoothingQuality = "high";
      g.drawImage(cur, 0, 0, c.width, c.height);
      cur = c;
    }
    const c = canvas(w, h);
    const g = c.getContext("2d");
    g.imageSmoothingEnabled = true;
    g.imageSmoothingQuality = "high";
    g.drawImage(cur, 0, 0, w, h);
    return c;
  }

  /** Adds the PNG as an inline attachment and inserts an <img> at the cursor with fixed width/height. */
  async function insert(b64, w, h) {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.body) throw new Error("Open an email you're writing, then try again.");

    const type = await call((cb) => item.body.getTypeAsync(cb));
    if (type !== Office.CoercionType.Html && type !== "html") {
      throw new Error("This email is in plain text. Switch it to HTML format to insert images.");
    }

    const name = `image-${Date.now()}.png`;
    await call((cb) => item.addFileAttachmentFromBase64Async(b64, name, { isInline: true }, cb));

    const html = `<img src="cid:${name}" width="${w}" height="${h}" style="width:${w}px;height:${h}px;border:0" alt="">`;
    await call((cb) => item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, cb));
  }

  function call(fn) {
    return new Promise((resolve, reject) =>
      fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message)))));
  }

  // ---------- helpers ----------

  function reset() {
    base = null;
    $("editor").hidden = true;
    $("empty").hidden = false;
    setStatus("");
    $("drop").focus();
  }

  function setButtons(on) { document.querySelectorAll("#editor button").forEach((b) => (b.disabled = !on)); }

  function setStatus(text, kind) {
    const s = $("status");
    s.textContent = text;
    s.className = "status" + (kind ? " " + kind : "");
  }

  function canvas(w, h) { const c = document.createElement("canvas"); c.width = w; c.height = h; return c; }

  function toBase64(bytes) {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  function kb(n) { return n < 1024 * 1024 ? Math.max(1, Math.round(n / 1024)) + " KB" : (n / 1048576).toFixed(1) + " MB"; }
})();
