/* global Office, SnapPng, SnapEditor, SnapSettings, SnapTest */
(function () {
  "use strict";

  const VERSION = "1.1.0";
  const $ = (id) => document.getElementById(id);
  const STEP = 1.1;

  // Greenshot-style tool icons (20×20, drawn in the current text colour)
  const S = (body, fill) => `<svg viewBox="0 0 20 20" aria-hidden="true" fill="${fill || "none"}" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;
  const ICONS = {
    select: S('<path d="M5 3v13l3.5-3.2L11 18l2-1-2.4-5H15z" fill="currentColor" stroke-width="1"/>'),
    rect: S('<rect x="3" y="5" width="14" height="10" rx="1"/>'),
    ellipse: S('<ellipse cx="10" cy="10" rx="7" ry="5.5"/>'),
    line: S('<path d="M4 16 16 4"/>'),
    arrow: S('<path d="M4 16 15 5"/><path d="M9 4.5h6.5V11" />'),
    free: S('<path d="M3 13c2-5 4-6 5-3s2 4 4 0 3-4 5-2"/>'),
    text: S('<path d="M5 4.5h10M10 4.5V16M7.5 16h5"/>'),
    bubble: S('<path d="M4 4h12a1.5 1.5 0 0 1 1.5 1.5v6A1.5 1.5 0 0 1 16 13H9l-4 3.5V13H4a1.5 1.5 0 0 1-1.5-1.5v-6A1.5 1.5 0 0 1 4 4Z"/>'),
    step: S('<circle cx="10" cy="10" r="7.5" fill="currentColor" stroke="none"/><path d="M8.6 7.2 10.4 6v8" stroke="#fff" stroke-width="1.8"/>'),
    highlight: S('<path d="m11 3.5 5.5 5.5-6 6H5v-5.5z"/><path d="M3 18h14" stroke="#e6c700" stroke-width="2.6"/>'),
    obfuscate: S('<rect x="3" y="3" width="4.5" height="4.5" fill="currentColor"/><rect x="12.5" y="3" width="4.5" height="4.5" fill="currentColor"/><rect x="7.75" y="7.75" width="4.5" height="4.5" fill="currentColor"/><rect x="3" y="12.5" width="4.5" height="4.5" fill="currentColor"/><rect x="12.5" y="12.5" width="4.5" height="4.5" fill="currentColor"/>', "none"),
    effects: S('<path d="M10 2.5 11.6 8l5.4 2-5.4 2L10 17.5 8.4 12 3 10l5.4-2z"/>'),
    crop: S('<path d="M6 2.5V14h11.5M2.5 6H14v11.5"/>'),
    rotateL: S('<path d="M4 8a6.5 6.5 0 1 1 1.5 6.5"/><path d="M3.5 3.5V8H8"/>'),
    rotateR: S('<path d="M16 8a6.5 6.5 0 1 0-1.5 6.5"/><path d="M16.5 3.5V8H12"/>'),
    resize: S('<path d="M3 8V3h5M17 12v5h-5M3 3l5.5 5.5M17 17l-5.5-5.5"/>'),
  };
  const ACTIONS = new Set(["effects", "rotateL", "rotateR", "resize"]);
  const NEEDS = {
    width: ["rect", "ellipse", "line", "arrow", "free", "bubble"],
    font: ["text", "bubble", "step"],
    fill: ["rect", "ellipse"],
  };

  let settings, editor, busy = false, origBytes = 0, lastInfo = null;

  Office.onReady(() => {
    settings = SnapSettings.load();
    editor = new SnapEditor.Editor({
      canvas: $("preview"), frame: $("frame"),
      settings: () => settings, onChange: update,
    });
    buildToolbar();
    wire();
    $("version").textContent = VERSION;
    $("drop").focus();
  });

  // ---------------------------------------------------------------- toolbar

  function buildToolbar() {
    const box = $("tools");
    box.innerHTML = "";
    for (const t of settings.toolbar) {
      if (!t.visible) continue;
      const info = SnapSettings.toolInfo(t.id);
      const b = document.createElement("button");
      b.innerHTML = ICONS[t.id];
      b.dataset.tool = t.id;
      const key = info.hint || (info.key && info.key.toUpperCase());
      b.title = info.label + (key ? ` (${key})` : "");
      b.setAttribute("aria-label", info.label);
      b.onclick = () => useTool(t.id);
      box.appendChild(b);
    }
    markActive();
  }

  function useTool(id) {
    if (!editor.base) return;
    if (id === "rotateL") return editor.rotate(-1);
    if (id === "rotateR") return editor.rotate(1);
    if (id === "resize") { $("widthPx").focus(); $("widthPx").select(); return; }
    if (id === "effects") { $("effectsMenu").hidden = !$("effectsMenu").hidden; markActive(); return; }
    editor.setTool(id);
    syncProps();
  }

  function markActive() {
    document.querySelectorAll("#tools button").forEach((b) => {
      const id = b.dataset.tool;
      const on = id === editor.tool || (id === "effects" && !$("effectsMenu").hidden);
      b.classList.toggle("active", on);
      b.setAttribute("aria-pressed", on ? "true" : "false");
    });
  }

  function syncProps() {
    const p = editor.props, t = editor.tool;
    $("pColor").value = p.color;
    $("pWidth").value = String(p.width);
    if (!$("pWidth").value) $("pWidth").value = "3";
    $("pFont").value = p.fontSize;
    $("pFill").checked = !!p.fill;
    const show = (list) => t === "select" ? !!editor.sel : list.includes(t);
    $("pWidthWrap").hidden = !show(NEEDS.width);
    $("pFontWrap").hidden = !show(NEEDS.font);
    $("pFillWrap").hidden = !show(NEEDS.fill);
    $("pColor").parentElement.hidden = ["crop", "obfuscate"].includes(t) || (t === "select" && !editor.sel);
  }

  /** Called by the editor after every change. */
  function update(info) {
    lastInfo = info;
    if (document.activeElement !== $("widthPx")) $("widthPx").value = info.w;
    $("info").textContent = `× ${info.h} px · ${Math.round(info.scale * 100)}%` + (info.fx.shadow ? " +shadow" : "");
    $("cropBar").hidden = !(info.tool === "crop");
    $("props").hidden = info.tool === "crop";
    $("cropApply").disabled = !info.cropPending;
    $("undo").disabled = !info.canUndo;
    document.querySelectorAll("#effectsMenu input").forEach((cb) => (cb.checked = !!info.fx[cb.dataset.fx]));
    markActive();
    syncProps();
  }

  // ---------------------------------------------------------------- wiring

  function wire() {
    document.addEventListener("paste", (e) => {
      if (!$("settings").hidden) return;
      if (e.target && e.target.tagName === "TEXTAREA") return;
      const item = [...((e.clipboardData && e.clipboardData.items) || [])].find((i) => i.type.startsWith("image/"));
      if (item) { e.preventDefault(); load(item.getAsFile()); }
      else if (!editor.base) setStatus("The clipboard doesn't contain an image.", "err");
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

    $("fit").onclick = () => editor.setScale(settings.fitWidth / editor.base.width);
    $("zoomIn").onclick = () => editor.setScale(editor.scale * STEP);
    $("zoomOut").onclick = () => editor.setScale(editor.scale / STEP);
    $("actual").onclick = () => editor.setScale(1);
    $("widthPx").addEventListener("change", (e) => {
      const v = parseInt(e.target.value, 10);
      if (v > 0) editor.setWidthPx(v); else e.target.value = lastInfo ? lastInfo.w : "";
    });
    $("widthPx").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); e.target.blur(); } });

    $("pColor").oninput = (e) => editor.setProp("color", e.target.value);
    $("pWidth").onchange = (e) => editor.setProp("width", Number(e.target.value));
    $("pFont").onchange = (e) => editor.setProp("fontSize", Math.max(8, Math.min(72, Number(e.target.value) || 16)));
    $("pFill").onchange = (e) => editor.setProp("fill", e.target.checked);
    $("undo").onclick = () => editor.undo();

    document.querySelectorAll("#effectsMenu input").forEach((cb) => (cb.onchange = () => editor.toggleEffect(cb.dataset.fx, cb.checked)));
    $("cropApply").onclick = () => editor.applyCrop();
    $("cropCancel").onclick = () => { editor.cancelCrop(); editor.setTool("select"); };

    $("done").onclick = done;
    $("discard").onclick = reset;

    $("openSettings").onclick = openSettings;
    $("closeSettings").onclick = closeSettings;
    $("runTest").onclick = runTest;
    $("insertTest").onclick = insertTest;

    document.addEventListener("keydown", onKey);
  }

  function onKey(e) {
    if (!$("settings").hidden) { if (e.key === "Escape") closeSettings(); return; }
    if (!editor.base || busy) return;
    const tag = e.target.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;

    const k = e.key;
    if ((e.ctrlKey || e.metaKey) && k.toLowerCase() === "z") { editor.undo(); e.preventDefault(); return; }
    if (e.ctrlKey || e.metaKey || e.altKey) return;

    if (k === "Enter") { if (editor.tool === "crop" && editor.crop) editor.applyCrop(); else done(); }
    else if (k === "Escape") { if (editor.crop) editor.cancelCrop(); useTool("select"); }
    else if (k === "Delete" || k === "Backspace") { if (!editor.deleteSelected()) return; }
    else if (k === "+" || k === "=") editor.setScale(editor.scale * STEP);
    else if (k === "-" || k === "_") editor.setScale(editor.scale / STEP);
    else if (k === "0") editor.setScale(1);
    else {
      const t = SnapSettings.TOOLS.find((x) => x.key && x.key.length === 1 && x.key === k.toLowerCase());
      if (!t) return;
      useTool(t.id);
    }
    e.preventDefault();
  }

  // ---------------------------------------------------------------- load / done

  async function load(blob) {
    try {
      const bmp = await createImageBitmap(blob);
      const base = SnapEditor.canvas(bmp.width, bmp.height);
      base.getContext("2d").drawImage(bmp, 0, 0);
      if (bmp.close) bmp.close();
      origBytes = blob.size;
      $("effectsMenu").hidden = true;
      $("empty").hidden = true;
      $("editor").hidden = false;
      editor.load(base);
      $("done").focus();
    } catch (e) {
      setStatus("Couldn't read that image.", "err");
    }
  }

  async function done() {
    if (!editor.base || busy) return;
    busy = true; setButtons(false);
    try {
      const out = editor.renderOutput();
      const bytes = await insertCanvas(out);
      reset();
      setStatus(`Inserted ${out.width} × ${out.height} px · ${kb(bytes)}` + (origBytes ? ` (pasted image was ${kb(origBytes)})` : ""), "ok");
    } catch (e) {
      setStatus(e.message || String(e), "err");
    } finally {
      busy = false; setButtons(true);
    }
  }

  /** Compresses a canvas and inserts it at the cursor. Returns the PNG size in bytes. */
  async function insertCanvas(c) {
    const w = c.width, h = c.height;
    let bytes = await SnapPng.encode(c.getContext("2d").getImageData(0, 0, w, h));
    const native = await new Promise((r) => c.toBlob(r, "image/png"));
    if (native && native.size < bytes.length) bytes = new Uint8Array(await native.arrayBuffer());
    await insert(toBase64(bytes), w, h);
    return bytes.length;
  }

  async function insert(b64, w, h) {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.body) throw new Error("Open an email you're writing, then try again.");
    const type = await call((cb) => item.body.getTypeAsync(cb));
    if (type !== "html") throw new Error("This email is in plain text. Switch it to HTML format to insert images.");
    const name = `image-${Date.now()}.png`;
    await call((cb) => item.addFileAttachmentFromBase64Async(b64, name, { isInline: true }, cb));
    const html = `<img src="cid:${name}" width="${w}" height="${h}" style="width:${w}px;height:${h}px;border:0" alt="">`;
    await call((cb) => item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, cb));
  }

  function call(fn) {
    return new Promise((resolve, reject) =>
      fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message)))));
  }

  // ---------------------------------------------------------------- settings & self-test

  function openSettings() {
    SnapSettings.mountUI($("settingsBody"), () => settings, (ns) => {
      settings = ns;
      SnapSettings.save(ns);
      buildToolbar();
    }, ICONS);
    $("main").hidden = true;
    $("settings").hidden = false;
    $("openSettings").hidden = true;
    $("closeSettings").hidden = false;
    $("title").textContent = "SnapSize settings";
    $("closeSettings").focus();
  }

  function closeSettings() {
    $("settings").hidden = true;
    $("main").hidden = false;
    $("openSettings").hidden = false;
    $("closeSettings").hidden = true;
    $("title").textContent = "SnapSize";
    buildToolbar();
    if (editor.base) editor.render(); else $("drop").focus();
  }

  async function runTest() {
    const list = $("testResults");
    list.innerHTML = "";
    $("runTest").disabled = true;
    const results = await SnapTest.run({ version: VERSION }, (r) => list.appendChild(resultRow(r)));
    const fails = results.filter((r) => r.status === "fail").length, warns = results.filter((r) => r.status === "warn").length;
    const li = document.createElement("li");
    li.className = fails ? "fail" : warns ? "warn" : "pass";
    li.innerHTML = `<span class="mark">${fails ? "✗" : warns ? "!" : "✓"}</span><span class="summary"></span>`;
    li.querySelector(".summary").textContent = fails ? `${fails} problem${fails > 1 ? "s" : ""} found` : warns ? "Working, with notes above" : "Everything is working";
    list.prepend(li);
    $("runTest").disabled = false;
  }

  function resultRow(r) {
    const li = document.createElement("li");
    li.className = r.status;
    const mark = document.createElement("span"); mark.className = "mark"; mark.textContent = r.status === "pass" ? "✓" : r.status === "warn" ? "!" : "✗";
    const name = document.createElement("span"); name.textContent = r.name;
    li.append(mark, name);
    if (r.detail) { const d = document.createElement("span"); d.className = "detail"; d.textContent = r.detail; li.append(d); }
    return li;
  }

  async function insertTest() {
    const list = $("testResults");
    list.innerHTML = "";
    try {
      const c = SnapTest.sampleImage();
      const bytes = await insertCanvas(c);
      list.appendChild(resultRow({ name: "Test image inserted at your cursor", status: "pass", detail: `${c.width} × ${c.height} px · ${kb(bytes)}. Check it in the email, then delete it.` }));
    } catch (e) {
      list.appendChild(resultRow({ name: "Couldn't insert the test image", status: "fail", detail: e.message || String(e) }));
    }
  }

  // ---------------------------------------------------------------- helpers

  function reset() {
    editor.base = null;
    $("editor").hidden = true;
    $("empty").hidden = false;
    $("effectsMenu").hidden = true;
    setStatus("");
    $("drop").focus();
  }

  function setButtons(on) { document.querySelectorAll("#editor button").forEach((b) => (b.disabled = !on)); if (on && lastInfo) update(lastInfo); }

  function setStatus(text, kind) {
    const s = $("status");
    s.textContent = text;
    s.className = "status" + (kind ? " " + kind : "");
  }

  function toBase64(bytes) {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  function kb(n) { return n < 1024 * 1024 ? Math.max(1, Math.round(n / 1024)) + " KB" : (n / 1048576).toFixed(1) + " MB"; }
})();
