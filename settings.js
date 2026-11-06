/* global Office */
/* SnapSize settings: which tools appear (and in what order), default colours/sizes, effects. */
(function (global) {
  "use strict";

  const TOOLS = [
    { id: "select",    label: "Select / move",   key: "Escape", hint: "Esc" },
    { id: "rect",      label: "Rectangle",       key: "r" },
    { id: "ellipse",   label: "Ellipse",         key: "e" },
    { id: "line",      label: "Line",            key: "l" },
    { id: "arrow",     label: "Arrow",           key: "a" },
    { id: "free",      label: "Freehand",        key: "f" },
    { id: "text",      label: "Text",            key: "t" },
    { id: "bubble",    label: "Speech bubble",   key: "s" },
    { id: "step",      label: "Step number",     key: "i" },
    { id: "highlight", label: "Highlight",       key: "h" },
    { id: "obfuscate", label: "Obfuscate (blur out)", key: "o" },
    { id: "effects",   label: "Effects (shadow, greyscale, border)", key: null },
    { id: "crop",      label: "Crop",            key: "c" },
    { id: "rotateL",   label: "Rotate left",     key: null },
    { id: "rotateR",   label: "Rotate right",    key: null },
    { id: "resize",    label: "Resize to exact width", key: null },
  ];

  const DEFAULTS = {
    fitWidth: 600,
    startSize: "fit",          // "fit" | "actual"
    color: "#e81123",
    highlight: "#fff200",
    lineWidth: 3,
    fontSize: 16,
    pixel: 8,
    shadow: false,
    grayscale: false,
    border: false,
    toolbar: TOOLS.map((t) => ({ id: t.id, visible: true })),
  };

  const KEY = "snapsize";

  function normalise(s) {
    const out = { ...DEFAULTS, ...s };
    const known = new Set(TOOLS.map((t) => t.id));
    const seen = new Set();
    const list = [];
    for (const t of Array.isArray(out.toolbar) ? out.toolbar : []) {
      if (t && known.has(t.id) && !seen.has(t.id)) { list.push({ id: t.id, visible: t.visible !== false }); seen.add(t.id); }
    }
    for (const t of TOOLS) if (!seen.has(t.id)) list.push({ id: t.id, visible: true }); // new tools in updates
    out.toolbar = list;
    out.fitWidth = clamp(out.fitWidth, 100, 2000, 600);
    out.lineWidth = clamp(out.lineWidth, 1, 20, 3);
    out.fontSize = clamp(out.fontSize, 8, 72, 16);
    out.pixel = clamp(out.pixel, 2, 40, 8);
    if (!/^#[0-9a-f]{6}$/i.test(out.color)) out.color = DEFAULTS.color;
    if (!/^#[0-9a-f]{6}$/i.test(out.highlight)) out.highlight = DEFAULTS.highlight;
    if (out.startSize !== "actual") out.startSize = "fit";
    return out;
  }

  function clamp(v, lo, hi, d) { v = Number(v); return Number.isFinite(v) ? Math.max(lo, Math.min(hi, Math.round(v))) : d; }

  function roaming() {
    try { return (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) || null; } catch (e) { return null; }
  }

  function load() {
    let s = null;
    const rs = roaming();
    if (rs) {
      try {
        const v = rs.get(KEY);
        if (v) s = typeof v === "string" ? JSON.parse(v) : v;
        else if (rs.get("fitWidth")) s = { fitWidth: rs.get("fitWidth") }; // v1.0 setting
      } catch (e) { /* ignore */ }
    }
    if (!s) { try { s = JSON.parse(localStorage.getItem(KEY) || "null"); } catch (e) { /* ignore */ } }
    return normalise(s || {});
  }

  /** Saves locally and to the mailbox (roams to all your devices). Resolves to true if roaming save worked. */
  function save(s) {
    try { localStorage.setItem(KEY, JSON.stringify(s)); } catch (e) { /* ignore */ }
    const rs = roaming();
    if (!rs) return Promise.resolve(false);
    return new Promise((resolve) => {
      try {
        rs.set(KEY, JSON.stringify(s));
        rs.saveAsync((r) => resolve(r.status === "succeeded"));
      } catch (e) { resolve(false); }
    });
  }

  function toolInfo(id) { return TOOLS.find((t) => t.id === id); }

  // ---------------------------------------------------------------- settings screen

  /**
   * Builds the settings form inside `root`. `get()` returns current settings,
   * `set(newSettings)` is called on every change.
   */
  function mountUI(root, get, set, icons) {
    root.innerHTML = "";
    const s = get();

    const h = (tag, attrs = {}, kids = []) => {
      const el = document.createElement(tag);
      for (const [k, v] of Object.entries(attrs)) {
        if (k === "class") el.className = v; else if (k === "text") el.textContent = v;
        else if (k.startsWith("on")) el.addEventListener(k.slice(2), v); else el.setAttribute(k, v);
      }
      for (const c of [].concat(kids)) if (c) el.append(c);
      return el;
    };
    const commit = (patch) => set(normalise({ ...get(), ...patch }));

    // Toolbar list
    root.append(h("h2", { text: "Toolbar" }), h("p", { class: "hint", text: "Tick the tools you want. Use the arrows to change their order." }));
    const list = h("ul", { class: "toolList" });
    const draw = () => {
      list.innerHTML = "";
      get().toolbar.forEach((t, i, arr) => {
        const info = toolInfo(t.id);
        const cb = h("input", { type: "checkbox", id: "tb-" + t.id, "aria-label": info.label });
        cb.checked = t.visible;
        cb.addEventListener("change", () => {
          const tb = get().toolbar.map((x) => (x.id === t.id ? { ...x, visible: cb.checked } : x));
          if (!tb.some((x) => x.visible)) { cb.checked = true; return; }
          commit({ toolbar: tb });
        });
        const move = (d) => {
          const tb = get().toolbar.slice(); const j = i + d;
          if (j < 0 || j >= tb.length) return;
          [tb[i], tb[j]] = [tb[j], tb[i]];
          commit({ toolbar: tb }); draw();
          const btn = list.children[j] && list.children[j].querySelector(d < 0 ? ".up" : ".down");
          if (btn && !btn.disabled) btn.focus();
        };
        const icon = h("span", { class: "ticon", "aria-hidden": "true" });
        icon.innerHTML = icons[t.id] || "";
        list.append(h("li", {}, [
          cb,
          h("label", { for: "tb-" + t.id }, [icon, info.label + (info.hint || info.key ? `  ·  ${info.hint || info.key.toUpperCase()}` : "")]),
          Object.assign(h("button", { class: "up", title: "Move up", "aria-label": "Move " + info.label + " up", onclick: () => move(-1) }, "↑"), { disabled: i === 0 }),
          Object.assign(h("button", { class: "down", title: "Move down", "aria-label": "Move " + info.label + " down", onclick: () => move(1) }, "↓"), { disabled: i === arr.length - 1 }),
        ]));
      });
    };
    draw();
    root.append(list);

    // Defaults
    root.append(h("h2", { text: "Defaults" }));
    const grid = h("div", { class: "grid" });
    const field = (label, input) => grid.append(h("label", {}, [h("span", { text: label }), input]));
    const inp = (attrs, key, parse = (v) => v) => {
      const el = h("input", attrs);
      if (attrs.type === "checkbox") el.checked = !!s[key]; else el.value = s[key];
      el.addEventListener("change", () => { commit({ [key]: attrs.type === "checkbox" ? el.checked : parse(el.value) }); el.value = attrs.type === "checkbox" ? el.value : get()[key]; });
      return el;
    };
    field("Line & text colour", inp({ type: "color" }, "color"));
    field("Highlight colour", inp({ type: "color" }, "highlight"));
    field("Line width (px)", inp({ type: "number", min: 1, max: 20 }, "lineWidth", Number));
    field("Text size (px)", inp({ type: "number", min: 8, max: 72 }, "fontSize", Number));
    field("Obfuscate block (px)", inp({ type: "number", min: 2, max: 40 }, "pixel", Number));
    field("Fit width (px)", inp({ type: "number", min: 100, max: 2000, step: 10 }, "fitWidth", Number));
    const start = h("select", {}, [h("option", { value: "fit", text: "Fit to width (if larger)" }), h("option", { value: "actual", text: "100% (actual pixels)" })]);
    start.value = s.startSize;
    start.addEventListener("change", () => commit({ startSize: start.value }));
    field("Start size", start);
    root.append(grid);

    root.append(h("h2", { text: "Effects on by default" }));
    const fx = h("div", { class: "checks" });
    for (const [key, label] of [["shadow", "Drop shadow"], ["grayscale", "Greyscale"], ["border", "Thin border"]]) {
      fx.append(h("label", { class: "check" }, [inp({ type: "checkbox" }, key), label]));
    }
    root.append(fx);

    root.append(h("button", {
      class: "linkbtn", text: "Reset all settings to defaults",
      onclick: () => { set(normalise({})); mountUI(root, get, set, icons); },
    }));
  }

  global.SnapSettings = { TOOLS, DEFAULTS, load, save, normalise, toolInfo, mountUI };
})(window);
