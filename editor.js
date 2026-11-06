/*
 * SnapSize editor — Greenshot-style annotation tools on a true-size preview.
 * All shapes are stored in ORIGINAL image pixels, so they scale with the image and are
 * drawn crisply at whatever final size you choose.
 */
(function (global) {
  "use strict";

  const SHADOW = 12; // css px of margin added by the drop-shadow effect
  const FONT = (px) => `600 ${px}px "Segoe UI", system-ui, -apple-system, sans-serif`;
  const DRAW_TOOLS = ["select", "rect", "ellipse", "line", "arrow", "free", "text", "bubble", "step", "highlight", "obfuscate", "crop"];

  // ---------------------------------------------------------------- pure helpers (also used by the self-test)

  function canvas(w, h) {
    const c = document.createElement("canvas");
    c.width = Math.max(1, Math.round(w));
    c.height = Math.max(1, Math.round(h));
    return c;
  }

  /** High-quality downscale: halve repeatedly, then one final step to the exact size. */
  function resample(src, w, h) {
    w = Math.max(1, Math.round(w)); h = Math.max(1, Math.round(h));
    const out = canvas(w, h);
    const og = out.getContext("2d");
    if (w === src.width && h === src.height) { og.drawImage(src, 0, 0); return out; }
    let cur = src;
    while (cur.width / 2 >= w && cur.height / 2 >= h) {
      const c = canvas(Math.floor(cur.width / 2), Math.floor(cur.height / 2));
      const g = c.getContext("2d");
      g.imageSmoothingQuality = "high";
      g.drawImage(cur, 0, 0, c.width, c.height);
      cur = c;
    }
    og.imageSmoothingEnabled = true;
    og.imageSmoothingQuality = "high";
    og.drawImage(cur, 0, 0, w, h);
    return out;
  }

  function grayscale(c) {
    const g = c.getContext("2d");
    const id = g.getImageData(0, 0, c.width, c.height), d = id.data;
    for (let i = 0; i < d.length; i += 4) {
      const y = Math.round(0.2126 * d[i] + 0.7152 * d[i + 1] + 0.0722 * d[i + 2]);
      d[i] = d[i + 1] = d[i + 2] = y;
    }
    g.putImageData(id, 0, 0);
  }

  /** Replaces a region with flat colour blocks (obfuscate). */
  function pixelate(g, x, y, w, h, block) {
    const cw = g.canvas.width, ch = g.canvas.height;
    let x0 = Math.max(0, Math.floor(x)), y0 = Math.max(0, Math.floor(y));
    const x1 = Math.min(cw, Math.ceil(x + w)), y1 = Math.min(ch, Math.ceil(y + h));
    const rw = x1 - x0, rh = y1 - y0;
    if (rw < 1 || rh < 1) return;
    block = Math.max(2, Math.round(block));
    const id = g.getImageData(x0, y0, rw, rh), d = id.data;
    for (let by = 0; by < rh; by += block) {
      for (let bx = 0; bx < rw; bx += block) {
        const ex = Math.min(bx + block, rw), ey = Math.min(by + block, rh);
        let r = 0, gg = 0, b = 0, a = 0, n = 0;
        for (let yy = by; yy < ey; yy++) for (let xx = bx; xx < ex; xx++) {
          const i = (yy * rw + xx) * 4; r += d[i]; gg += d[i + 1]; b += d[i + 2]; a += d[i + 3]; n++;
        }
        r = Math.round(r / n); gg = Math.round(gg / n); b = Math.round(b / n); a = Math.round(a / n);
        for (let yy = by; yy < ey; yy++) for (let xx = bx; xx < ex; xx++) {
          const i = (yy * rw + xx) * 4; d[i] = r; d[i + 1] = gg; d[i + 2] = b; d[i + 3] = a;
        }
      }
    }
    g.putImageData(id, x0, y0);
  }

  function wrapLines(g, text, maxW) {
    const out = [];
    for (const para of String(text || "").split("\n")) {
      if (!maxW) { out.push(para); continue; }
      let line = "";
      for (const word of para.split(" ")) {
        const t = line ? line + " " + word : word;
        if (line && g.measureText(t).width > maxW) { out.push(line); line = word; } else line = t;
      }
      out.push(line);
    }
    return out;
  }

  function roundRectPath(g, x, y, w, h, r) {
    r = Math.min(r, w / 2, h / 2);
    g.moveTo(x + r, y);
    g.arcTo(x + w, y, x + w, y + h, r);
    g.arcTo(x + w, y + h, x, y + h, r);
    g.arcTo(x, y + h, x, y, r);
    g.arcTo(x, y, x + w, y, r);
  }

  /** Draws one annotation. k = output pixels per original-image pixel. */
  function drawShape(g, s, k) {
    g.save();
    const lw = Math.max(1, s.width * k);
    g.lineWidth = lw; g.strokeStyle = s.color; g.fillStyle = s.color;
    g.lineCap = "round"; g.lineJoin = "round";
    const x1 = s.x1 * k, y1 = s.y1 * k, x2 = (s.x2 ?? s.x1) * k, y2 = (s.y2 ?? s.y1) * k;
    const rx = Math.min(x1, x2), ry = Math.min(y1, y2), rw = Math.abs(x2 - x1), rh = Math.abs(y2 - y1);
    const fs = (s.fontSize || 16) * k;

    switch (s.type) {
      case "rect":
        if (s.fill) g.fillRect(rx, ry, rw, rh); else g.strokeRect(rx, ry, rw, rh);
        break;
      case "ellipse":
        g.beginPath(); g.ellipse(rx + rw / 2, ry + rh / 2, rw / 2, rh / 2, 0, 0, Math.PI * 2);
        if (s.fill) g.fill(); else g.stroke();
        break;
      case "line":
        g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke();
        break;
      case "arrow": {
        const ang = Math.atan2(y2 - y1, x2 - x1), hl = Math.max(10 * k, lw * 3.5), hw = hl * 0.55;
        const cos = Math.cos(ang), sin = Math.sin(ang);
        g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2 - cos * hl * 0.7, y2 - sin * hl * 0.7); g.stroke();
        g.beginPath();
        g.moveTo(x2, y2);
        g.lineTo(x2 - hl * cos + hw * sin, y2 - hl * sin - hw * cos);
        g.lineTo(x2 - hl * cos - hw * sin, y2 - hl * sin + hw * cos);
        g.closePath(); g.fill();
        break;
      }
      case "free": {
        const p = s.points || [];
        if (!p.length) break;
        g.beginPath(); g.moveTo(p[0][0] * k, p[0][1] * k);
        if (p.length === 1) g.lineTo(p[0][0] * k + 0.1, p[0][1] * k);
        for (let i = 1; i < p.length; i++) g.lineTo(p[i][0] * k, p[i][1] * k);
        g.stroke();
        break;
      }
      case "text": {
        g.font = FONT(fs); g.textBaseline = "top";
        wrapLines(g, s.text, 0).forEach((l, i) => g.fillText(l, x1, y1 + i * fs * 1.25));
        break;
      }
      case "bubble": {
        const pad = 6 * k * (s.width ? 1 : 1), tailH = Math.max(14 * k, rh * 0.35);
        g.beginPath();
        roundRectPath(g, rx, ry, rw, rh, 8 * k);
        g.moveTo(rx + rw * 0.22, ry + rh);
        g.lineTo(rx + rw * 0.1, ry + rh + tailH);
        g.lineTo(rx + rw * 0.4, ry + rh);
        g.fillStyle = "#ffffff"; g.fill();
        g.stroke();
        // cover the seam between box and tail
        g.fillRect(rx + rw * 0.22 + lw / 2, ry + rh - lw / 2 - 0.5, rw * 0.18 - lw, lw + 1);
        g.fillStyle = s.color; g.font = FONT(fs); g.textBaseline = "top";
        g.save(); g.beginPath(); g.rect(rx, ry, rw, rh); g.clip();
        wrapLines(g, s.text, rw - pad * 2).forEach((l, i) => g.fillText(l, rx + pad, ry + pad + i * fs * 1.25));
        g.restore();
        break;
      }
      case "step": {
        const r = fs * 0.85;
        g.beginPath(); g.arc(x1, y1, r, 0, Math.PI * 2); g.fill();
        g.fillStyle = "#ffffff"; g.font = `700 ${Math.round(r * 1.15)}px "Segoe UI", system-ui, sans-serif`;
        g.textAlign = "center"; g.textBaseline = "middle";
        g.fillText(String(s.n), x1, y1 + r * 0.05);
        break;
      }
      case "highlight":
        g.globalCompositeOperation = "multiply";
        g.fillRect(rx, ry, rw, rh);
        break;
      case "obfuscate":
        pixelate(g, rx, ry, rw, rh, (s.pixel || 8) * k);
        break;
    }
    g.restore();
  }

  const measureCtx = (() => { let c; return () => (c = c || document.createElement("canvas").getContext("2d")); })();

  /** Bounding box in original-image pixels. */
  function bbox(s) {
    const pad = (s.width || 2) / 2 + 2;
    switch (s.type) {
      case "free": {
        const xs = s.points.map((p) => p[0]), ys = s.points.map((p) => p[1]);
        return { x: Math.min(...xs) - pad, y: Math.min(...ys) - pad, w: Math.max(...xs) - Math.min(...xs) + pad * 2, h: Math.max(...ys) - Math.min(...ys) + pad * 2 };
      }
      case "text": {
        const g = measureCtx(); g.font = FONT(s.fontSize);
        const lines = wrapLines(g, s.text, 0);
        const w = Math.max(10, ...lines.map((l) => g.measureText(l).width));
        return { x: s.x1, y: s.y1, w, h: lines.length * s.fontSize * 1.25 };
      }
      case "step": {
        const r = s.fontSize * 0.85;
        return { x: s.x1 - r, y: s.y1 - r, w: r * 2, h: r * 2 };
      }
      default: {
        const x = Math.min(s.x1, s.x2), y = Math.min(s.y1, s.y2);
        let h = Math.abs(s.y2 - s.y1);
        if (s.type === "bubble") h += Math.max(14, h * 0.35);
        return { x: x - pad, y: y - pad, w: Math.abs(s.x2 - s.x1) + pad * 2, h: h + pad * 2 };
      }
    }
  }

  function moveShape(s, dx, dy) {
    s.x1 += dx; s.y1 += dy;
    if (s.x2 != null) { s.x2 += dx; s.y2 += dy; }
    if (s.points) s.points = s.points.map((p) => [p[0] + dx, p[1] + dy]);
  }

  // ---------------------------------------------------------------- the interactive editor

  class Editor {
    /**
     * @param {{canvas:HTMLCanvasElement, frame:HTMLElement, settings:()=>object, onChange:(info:object)=>void}} o
     */
    constructor(o) {
      this.cv = o.canvas;
      this.frame = o.frame;
      this.settings = o.settings;
      this.onChange = o.onChange || (() => {});
      this.base = null;
      this.shapes = [];
      this.history = [];
      this.scale = 1;
      this.tool = "select";
      this.counter = 0;
      this.fx = { shadow: false, grayscale: false, border: false };
      this.props = { color: "#e81123", width: 3, fontSize: 16, fill: false };
      this.sel = null; this.temp = null; this.crop = null; this.drag = null;
      this._cache = null; this._text = null;
      this._bind();
    }

    // ----- state

    load(base) {
      const st = this.settings();
      this.base = base;
      this.shapes = []; this.history = []; this.counter = 0;
      this.sel = this.temp = this.crop = this.drag = null; this._cache = null;
      this._closeText(false);
      this.fx = { shadow: !!st.shadow, grayscale: !!st.grayscale, border: !!st.border };
      this.props = { color: st.color, width: st.lineWidth, fontSize: st.fontSize, fill: false };
      this.tool = "select";
      this.scale = st.startSize === "actual" ? 1 : Math.min(1, st.fitWidth / base.width);
      this.render();
    }

    snapshot() {
      this.history.push({ base: this.base, shapes: JSON.parse(JSON.stringify(this.shapes)), counter: this.counter, fx: { ...this.fx } });
      if (this.history.length > 60) this.history.shift();
    }

    undo() {
      const h = this.history.pop();
      if (!h) return false;
      this.base = h.base; this.shapes = h.shapes; this.counter = h.counter; this.fx = h.fx;
      this.sel = this.crop = null; this._cache = null;
      this.render();
      return true;
    }

    setTool(id) {
      if (!DRAW_TOOLS.includes(id)) return;
      this._closeText(true);
      const st = this.settings();
      if (id === "highlight" && this.tool !== "highlight") { this._lineColor = this.props.color; this.props.color = st.highlight; }
      else if (id !== "highlight" && this.tool === "highlight") { this.props.color = this._lineColor || st.color; }
      if (id !== "crop") this.crop = null;
      if (id !== "select") this.sel = null;
      this.tool = id;
      this.cv.style.cursor = id === "select" ? "default" : id === "text" ? "text" : "crosshair";
      this.render();
    }

    /** Apply a property change; also updates the selected shape. */
    setProp(name, value) {
      this.props[name] = value;
      if (this.sel) {
        this.snapshot();
        const map = { width: "width", fontSize: "fontSize" };
        if (name === "color") this.sel.color = value;
        else if (name === "fill") this.sel.fill = value;
        else if (map[name]) this.sel[map[name]] = value / this.scale;
        this.render();
      }
    }

    setScale(s) {
      this._closeText(true);
      s = Math.max(0.05, Math.min(4, s));
      if (Math.abs(s - 1) < 0.03) s = 1;
      this.scale = s;
      this.render();
    }

    setWidthPx(w) { if (this.base) this.setScale(w / this.base.width); }

    toggleEffect(name, on) { this.snapshot(); this.fx[name] = on; this._cache = null; this.render(); }

    deleteSelected() {
      if (!this.sel) return false;
      this.snapshot();
      this.shapes = this.shapes.filter((s) => s !== this.sel);
      this.sel = null; this.render();
      return true;
    }

    /** Burn annotations into the image (needed before crop/rotate). */
    flatten() {
      if (!this.shapes.length) return;
      const c = canvas(this.base.width, this.base.height);
      const g = c.getContext("2d");
      g.drawImage(this.base, 0, 0);
      for (const s of this.shapes) drawShape(g, s, 1);
      this.base = c; this.shapes = []; this.sel = null; this._cache = null;
    }

    rotate(dir) {
      this._closeText(true);
      this.snapshot();
      this.flatten();
      const b = this.base, r = canvas(b.height, b.width), g = r.getContext("2d");
      if (dir > 0) { g.translate(r.width, 0); g.rotate(Math.PI / 2); }
      else { g.translate(0, r.height); g.rotate(-Math.PI / 2); }
      g.drawImage(b, 0, 0);
      this.base = r; this._cache = null; this.crop = null;
      this.render();
    }

    applyCrop() {
      const c = this.crop;
      if (!c) return false;
      const x = Math.max(0, Math.round(Math.min(c.x1, c.x2))), y = Math.max(0, Math.round(Math.min(c.y1, c.y2)));
      const w = Math.min(this.base.width - x, Math.round(Math.abs(c.x2 - c.x1)));
      const h = Math.min(this.base.height - y, Math.round(Math.abs(c.y2 - c.y1)));
      this.crop = null;
      if (w < 4 || h < 4) { this.render(); return false; }
      this.snapshot();
      this.flatten();
      const out = canvas(w, h);
      out.getContext("2d").drawImage(this.base, x, y, w, h, 0, 0, w, h);
      this.base = out; this._cache = null;
      this.tool = "select";
      this.render();
      return true;
    }

    cancelCrop() { this.crop = null; this.render(); }

    size() {
      const w = Math.max(1, Math.round(this.base.width * this.scale));
      const h = Math.max(1, Math.round(this.base.height * this.scale));
      const m = this.fx.shadow ? SHADOW : 0;
      return { w, h, outW: w + m, outH: h + m };
    }

    // ----- rendering

    /** Builds the image at `scale` (email px per image px) × `dpr` (screen px per email px). */
    compose(scale, dpr, withTemp) {
      const w = Math.max(1, Math.round(this.base.width * scale)), h = Math.max(1, Math.round(this.base.height * scale));
      const pw = Math.max(1, Math.round(w * dpr)), ph = Math.max(1, Math.round(h * dpr));
      const key = pw + "x" + ph + "|" + this.fx.grayscale;
      if (!this._cache || this._cache.key !== key || this._cache.base !== this.base) {
        const b = resample(this.base, pw, ph);
        if (this.fx.grayscale) grayscale(b);
        this._cache = { key, base: this.base, canvas: b };
      }
      const layer = canvas(pw, ph);
      const g = layer.getContext("2d");
      g.drawImage(this._cache.canvas, 0, 0);
      const k = pw / this.base.width;
      for (const s of this.shapes) drawShape(g, s, k);
      if (withTemp && this.temp) drawShape(g, this.temp, k);
      if (this.fx.border) {
        const lw = Math.max(1, Math.round(dpr));
        g.strokeStyle = "#aaaaaa"; g.lineWidth = lw;
        g.strokeRect(lw / 2, lw / 2, pw - lw, ph - lw);
      }
      if (!this.fx.shadow) return layer;
      const m = Math.round(SHADOW * dpr);
      const out = canvas(pw + m, ph + m), o = out.getContext("2d");
      o.fillStyle = "#ffffff"; o.fillRect(0, 0, out.width, out.height);
      o.shadowColor = "rgba(0,0,0,0.35)"; o.shadowBlur = 6 * dpr; o.shadowOffsetX = o.shadowOffsetY = 3 * dpr;
      o.drawImage(layer, 0, 0);
      return out;
    }

    /** Final image at exactly the size shown. */
    renderOutput() {
      this._closeText(true);
      return this.compose(this.scale, 1, false);
    }

    render() {
      if (!this.base) return;
      const dpr = window.devicePixelRatio || 1;
      const img = this.compose(this.scale, dpr, true);
      const cv = this.cv;
      if (cv.width !== img.width || cv.height !== img.height) { cv.width = img.width; cv.height = img.height; }
      cv.style.width = img.width / dpr + "px";
      cv.style.height = img.height / dpr + "px";
      const g = cv.getContext("2d");
      g.clearRect(0, 0, cv.width, cv.height);
      g.drawImage(img, 0, 0);

      const k = this.scale * dpr;
      if (this.sel) {
        const b = bbox(this.sel);
        g.save(); g.setLineDash([4 * dpr, 3 * dpr]); g.strokeStyle = "#0f6cbd"; g.lineWidth = dpr;
        g.strokeRect(b.x * k, b.y * k, b.w * k, b.h * k); g.restore();
      }
      if (this.crop) {
        const c = this.crop, x = Math.min(c.x1, c.x2) * k, y = Math.min(c.y1, c.y2) * k;
        const w = Math.abs(c.x2 - c.x1) * k, h = Math.abs(c.y2 - c.y1) * k;
        g.save();
        g.fillStyle = "rgba(0,0,0,0.45)";
        g.beginPath(); g.rect(0, 0, cv.width, cv.height); g.rect(x, y, w, h); g.fill("evenodd");
        g.setLineDash([5 * dpr, 4 * dpr]); g.strokeStyle = "#ffffff"; g.lineWidth = dpr; g.strokeRect(x, y, w, h);
        g.restore();
      }
      const sz = this.size();
      this.onChange({ ...sz, scale: this.scale, tool: this.tool, cropPending: !!(this.crop && !this.drag), canUndo: this.history.length > 0, selected: !!this.sel, fx: { ...this.fx } });
    }

    // ----- pointer interaction

    _toImg(e) {
      const r = this.cv.getBoundingClientRect();
      return [(e.clientX - r.left) / this.scale, (e.clientY - r.top) / this.scale];
    }

    _hit(p) {
      const tol = 4 / this.scale;
      for (let i = this.shapes.length - 1; i >= 0; i--) {
        const b = bbox(this.shapes[i]);
        if (p[0] >= b.x - tol && p[0] <= b.x + b.w + tol && p[1] >= b.y - tol && p[1] <= b.y + b.h + tol) return this.shapes[i];
      }
      return null;
    }

    _newShape(type, p) {
      const z = this.scale;
      return {
        type, x1: p[0], y1: p[1], x2: p[0], y2: p[1],
        color: this.props.color, width: this.props.width / z, fontSize: this.props.fontSize / z,
        fill: !!this.props.fill, pixel: (this.settings().pixel || 8) / z
      };
    }

    _bind() {
      const cv = this.cv;
      cv.style.touchAction = "none";

      cv.addEventListener("pointerdown", (e) => {
        if (!this.base || e.button > 0) return;
        e.preventDefault(); // keep keyboard focus where it is (e.g. a new text box)
        if (this._text) { this._closeText(true); return; }
        const p = this._toImg(e);
        cv.setPointerCapture(e.pointerId);
        switch (this.tool) {
          case "select": {
            this.sel = this._hit(p);
            if (this.sel) { this.snapshot(); this.drag = { mode: "move", last: p, moved: false }; }
            this.render();
            break;
          }
          case "text":
            this._openText(p, null);
            break;
          case "step": {
            this.snapshot();
            const s = this._newShape("step", p);
            s.n = ++this.counter;
            this.shapes.push(s);
            this.render();
            break;
          }
          case "crop":
            this.crop = { x1: p[0], y1: p[1], x2: p[0], y2: p[1] };
            this.drag = { mode: "crop" };
            break;
          case "free":
            this.temp = this._newShape("free", p);
            this.temp.points = [p];
            this.drag = { mode: "draw" };
            break;
          default:
            this.temp = this._newShape(this.tool, p);
            this.drag = { mode: "draw" };
        }
      });

      cv.addEventListener("pointermove", (e) => {
        if (!this.drag) return;
        const p = this._toImg(e);
        if (this.drag.mode === "move") {
          moveShape(this.sel, p[0] - this.drag.last[0], p[1] - this.drag.last[1]);
          this.drag.last = p; this.drag.moved = true;
        } else if (this.drag.mode === "crop") {
          this.crop.x2 = Math.max(0, Math.min(this.base.width, p[0]));
          this.crop.y2 = Math.max(0, Math.min(this.base.height, p[1]));
        } else if (this.temp.type === "free") {
          this.temp.points.push(p);
        } else {
          if (e.shiftKey && (this.temp.type === "line" || this.temp.type === "arrow")) {
            // snap to 45°
            const dx = p[0] - this.temp.x1, dy = p[1] - this.temp.y1;
            const a = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4), len = Math.hypot(dx, dy);
            p[0] = this.temp.x1 + Math.cos(a) * len; p[1] = this.temp.y1 + Math.sin(a) * len;
          }
          this.temp.x2 = p[0]; this.temp.y2 = p[1];
        }
        this.render();
      });

      const end = () => {
        if (!this.drag) return;
        const d = this.drag; this.drag = null;
        if (d.mode === "move") {
          if (!d.moved) this.history.pop(); // a click, not a move
        } else if (d.mode === "crop") {
          const c = this.crop;
          if (Math.abs(c.x2 - c.x1) * this.scale < 4 || Math.abs(c.y2 - c.y1) * this.scale < 4) this.crop = null;
        } else if (this.temp) {
          const t = this.temp; this.temp = null;
          const tiny = t.type === "free" ? t.points.length < 2 : Math.hypot(t.x2 - t.x1, t.y2 - t.y1) * this.scale < 4;
          if (!tiny) {
            this.snapshot();
            this.shapes.push(t);
            if (t.type === "bubble") { t.text = ""; this._openText([Math.min(t.x1, t.x2), Math.min(t.y1, t.y2)], t); }
          }
        }
        this.render();
      };
      cv.addEventListener("pointerup", end);
      cv.addEventListener("pointercancel", end);

      cv.addEventListener("dblclick", (e) => {
        if (this.tool !== "select") return;
        const s = this._hit(this._toImg(e));
        if (s && (s.type === "text" || s.type === "bubble")) this._openText([s.x1, s.y1], s);
      });
    }

    // ----- inline text box (task panes can't use prompt())

    _openText(p, shape) {
      this._closeText(true);
      const ta = document.createElement("textarea");
      const z = this.scale;
      const fsImg = shape ? shape.fontSize : this.props.fontSize / z;
      const color = shape ? shape.color : this.props.color;
      let x = p[0], y = p[1], w = null, h = null;
      if (shape && shape.type === "bubble") {
        x = Math.min(shape.x1, shape.x2); y = Math.min(shape.y1, shape.y2);
        w = Math.abs(shape.x2 - shape.x1); h = Math.abs(shape.y2 - shape.y1);
      }
      Object.assign(ta.style, {
        position: "absolute", left: x * z + 3 + "px", top: y * z + 3 + "px",
        font: FONT(fsImg * z), color, lineHeight: "1.25",
        width: w ? w * z + "px" : Math.max(120, 200) + "px", height: h ? h * z + "px" : fsImg * z * 1.25 * 2 + 8 + "px",
        background: "rgba(255,255,255,0.9)", border: "1px dashed #0f6cbd", outline: "none",
        resize: "none", padding: shape && shape.type === "bubble" ? 6 * z * 1 + "px" : "0", margin: "0", zIndex: 5
      });
      ta.value = shape ? shape.text || "" : "";
      ta.setAttribute("aria-label", "Annotation text");
      this.frame.appendChild(ta);
      this._text = { ta, shape, p };
      ta.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); this._closeText(true); }
        else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this._closeText(false); }
        else e.stopPropagation();
      });
      ta.addEventListener("blur", () => setTimeout(() => { if (this._text && this._text.ta === ta && document.activeElement !== ta) this._closeText(true); }, 0));
      ta.focus();
      setTimeout(() => { if (this._text && this._text.ta === ta) ta.focus(); }, 0);
    }

    _closeText(commit) {
      const t = this._text;
      if (!t) return;
      this._text = null;
      const text = t.ta.value.replace(/\s+$/, "");
      t.ta.remove();
      if (commit) {
        if (t.shape) {
          if (t.shape.type === "text" && !text) this.shapes = this.shapes.filter((s) => s !== t.shape);
          else t.shape.text = text;
        } else if (text) {
          this.snapshot();
          const s = this._newShape("text", t.p);
          s.text = text;
          this.shapes.push(s);
        }
      }
      this.render();
    }
  }

  global.SnapEditor = { Editor, drawShape, pixelate, resample, grayscale, bbox, canvas, SHADOW };
})(window);
