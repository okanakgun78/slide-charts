// Minimal stand-in for Office.js / PowerPoint used by the end-to-end test.
// Operations apply eagerly to an in-memory deck; reading a property requires load() + sync(),
// as in the real API (PropertyNotLoaded otherwise). Reproduces the Mac quirk measured in the pilot:
// a 0 width/height in add options becomes 72 pt.
(() => {
  let nextId = 100;
  const newId = () => String(nextId++);
  const deck = {
    w: 960, h: 540,
    slides: [{ id: "256#1", shapes: [], bg: "#FFFFFF" }],
    selSlides: ["256#1"],
    selShapes: [],
  };
  const handlers = [];
  window.__deck = deck;
  window.__errors = [];

  const fireSelection = () => setTimeout(() => handlers.forEach((h) => h({})), 10);
  window.__select = (ids) => { deck.selShapes = ids.slice(); fireSelection(); };

  const slideOf = (id) => deck.slides.find((s) => s.id === id);
  function findShape(id) {
    for (const s of deck.slides) {
      const walk = (list) => { for (const sh of list) { if (sh.id === id) return sh; if (sh.children) { const r = walk(sh.children); if (r) return r; } } return null; };
      const r = walk(s.shapes);
      if (r) return { slide: s, shape: r };
    }
    return null;
  }

  class NotLoaded extends Error { constructor(p) { super(`PropertyNotLoaded: ${p}`); this.code = "PropertyNotLoaded"; } }

  // Context: tracks pending loads and deferred errors.
  class Ctx {
    constructor() { this.pending = []; this.errors = []; }
    async sync() {
      await new Promise((r) => setTimeout(r, 1));
      for (const [obj, props] of this.pending) props.forEach((p) => obj._loaded.add(p));
      this.pending = [];
      if (this.errors.length) { const e = this.errors[0]; this.errors = []; throw e; }
    }
    queueLoad(obj, spec) {
      const props = typeof spec === "string" ? spec.split(",").map((s) => s.trim()) : Object.keys(spec || {});
      this.pending.push([obj, props]);
      return obj;
    }
  }

  const guard = (obj, prop, value) => { if (!obj._loaded.has(prop) && !obj._loaded.has("*")) throw new NotLoaded(prop); return value; };

  class Result { constructor(ctx, v) { this._loaded = new Set(); this._v = v; ctx.pending.push([this, ["value"]]); } get value() { return guard(this, "value", this._v); } }

  class Tag {
    constructor(ctx, shape, key) { this.ctx = ctx; this._shape = shape; this._key = key; this._loaded = new Set(); }
    get isNullObject() { return !this._shape || !this._shape.tags || !(this._key in this._shape.tags); }
    get value() { return guard(this, "value", this._shape.tags[this._key]); }
    load(p) { return this.ctx.queueLoad(this, p); }
  }
  class Tags {
    constructor(ctx, shape) { this.ctx = ctx; this._shape = shape; }
    add(k, v) { this._shape.tags[k.toUpperCase()] = String(v); }
    delete(k) { delete this._shape.tags[k.toUpperCase()]; }
    getItem(k) { const t = new Tag(this.ctx, this._shape, k.toUpperCase()); if (t.isNullObject) this.ctx.errors.push(new Error("ItemNotFound tag")); return t; }
    getItemOrNullObject(k) { return new Tag(this.ctx, this._shape, k.toUpperCase()); }
  }

  const H = (rec, path, init) => { let o = rec; for (const k of path) o = o[k] ??= {}; Object.assign(o, init ?? {}); return o; };

  class Font { constructor(ctx, rec) { this.ctx = ctx; this._rec = rec; this._loaded = new Set(); }
    load(p) { return this.ctx.queueLoad(this, p); }
    get name() { return guard(this, "name", "Avenir Next LT Pro"); }
    set size(v) { this._rec.size = v; } set color(v) { this._rec.color = v; } set bold(v) { this._rec.bold = v; } set italic(v) { this._rec.italic = v; } }
  class TextFrame {
    constructor(ctx, rec) { this.ctx = ctx; this._rec = rec; }
    set autoSizeSetting(v) {
      this._rec.tf.autoSize = v;
      if (v === "AutoSizeShapeToFitText") { const ls = this._rec.text.split("\n"); this._rec.width = Math.max(...ls.map((l) => l.length)) * 6.6; this._rec.height = ls.length * 14.4; }
    }
    set wordWrap(v) { this._rec.tf.wrap = v; }
    set leftMargin(v) { this._rec.tf.l = v; } set rightMargin(v) { this._rec.tf.r = v; }
    set topMargin(v) { this._rec.tf.t = v; } set bottomMargin(v) { this._rec.tf.b = v; }
    set verticalAlignment(v) { this._rec.tf.valign = v; }
    get textRange() { const rec = this._rec; const ctx = this.ctx; return (this._tr ??= { font: new Font(ctx, H(rec, ["font"])), paragraphFormat: { set horizontalAlignment(v) { rec.align = v; } } }); }
  }

  class Shape {
    constructor(ctx, rec, missing) { this.ctx = ctx; this._rec = rec; this._loaded = new Set(); if (missing) ctx.errors.push(new Error("ItemNotFound shape")); }
    load(p) { return this.ctx.queueLoad(this, p); }
    get isNullObject() { return !this._rec; }
    get id() { return guard(this, "id", this._rec.id); }
    get type() { return guard(this, "type", this._rec.type); }
    get level() { return guard(this, "level", this._rec.parent ? 1 : 0); }
    get name() { return guard(this, "name", this._rec.name); } set name(v) { this._rec.name = v; }
    get left() { return guard(this, "left", this._rec.left); } set left(v) { this._move("left", v); }
    get top() { return guard(this, "top", this._rec.top); } set top(v) { this._move("top", v); }
    get width() { return guard(this, "width", this._rec.width); } set width(v) { this._scale("width", v); }
    get height() { return guard(this, "height", this._rec.height); } set height(v) { this._scale("height", v); }
    get rotation() { return guard(this, "rotation", this._rec.rotation || 0); } set rotation(v) { this._rec.rotation = v; }
    _move(k, v) { const d = v - this._rec[k]; this._rec[k] = v; (this._rec.children || []).forEach((c) => (c[k] += d)); }
    _scale(k, v) {
      const r = this._rec, old = r[k];
      if (r.children && old > 0) { const o = k === "width" ? "left" : "top"; const f = v / old; r.children.forEach((c) => { c[o] = r[o] + (c[o] - r[o]) * f; c[k] *= f; }); }
      r[k] = v;
    }
    get fill() { const rec = this._rec; return { setSolidColor(c) { rec.fill = c; }, clear() { rec.fill = null; } }; }
    get lineFormat() { const rec = H(this._rec, ["line"]); return { set color(v) { rec.color = v; }, set weight(v) { rec.weight = v; }, set visible(v) { rec.visible = v; }, set dashStyle(v) { rec.dash = v; } }; }
    get textFrame() { return (this._tf ??= new TextFrame(this.ctx, this._rec)); }
    get tags() { return new Tags(this.ctx, this._rec); }
    get adjustments() { const rec = this._rec; return { set(i, v) { (rec.adj ??= [])[i] = v; } }; }
    get group() {
      const rec = this._rec, ctx = this.ctx;
      if (rec.type !== "Group") ctx.errors.push(new Error("InvalidArgument: not a group"));
      return {
        shapes: new Shapes(ctx, () => rec.children || [], null),
        ungroup() { const sl = findShape(rec.id).slide; const i = sl.shapes.indexOf(rec); rec.children.forEach((c) => (c.parent = null)); sl.shapes.splice(i, 1, ...rec.children); },
      };
    }
    get parentGroup() { const p = this._rec.parent; return new Shape(this.ctx, p, !p); }
    getParentSlide() { const f = findShape(this._rec.id); return new Slide(this.ctx, f.slide); }
    delete() {
      const f = findShape(this._rec.id);
      if (!f) { this.ctx.errors.push(new Error("ItemNotFound on delete")); return; }
      const list = this._rec.parent ? this._rec.parent.children : f.slide.shapes;
      list.splice(list.indexOf(this._rec), 1);
      if (this._rec.parent && this._rec.parent.children.length === 0) { const g = this._rec.parent; const gl = f.slide.shapes; gl.splice(gl.indexOf(g), 1); }
    }
  }

  class Shapes {
    constructor(ctx, list, slide) { this.ctx = ctx; this._list = list; this._slide = slide; this._loaded = new Set(); }
    load(p) { return this.ctx.queueLoad(this, typeof p === "string" ? p.split(",").map((x) => x.split("/")[0]).join(",") : p); }
    get items() { return guard(this, "items", this._list().map((r) => { const s = new Shape(this.ctx, r); s._loaded.add("*"); return s; })); }
    getCount() { return new Result(this.ctx, this._list().length); }
    _add(rec, o) {
      o = o || {};
      // Mac quirk: 0 (or missing) width/height becomes 72 pt.
      Object.assign(rec, { id: newId(), left: o.left ?? 0, top: o.top ?? 0, width: o.width || 72, height: o.height || 72, tags: {}, rotation: 0 });
      this._slide.shapes.push(rec);
      return new Shape(this.ctx, rec);
    }
    addGeometricShape(t, o) { return this._add({ type: "GeometricShape", geom: t }, o); }
    addLine(t, o) { return this._add({ type: "Line", geom: t }, o); }
    addTextBox(text, o) { return this._add({ type: "TextBox", text, tf: {} }, o); }
    addGroup(ids) {
      const recs = ids.map((x) => (typeof x === "string" ? x : x._rec.id)).map((id) => this._slide.shapes.find((r) => r.id === id));
      if (recs.some((r) => !r)) { this.ctx.errors.push(new Error("InvalidArgument: addGroup id not top-level")); return new Shape(this.ctx, null); }
      const l = Math.min(...recs.map((r) => r.left)), t = Math.min(...recs.map((r) => r.top));
      const rr = Math.max(...recs.map((r) => r.left + r.width)), b = Math.max(...recs.map((r) => r.top + r.height));
      const g = { id: newId(), type: "Group", left: l, top: t, width: rr - l, height: b - t, tags: {}, children: recs, rotation: 0 };
      recs.forEach((r) => { this._slide.shapes.splice(this._slide.shapes.indexOf(r), 1); r.parent = g; });
      this._slide.shapes.push(g);
      return new Shape(this.ctx, g);
    }
    getItem(id) { const f = findShape(id); return new Shape(this.ctx, f && f.shape, !f); }
    getItemOrNullObject(id) { const f = findShape(id); return new Shape(this.ctx, f ? f.shape : null); }
  }

  class SolidFill { constructor(ctx, color) { this.ctx = ctx; this._c = color; this._loaded = new Set(); } get isNullObject() { return !this._c; } get color() { return guard(this, "color", this._c); } load(p) { return this.ctx.queueLoad(this, p); } }
  class Slide {
    constructor(ctx, rec, missing) { this.ctx = ctx; this._rec = rec; this._loaded = new Set(); if (missing) ctx.errors.push(new Error("ItemNotFound slide")); }
    load(p) { return this.ctx.queueLoad(this, p); }
    get id() { return guard(this, "id", this._rec.id); }
    get shapes() { return new Shapes(this.ctx, () => this._rec.shapes, this._rec); }
    setSelectedShapes(ids) { window.__select(ids); }
    get themeColorScheme() { const ctx = this.ctx; return { getThemeColor: (k) => new Result(ctx, { Accent1: "#616C8B", Accent2: "#CD9700", Accent3: "#C00000", Accent4: "#003760", Accent5: "#6DAA2D", Accent6: "#7030A0" }[k]) }; }
    get background() { const ctx = this.ctx, rec = this._rec; const bg = { _loaded: new Set(["isMasterBackgroundFollowed"]), isMasterBackgroundFollowed: false, fill: { getSolidFillOrNullObject: () => new SolidFill(ctx, rec.bg) } }; bg.load = () => bg; return bg; }
    get layout() { const ctx = this.ctx; return { background: { fill: { getSolidFillOrNullObject: () => new SolidFill(ctx, null) } } }; }
    get slideMaster() { const ctx = this.ctx; return { background: { fill: { getSolidFillOrNullObject: () => new SolidFill(ctx, null) } } }; }
  }

  class Scoped {
    constructor(ctx, items) { this.ctx = ctx; this._items = items; this._loaded = new Set(); }
    load(p) { return this.ctx.queueLoad(this, "items"); }
    get items() { return guard(this, "items", this._items.map((x) => { x._loaded.add("*"); return x; })); }
  }

  class Presentation {
    constructor(ctx) { this.ctx = ctx; }
    getSelectedSlides() { return new Scoped(this.ctx, deck.selSlides.map((id) => new Slide(this.ctx, slideOf(id)))); }
    getSelectedShapes() { return new Scoped(this.ctx, deck.selShapes.map((id) => findShape(id)).filter(Boolean).map((f) => new Shape(this.ctx, f.shape))); }
    get slides() { const ctx = this.ctx; return { getItem: (id) => new Slide(ctx, slideOf(id), !slideOf(id)), getItemAt: (i) => new Slide(ctx, deck.slides[i]) }; }
    get pageSetup() { const ps = { _loaded: new Set(), get slideWidth() { return guard(ps, "slideWidth", deck.w); }, get slideHeight() { return guard(ps, "slideHeight", deck.h); } }; ps.load = (p) => this.ctx.queueLoad(ps, p); return ps; }
    setSelectedSlides(ids) { deck.selSlides = ids.slice(); }
  }

  window.PowerPoint = {
    run: async (fn) => { const ctx = new Ctx(); ctx.presentation = new Presentation(ctx); try { return await fn(ctx); } catch (e) { window.__errors.push(String(e && e.message || e)); throw e; } },
  };
  window.Office = {
    HostType: { PowerPoint: "PowerPoint" },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    context: {
      diagnostics: { host: "PowerPoint", platform: "Mock", version: "test" },
      requirements: { isSetSupported: () => true },
      document: { addHandlerAsync: (_t, h, cb) => { handlers.push(h); cb && cb({ status: "succeeded" }); } },
    },
    onReady: (cb) => setTimeout(() => cb({ host: "PowerPoint" }), 0),
  };
})();
