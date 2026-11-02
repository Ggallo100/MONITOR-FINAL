/* Sistema POA - DPDI · versión web. Componentes genéricos y páginas. Sin dependencias externas. */
(function () {
  "use strict";
  const POA = (window.POA = {});
  const CTX = window.POA_CTX || {};
  const MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"];
  const OK = new Set(["OK", "Cumple", "Dentro del techo", "Autofinanciada", "Superávit", "Alcanza la meta", "Aprobada"]);
  const BAD = new Set(["Revisar", "Excede techo", "Déficit", "Sin techo", "Bajo la meta", "Sin meta", "Observada"]);
  const $ = (sel, el = document) => el.querySelector(sel);
  const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const csrf = () => document.querySelector('meta[name="csrf-token"]').content;
  const nf2 = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const nf0 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });

  // ------------------------------------------------------------------ utilidades
  POA.fmt = function (v, fmt) {
    if (v === null || v === undefined || v === "") return "";
    const n = typeof v === "number" ? v : Number(v);
    switch (fmt) {
      case "money": return isNaN(n) ? esc(v) : nf2.format(Math.abs(n) < 0.005 ? 0 : n);
      case "pct": return isNaN(n) ? esc(v) : (n * 100).toFixed(1) + "%";
      case "int": return isNaN(n) ? esc(v) : nf0.format(n);
      case "num": return isNaN(n) ? esc(v) : (Number.isInteger(n) ? nf0.format(n) : n.toLocaleString("en-US", { maximumFractionDigits: 4 }));
      case "sino": return v === 1 || v === true || v === "1" ? "Sí" : "No";
      case "status": return v ? `<span class="chip ${OK.has(v) ? "ok" : BAD.has(v) ? "bad" : ""}">${esc(v)}</span>` : "";
      default: return esc(v);
    }
  };

  POA.toast = function (msg, tipo = "ok") {
    const t = document.createElement("div");
    t.className = "toast toast-" + tipo;
    t.textContent = msg;
    $("#toasts").appendChild(t);
    setTimeout(() => t.remove(), tipo === "error" ? 7000 : 3500);
  };

  POA.api = async function (method, url, body, opts = {}) {
    const init = { method, headers: { "X-CSRF-Token": csrf() }, credentials: "same-origin" };
    if (body instanceof FormData) init.body = body;
    else if (body !== undefined) { init.headers["Content-Type"] = "application/json"; init.body = JSON.stringify(body); }
    const r = await fetch(url, init);
    if (r.status === 401) { location.href = "/login?next=" + encodeURIComponent(location.pathname); throw new Error("Sesión expirada"); }
    let data = {};
    try { data = await r.json(); } catch (e) { data = { error: r.ok ? "" : "Error del servidor (" + r.status + ")" }; }
    if (!r.ok) { if (!opts.silencioso) POA.toast(data.error || "Error " + r.status, "error"); throw new Error(data.error || r.status); }
    return data;
  };

  POA.confirmar = (msg) => window.confirm(msg);

  POA.modal = function (titulo, html) {
    $("#modal-titulo").textContent = titulo;
    $("#modal-cuerpo").innerHTML = html;
    $("#modal").hidden = false;
    return $("#modal-cuerpo");
  };
  document.addEventListener("click", (e) => {
    if (e.target.matches("[data-cerrar]") || e.target.id === "modal") $("#modal").hidden = true;
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") $("#modal").hidden = true; });

  let listas = null;
  POA.listas = async function (forzar = false) {
    if (!listas || forzar) listas = await POA.api("GET", "/api/listas");
    return listas;
  };

  const debounce = (fn, ms = 250) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

  // ------------------------------------------------------- selector de unidad
  POA.buscarEnLista = function (titulo, items, onPick, describir) {
    const cuerpo = POA.modal(titulo, `<input class="buscador" placeholder="Escriba para filtrar…" autofocus>
      <div class="lista-busqueda"></div><div class="desc-busqueda nota"></div>`);
    const input = $("input", cuerpo), lista = $(".lista-busqueda", cuerpo), desc = $(".desc-busqueda", cuerpo);
    const pintar = () => {
      const w = input.value.toUpperCase().split(/\s+/).filter(Boolean);
      const vis = items.filter((it) => w.every((x) => it.label.toUpperCase().includes(x)));
      lista.innerHTML = vis.slice(0, 400).map((it) => `<button type="button" data-v="${esc(it.value)}">${esc(it.label)}</button>`).join("") ||
        '<p class="nota">Sin coincidencias.</p>';
    };
    lista.addEventListener("mouseover", (e) => { const b = e.target.closest("button"); if (b && describir) desc.textContent = describir(b.dataset.v); });
    lista.addEventListener("click", (e) => { const b = e.target.closest("button"); if (!b) return; $("#modal").hidden = true; onPick(b.dataset.v); });
    input.addEventListener("input", pintar);
    input.addEventListener("keydown", (e) => { if (e.key === "Enter") { const b = $("button", lista); if (b) b.click(); } });
    pintar();
    setTimeout(() => input.focus(), 30);
  };

  const btnUnidad = $("#btn-unidad");
  if (btnUnidad) btnUnidad.addEventListener("click", () => {
    POA.buscarEnLista("Cambiar unidad orgánica (código / centro de costo / unidad)",
      (CTX.unidades || []).map((u) => ({ value: u.id, label: u.label })), async (v) => {
        await POA.api("POST", "/api/contexto", { uid: Number(v) });
        location.reload();
      });
  });

  // ------------------------------------------------------------------- tabla
  class Grid {
    constructor(el, columns, { onSelect, rowClass, buscador } = {}) {
      this.el = el; this.columns = columns; this.onSelect = onSelect; this.rowClass = rowClass;
      this.rows = []; this.sel = null; this.filtro = "";
      if (buscador) buscador.addEventListener("input", () => { this.filtro = buscador.value.toUpperCase(); this.render(); });
      el.addEventListener("click", (e) => {
        const tr = e.target.closest("tr[data-id]");
        if (!tr) return;
        this.select(tr.dataset.id, true);
      });
    }
    setRows(rows) { this.rows = rows || []; this.render(); }
    visibles() {
      if (!this.filtro) return this.rows;
      const w = this.filtro.split(/\s+/).filter(Boolean);
      return this.rows.filter((r) => { const t = this.columns.map((c) => r[c.key] ?? "").join(" ").toUpperCase(); return w.every((x) => t.includes(x)); });
    }
    render() {
      const cab = this.columns.map((c) => `<th class="${["money", "pct", "int", "num"].includes(c.fmt) ? "num" : ""}">${esc(c.label)}</th>`).join("");
      const filas = this.visibles().map((r) => {
        const cls = [String(r.id) === String(this.sel) ? "sel" : "", this.rowClass ? this.rowClass(r) : ""].join(" ");
        return `<tr data-id="${esc(r.id)}" class="${cls}">` + this.columns.map((c) =>
          `<td class="${["money", "pct", "int", "num"].includes(c.fmt) ? "num" : ""}">${POA.fmt(r[c.key], c.fmt)}</td>`).join("") + "</tr>";
      }).join("");
      this.el.innerHTML = `<table class="tabla"><thead><tr>${cab}</tr></thead><tbody>${filas ||
        `<tr><td colspan="${this.columns.length}" class="vacio">Sin registros</td></tr>`}</tbody></table>`;
    }
    select(id, fire = false) {
      this.sel = id; this.render();
      const r = this.rows.find((x) => String(x.id) === String(id));
      if (fire && r && this.onSelect) this.onSelect(r);
      return r;
    }
    get(id) { return this.rows.find((x) => String(x.id) === String(id)); }
  }
  POA.Grid = Grid;

  // --------------------------------------------------------------- formulario
  class Form {
    constructor(el, fields, { onChange, kind } = {}) {
      this.el = el; this.fields = fields; this.onChange = onChange; this.kind = kind; this.values = {};
      this.el.classList.add("formulario", "grilla-form");
      this.el.innerHTML = fields.map((f) => this.html(f)).join("");
      this.el.addEventListener("input", (e) => this.cambio(e));
      this.el.addEventListener("change", (e) => this.cambio(e));
      this.el.addEventListener("click", (e) => this.clicMeses(e));
      this.el.addEventListener("change", (e) => {
        const cb = e.target.closest(".multi-lista input[type=checkbox]"); if (!cb) return;
        const caja = cb.closest(".multi"), sel = this.input(caja.dataset.multi);
        const op = Array.from(sel.options).find((o) => o.value === cb.value); if (op) op.selected = cb.checked;
        this.contarMulti(caja);
      });
      this.el.addEventListener("input", (e) => { if (e.target.matches(".multi-buscar")) this.pintarMulti(e.target.closest(".multi")); });
      this.el.addEventListener("click", (e) => {
        if (!e.target.matches("[data-multi-limpiar]")) return;
        const caja = e.target.closest(".multi");
        Array.from(this.input(caja.dataset.multi).options).forEach((o) => (o.selected = false));
        this.pintarMulti(caja);
      });
      this.refrescarOpciones();
    }
    pintarMulti(caja) {
      const sel = this.input(caja.dataset.multi), q = $(".multi-buscar", caja).value.trim().toLowerCase();
      const ops = Array.from(sel.options).filter((o) => !q || o.textContent.toLowerCase().includes(q))
        .sort((a, b) => (b.selected ? 1 : 0) - (a.selected ? 1 : 0));  // las elegidas primero
      $(".multi-lista", caja).innerHTML = ops.length ? ops.map((o) => `<label class="${o.selected ? "elegida" : ""}"><input type="checkbox" value="${esc(o.value)}" ${o.selected ? "checked" : ""}><span>${esc(o.textContent)}</span></label>`).join("")
        : '<p class="nota">Sin coincidencias.</p>';
      this.contarMulti(caja);
    }
    contarMulti(caja) {
      const n = this.input(caja.dataset.multi).selectedOptions.length;
      $(".multi-cuenta", caja).textContent = n === 1 ? "1 unidad seleccionada" : `${n} unidades seleccionadas`;
      caja.querySelectorAll(".multi-lista label").forEach((l) => l.classList.toggle("elegida", l.querySelector("input").checked));
    }
    html(f) {
      const id = "f_" + f.key + "_" + Math.random().toString(36).slice(2, 7);
      const cls = "campo" + (f.wide ? " ancho" : "") + (f.type === "months" || f.type === "marks" ? " meses" : "");
      const lab = `<label for="${id}">${esc(f.label)}${f.required ? " *" : ""}</label>`;
      switch (f.type) {
        case "textarea": return `<div class="${cls}">${lab}<textarea id="${id}" name="${f.key}" rows="3"></textarea></div>`;
        case "select":
          if (f.search || f.free) return `<div class="${cls}">${lab}<div class="combo"><input id="${id}" name="${f.key}" list="${id}_l" autocomplete="off"><datalist id="${id}_l"></datalist>${f.search ? `<button type="button" class="btn-mini" data-buscar="${f.key}" title="Buscar">🔍</button>` : ""}</div></div>`;
          return `<div class="${cls}">${lab}<select id="${id}" name="${f.key}"></select></div>`;
        case "multiselect": return `<div class="${cls}">${lab}<div class="multi" data-multi="${f.key}">
          <input type="search" class="multi-buscar" placeholder="Buscar por código, centro de costo o unidad…" autocomplete="off">
          <div class="multi-lista"></div>
          <div class="multi-pie"><span class="multi-cuenta"></span><button type="button" class="btn-mini" data-multi-limpiar>Quitar todas</button></div>
        </div><select id="${id}" name="${f.key}" multiple hidden></select></div>`;
        case "calc": return `<div class="${cls}"><span class="lbl">${esc(f.label)}</span><output data-calc="${f.key}"></output></div>`;
        case "check": return `<div class="${cls} campo-check"><label><input type="checkbox" name="${f.key}"> ${esc(f.label)}</label></div>`;
        case "password": return `<div class="${cls}">${lab}<input id="${id}" name="${f.key}" type="password" autocomplete="new-password"></div>`;
        case "color": return `<div class="${cls}">${lab}<div class="combo"><input type="color" data-color="${f.key}"><input id="${id}" name="${f.key}" placeholder="#RRGGBB"></div></div>`;
        case "months": case "marks": {
          const celdas = MESES.map((m, i) => `<div class="mes"><label><input type="checkbox" data-sel="${i}"> ${m}</label><input name="${f.key}__${i}" inputmode="decimal"><small data-mcalc="${i}"></small></div>`).join("");
          const btns = f.type === "months"
            ? `<button type="button" class="btn-mini" data-mes="dist">Distribuir equitativamente en meses seleccionados</button><button type="button" class="btn-mini" data-mes="prorr">Prorrateo anual 1/12</button>`
            : `<button type="button" class="btn-mini" data-mes="prorr">Prorrateo anual 1/12 (0.08333333)</button><button type="button" class="btn-mini" data-mes="x">Marcar «x» en seleccionados</button><button type="button" class="btn-mini" data-mes="factor">Factor 1/n en seleccionados</button>`;
          return `<div class="${cls}" data-meses="${f.key}" data-tipo="${f.type}"><span class="lbl">${esc(f.label)}</span><div class="grilla-meses">${celdas}</div><div class="botones-mini">${btns}<button type="button" class="btn-mini" data-mes="todos">Seleccionar todos</button><button type="button" class="btn-mini" data-mes="limpiar">Limpiar</button></div></div>`;
        }
        default: {
          const im = ["number", "money", "pct"].includes(f.type) ? ' inputmode="decimal"' : "";
          const ph = f.type === "pct" ? ' placeholder="ej. 10%"' : "";
          return `<div class="${cls}">${lab}<input id="${id}" name="${f.key}"${im}${ph}></div>`;
        }
      }
    }
    opciones(f) {
      if (f.options) return f.options;
      const L = listas || {};
      const v = this.get();
      switch (f.src) {
        case "objs": return (L.objs || {})[v.eje] || [];
        case "gg": return (L.gg || {})[v.tipo || "Egresos"] || [];
        case "sustento": return (L.sustento || {})[v.tipo || "Egresos"] || [];
        case "tipos_detalle": return ((L.tipos_detalle || {})[this.kind] || []).map((x) => ({ value: x, label: x }));
        default: return L[f.src] || [];
      }
    }
    refrescarOpciones(soloDe) {
      for (const f of this.fields) {
        if (!["select", "multiselect"].includes(f.type)) continue;
        if (soloDe && f.dep !== soloDe) continue;
        const ops = this.opciones(f);
        const el = this.input(f.key);
        if (!el) continue;
        if (f.search || f.free) {
          $("datalist", el.parentElement).innerHTML = ops.map((o) => `<option value="${esc(o.value)}">${esc(o.label !== o.value ? o.label : "")}</option>`).join("");
        } else if (f.type === "multiselect") {
          const actuales = new Set(Array.from(el.selectedOptions).map((o) => o.value));
          el.innerHTML = ops.map((o) => `<option value="${esc(o.value)}" ${actuales.has(String(o.value)) ? "selected" : ""}>${esc(o.label)}</option>`).join("");
          const caja = el.parentElement.querySelector(".multi"); if (caja) this.pintarMulti(caja);
        } else {
          const actual = el.value;
          el.innerHTML = '<option value=""></option>' + ops.map((o) => `<option value="${esc(o.value)}">${esc(o.label)}</option>`).join("");
          if (ops.some((o) => String(o.value) === actual)) el.value = actual;
        }
      }
    }
    input(key) { return this.el.querySelector(`[name="${key}"]`); }
    get() {
      const out = {};
      for (const f of this.fields) {
        if (f.type === "calc") continue;
        if (f.type === "months" || f.type === "marks") { out[f.key] = MESES.map((_, i) => this.input(`${f.key}__${i}`).value.trim()); continue; }
        const el = this.input(f.key);
        if (!el) continue;
        if (f.type === "check") out[f.key] = el.checked ? 1 : 0;
        else if (f.type === "multiselect") out[f.key] = Array.from(el.selectedOptions).map((o) => o.value);
        else out[f.key] = el.value.trim();
      }
      return out;
    }
    set(row) {
      this.limpiar(false);
      row = row || {};
      // primero los campos de los que dependen otros (eje, tipo)
      const orden = [...this.fields].sort((a, b) => (a.dep ? 1 : 0) - (b.dep ? 1 : 0));
      for (const f of orden) {
        if (f.type === "calc" || !(f.key in row)) continue;
        let v = row[f.key];
        if (f.type === "months" || f.type === "marks") {
          (v || []).forEach((x, i) => { this.input(`${f.key}__${i}`).value = POA.valorMes(x, f.type); });
          continue;
        }
        const el = this.input(f.key);
        if (!el) continue;
        if (f.type === "check") el.checked = v === 1 || v === true || v === "1";
        else if (f.type === "multiselect") {
          Array.from(el.options).forEach((o) => { o.selected = (v || []).map(String).includes(o.value); });
          const caja = el.parentElement.querySelector(".multi"); if (caja) { $(".multi-buscar", caja).value = ""; this.pintarMulti(caja); }
        }
        else if (f.type === "pct") el.value = v === null || v === undefined || v === "" ? "" : (Number(v) * 100).toFixed(2).replace(/\.?0+$/, "") + "%";
        else el.value = v ?? "";
        if (f.type === "color") { const c = this.el.querySelector(`[data-color="${f.key}"]`); if (c && /^#[0-9a-f]{6}$/i.test(v)) c.value = v; }
        if (this.fields.some((x) => x.dep === f.key)) this.refrescarOpciones(f.key);
      }
      this.values = this.get();
    }
    limpiar(fire = true) {
      this.el.querySelectorAll("input, textarea, select").forEach((el) => {
        if (el.type === "checkbox") el.checked = false;
        else if (el.tagName === "SELECT" && el.multiple) Array.from(el.options).forEach((o) => (o.selected = false));
        else if (el.type !== "color") el.value = "";
      });
      this.el.querySelectorAll("output, [data-mcalc]").forEach((o) => (o.innerHTML = ""));
      this.refrescarOpciones();
      if (fire && this.onChange) this.onChange(this.get(), null);
    }
    setCalc(obj) {
      for (const [k, v] of Object.entries(obj || {})) {
        const o = this.el.querySelector(`[data-calc="${k}"]`);
        if (o) { o.innerHTML = POA.fmt(v, OK.has(v) || BAD.has(v) ? "status" : "text"); }
      }
      if (obj && obj.meses_calc) obj.meses_calc.forEach((v, i) => { const m = this.el.querySelector(`[data-mcalc="${i}"]`); if (m) m.textContent = v; });
    }
    cambio(e) {
      const t = e.target;
      if (t.dataset.color) { const txt = this.input(t.dataset.color); if (txt) txt.value = t.value; }
      const nombre = t.name || "";
      if (this.fields.some((f) => f.dep === nombre)) this.refrescarOpciones(nombre);
      if (this.onChange) this.onChange(this.get(), nombre);
    }
    clicMeses(e) {
      const b = e.target.closest("[data-mes]");
      const bus = e.target.closest("[data-buscar]");
      if (bus) {
        const f = this.fields.find((x) => x.key === bus.dataset.buscar);
        const ops = this.opciones(f);
        const desc = f.src === "gg" ? (v) => ((listas.gg_desc || {})[this.get().tipo || "Egresos"] || {})[v] || "" : null;
        POA.buscarEnLista("Buscar " + f.label, ops, (v) => { this.input(f.key).value = v; this.cambio({ target: this.input(f.key) }); }, desc);
        return;
      }
      if (!b) return;
      const cont = b.closest("[data-meses]");
      const key = cont.dataset.meses, tipo = cont.dataset.tipo;
      const ins = MESES.map((_, i) => this.input(`${key}__${i}`));
      const sel = $$("[data-sel]", cont).filter((c) => c.checked).map((c) => Number(c.dataset.sel));
      const total = () => Number(String((this._total ?? 0)).replace(/,/g, "")) || 0;
      switch (b.dataset.mes) {
        case "todos": $$("[data-sel]", cont).forEach((c) => (c.checked = true)); return;
        case "limpiar": ins.forEach((i) => (i.value = "")); $$("[data-sel]", cont).forEach((c) => (c.checked = false)); break;
        case "prorr":
          if (tipo === "marks") ins.forEach((i) => (i.value = "0.08333333"));
          else POA.distribuir(total(), [...Array(12).keys()]).forEach((v, i) => (ins[i].value = v ?? ""));
          break;
        case "dist": case "factor": case "x":
          if (!sel.length) { POA.toast("Marque primero los meses (casillas).", "aviso"); return; }
          if (b.dataset.mes === "dist") POA.distribuir(total(), sel).forEach((v, i) => (ins[i].value = v ?? ""));
          else ins.forEach((inp, i) => (inp.value = sel.includes(i) ? (b.dataset.mes === "x" ? "x" : (1 / sel.length).toFixed(8)) : ""));
          break;
      }
      if (this.onChange) this.onChange(this.get(), key);
    }
    habilitar(si) {
      this.el.querySelectorAll("input, textarea, select, button").forEach((el) => (el.disabled = !si));
    }
  }
  POA.Form = Form;

  POA.valorMes = (x, tipo) => {
    if (x === null || x === undefined || x === "") return "";
    if (String(x).toLowerCase() === "x") return "x";
    const n = Number(x);
    if (isNaN(n)) return String(x);
    return tipo === "months" ? n.toFixed(2) : String(+n.toFixed(8));
  };

  POA.distribuir = function (total, meses) {
    const out = Array(12).fill(null);
    if (!meses.length) return out;
    const base = Math.round((total / meses.length) * 100) / 100;
    meses.forEach((m) => (out[m] = base.toFixed(2)));
    out[meses[meses.length - 1]] = (Math.round((total - base * (meses.length - 1)) * 100) / 100).toFixed(2);
    return out;
  };

  function techoHtml(t, etiqueta) {
    if (!t) return "";
    const uso = t.uso === null || t.uso === undefined ? null : t.uso;
    const w = Math.min(100, (uso || 0) * 100);
    return `<div class="techo"><div><strong>${esc(etiqueta)}</strong> · Ingresos S/ ${POA.fmt(t.ingresos, "money")} · Egresos S/ ${POA.fmt(t.egresos, "money")}
      ${t.techo ? ` de S/ ${POA.fmt(t.techo, "money")} (${POA.fmt(uso, "pct")}) · saldo S/ ${POA.fmt(t.saldo, "money")}` : ""}</div>
      ${t.techo ? `<div class="barra"><span style="width:${w}%" class="${uso > 1 ? "exceso" : ""}"></span></div>` : ""}
      ${POA.fmt(t.estado, "status")}</div>`;
  }

  function modoEdicion(editable) {
    $$("[data-edicion] button").forEach((b) => (b.disabled = !editable));
  }

  // ===================================================== página maestro-detalle
  POA.paginaMaestroDetalle = async function (cfg) {
    await POA.listas();
    const kind = cfg.kind, M = cfg.master;
    let mid = null, did = null, editable = CTX.editable, siguiente = "";
    const calcM = debounce(async (v) => {
      if (kind === "INI" || kind === "AP") mform.setCalc(await POA.api("POST", "/api/calc/" + kind, v, { silencioso: true }));
    });
    const mform = new Form($("#m-form"), M.fields, { kind, onChange: (v) => calcM(v) });
    const calcD = debounce(async (v) => {
      const r = await POA.api("POST", "/api/calc/detalle", v, { silencioso: true });
      dform._total = r.total; dform.setCalc(r);
    });
    const dform = new Form($("#d-form"), cfg.detalle_fields, { kind, onChange: (v) => calcD(v) });
    const mgrid = new Grid($("#m-tabla"), M.columns, {
      buscador: $("#m-buscar"), onSelect: (r) => cargarMaster(r),
      rowClass: (r) => (!(r.nombre || r.proyecto) ? "tenue" : (r._control_meta && r._control_meta !== "OK") || r._estado === "Déficit" ? "alerta-fila" : ""),
    });
    const dgrid = new Grid($("#d-tabla"), cfg.detalle_columns, {
      onSelect: (r) => { did = r.id; dform.set(r); calcD(dform.get()); $("#d-modo").textContent = "Editando"; },
      rowClass: (r) => (r._control === "Revisar" || (r._total > 0 && !r.sustento) ? "alerta-fila" : ""),
    });

    async function recargar(seleccionar) {
      const r = await POA.api("GET", `/api/md/${kind}`);
      editable = r.editable; siguiente = r.siguiente;
      modoEdicion(editable); mform.habilitar(editable); dform.habilitar(editable);
      $("#resumen-techo").innerHTML = techoHtml(r.techo, kind === "AP" ? "Actividades propias (autofinanciadas)" : "Consumo del techo");
      mgrid.setRows(r.rows);
      const objetivo = seleccionar ?? mid;
      if (objetivo && mgrid.get(objetivo)) cargarMaster(mgrid.select(objetivo));
      else if (!mid) nuevoMaster();
    }
    function nuevoMaster() {
      mid = null; mgrid.select(null); mform.limpiar(false);
      mform.set({ codigo: siguiente, ficha_sustento: kind === "INV" ? "No" : undefined, moneda: "S/" });
      $("#m-modo").textContent = "Nuevo"; calcM(mform.get());
      dgrid.setRows([]); $("#d-titulo").textContent = "Detalle — guarde o seleccione un registro maestro"; $("#d-info").textContent = "";
      nuevoDetalle();
    }
    async function cargarMaster(r) {
      if (!r) return;
      mid = r.id; mform.set(r); calcM(mform.get());
      $("#m-modo").textContent = "Editando " + r.codigo;
      await recargarDetalle();
    }
    async function recargarDetalle(selDid) {
      if (!mid) return;
      const r = await POA.api("GET", `/api/md/${kind}/${mid}/detalle`);
      const m = r.master;
      $("#d-titulo").textContent = `Detalle de ${m.codigo} · ${(m.nombre || m.proyecto || "").slice(0, 80)}`;
      let extra = "";
      if (kind === "AP") extra = ` · Punto de equilibrio: ${m._punto_equilibrio ?? "—"} participantes · Estado: ${m._estado || "—"}`;
      $("#d-info").innerHTML = `${r.rows.length} ítems · Ingresos S/ ${POA.fmt(m._ingresos, "money")} · Egresos S/ ${POA.fmt(m._egresos, "money")} · Resultado S/ ${POA.fmt(m._resultado, "money")}${esc(extra)}`;
      dgrid.setRows(r.rows);
      if (selDid && dgrid.get(selDid)) { dgrid.select(selDid); did = selDid; } else nuevoDetalle();
    }
    function nuevoDetalle() {
      did = null; dgrid.select(null); dform.limpiar(false);
      dform.set({ tipo: (listas.tipos_detalle[kind] || ["Egresos"])[0], moneda: "S/" });
      $("#d-modo").textContent = "Nuevo"; calcD(dform.get());
    }
    document.addEventListener("click", async (e) => {
      const a = e.target.closest("[data-accion]");
      if (!a) return;
      try {
        switch (a.dataset.accion) {
          case "m-nuevo": nuevoMaster(); break;
          case "m-guardar": {
            const v = mform.get();
            const r = await POA.api(mid ? "PUT" : "POST", mid ? `/api/md/${kind}/${mid}` : `/api/md/${kind}`, v);
            mid = r.id; POA.toast("Registro guardado."); await recargar(r.id); break;
          }
          case "m-duplicar": if (mid) { const r = await POA.api("POST", `/api/md/${kind}/${mid}/duplicar`); mid = r.id; await recargar(r.id); POA.toast("Registro duplicado con su detalle."); } break;
          case "m-ingreso": if (mid) { await POA.api("POST", `/api/md/AP/${mid}/ingreso-tarifa`); await recargar(mid); POA.toast("Ítem de ingreso generado."); } break;
          case "m-eliminar":
            if (mid && POA.confirmar("¿Eliminar el registro y todo su detalle?")) { await POA.api("DELETE", `/api/md/${kind}/${mid}`); mid = null; await recargar(); POA.toast("Registro eliminado."); }
            break;
          case "d-nuevo": nuevoDetalle(); break;
          case "d-guardar": case "d-duplicar": {
            if (!mid) { POA.toast("Guarde o seleccione primero el registro maestro.", "aviso"); return; }
            const v = dform.get();
            const dup = a.dataset.accion === "d-duplicar";
            if (dup && !did) return;
            const r = await POA.api(did && !dup ? "PUT" : "POST", did && !dup ? `/api/det/${kind}/${did}` : `/api/md/${kind}/${mid}/detalle`, v);
            POA.toast(dup ? "Ítem duplicado." : "Ítem guardado.");
            await recargar(mid); await recargarDetalle(r.id); break;
          }
          case "d-eliminar":
            if (did && POA.confirmar("¿Eliminar el ítem seleccionado?")) { await POA.api("DELETE", `/api/det/${kind}/${did}`); did = null; await recargar(mid); }
            break;
        }
      } catch (err) { /* el error ya se mostró */ }
    });
    await recargar();
  };

  // ====================================================== página funcionamiento
  POA.paginaFuncionamiento = async function (FLAT) {
    await POA.listas();
    let kind = "FE", rid = null, editable = CTX.editable, siguiente = "", form, grid;
    const calc = debounce(async (v) => form.setCalc(await POA.api("POST", "/api/calc/" + kind, v, { silencioso: true })));
    const renovar = (sel) => { const viejo = $(sel); const nuevo = viejo.cloneNode(false); viejo.replaceWith(nuevo); return nuevo; };
    function montar() {
      const C = FLAT[kind];
      $("#f-titulo").textContent = C.titulo;
      // contenedores nuevos: los eventos de la pestaña anterior no deben seguir activos
      form = new Form(renovar("#f-form"), C.fields, { kind, onChange: (v) => calc(v) });
      renovar("#f-tabla");
      const nuevo = $("#f-buscar").cloneNode(true); $("#f-buscar").replaceWith(nuevo); nuevo.value = "";
      grid = new Grid($("#f-tabla"), C.columns, {
        buscador: nuevo, onSelect: (r) => { rid = r.id; form.set(r); calc(form.get()); $("#f-modo").textContent = "Editando " + r.codigo; },
        rowClass: (r) => (r._control === "Revisar" || ((r._total || r._neto) > 0 && !r.sustento) ? "alerta-fila" : ""),
      });
    }
    async function recargar(sel) {
      const r = await POA.api("GET", `/api/fun/${kind}`);
      editable = r.editable; siguiente = r.siguiente;
      modoEdicion(editable); form.habilitar(editable);
      const i = r.info;
      $("#f-info").innerHTML = kind === "FE"
        ? techoHtml(i.techo, "Techo de funcionamiento") + `<div class="nota">Reajuste ${POA.fmt(i.reajuste, "pct")} desde ${esc(i.mes_reajuste)} en filas con «Aplica reajuste = Sí».</div>`
        : `<div class="techo"><strong>Ingreso neto anual devengado S/ ${POA.fmt(r.rows.reduce((a, x) => a + (x._neto || 0), 0), "money")}</strong> · Meta S/ ${POA.fmt(i.meta_ingresos, "money")} · Morosidad incobrable (provisión) ${POA.fmt(i.morosidad, "pct")} · <a href="/resumen">ver cobranza (2.1)</a></div>`;
      grid.setRows(r.rows);
      if (sel && grid.get(sel)) { grid.select(sel); rid = sel; } else nuevo();
    }
    function nuevo() {
      rid = null; grid.select(null); form.limpiar(false);
      form.set({ codigo: siguiente, moneda: "S/", aplica_morosidad: "No", aplica_reajuste: "No" });
      $("#f-modo").textContent = "Nuevo"; calc(form.get());
    }
    $$(".pestanas button").forEach((b) => b.addEventListener("click", async () => {
      $$(".pestanas button").forEach((x) => x.classList.toggle("activa", x === b));
      kind = b.dataset.tab; montar(); await recargar();
    }));
    document.addEventListener("click", async (e) => {
      const a = e.target.closest("[data-accion]");
      if (!a) return;
      try {
        switch (a.dataset.accion) {
          case "f-nuevo": nuevo(); break;
          case "f-guardar": case "f-duplicar": {
            const dup = a.dataset.accion === "f-duplicar";
            if (dup && !rid) return;
            const v = form.get(); if (dup) v.codigo = siguiente;
            const r = await POA.api(rid && !dup ? "PUT" : "POST", rid && !dup ? `/api/fun/${kind}/${rid}` : `/api/fun/${kind}`, v);
            POA.toast("Guardado."); await recargar(r.id); break;
          }
          case "f-eliminar": if (rid && POA.confirmar("¿Eliminar la fila seleccionada?")) { await POA.api("DELETE", `/api/fun/${kind}/${rid}`); await recargar(); } break;
          case "f-renumerar": await POA.api("POST", `/api/fun/${kind}/renumerar`); await recargar(); POA.toast("Códigos renumerados."); break;
        }
      } catch (err) { /* mostrado */ }
    });
    montar(); await recargar();
  };

  // ================================================================ página inicio
  POA.paginaInicio = async function () {
    const r = await POA.api("GET", "/api/inicio");
    const u = r.unidad, ed = r.editable, rol = CTX.rol;
    // estado de la ficha
    const est = u.estado_ficha || "Borrador";
    let acciones = "";
    if ((rol === "RESPONSABLE" || rol === "ADMIN") && ["Borrador", "Observada"].includes(est)) acciones += '<button class="btn btn-primario" data-est="enviar">Enviar a revisión de la DPDI</button>';
    if (rol === "ADMIN") {
      if (est === "Enviada") acciones += '<button class="btn btn-primario" data-est="aprobar">Aprobar</button>';
      if (["Enviada", "Aprobada"].includes(est)) acciones += '<button class="btn" data-est="observar">Observar</button>';
      if (est !== "Borrador") acciones += '<button class="btn" data-est="reabrir">Reabrir (borrador)</button>';
    }
    $("#estado").innerHTML = `<p><span class="chip grande estado-${est.toLowerCase()}">${esc(est)}</span></p>
      ${u.estado_fecha ? `<p class="nota">Último cambio: ${esc(u.estado_fecha)} por ${esc(u.estado_usuario)}</p>` : ""}
      ${u.estado_obs ? `<div class="alerta alerta-aviso"><strong>Observación DPDI:</strong> ${esc(u.estado_obs)}</div>` : ""}
      <p class="nota">Borrador u Observada: la unidad puede editar. Enviada o Aprobada: solo lectura para la unidad.</p>
      <div class="botones">${acciones}</div>`;
    $("#estado").addEventListener("click", async (e) => {
      const b = e.target.closest("[data-est]");
      if (!b) return;
      let obs = "";
      if (b.dataset.est === "observar") { obs = prompt("Observación para la unidad:") || ""; if (!obs) return; }
      if (b.dataset.est === "enviar" && !POA.confirmar("Al enviar, la ficha queda en solo lectura hasta que la DPDI la observe o la reabra. ¿Continuar?")) return;
      try { await POA.api("POST", "/api/estado", { accion: b.dataset.est, observacion: obs }); location.reload(); } catch (err) { /* */ }
    });
    // datos de la unidad
    const campos = [["responsable", "Responsable de la unidad"], ["cargo", "Cargo"], ["correo", "Correo electrónico"],
      ["fecha_presentacion", "Fecha de presentación"], ["version_ficha", "Versión de la ficha"], ["elaborado_nombre", "Elaborado por (nombre)"],
      ["elaborado_cargo", "Elaborado por (cargo)"], ["elaborado_fecha", "Fecha (elaborado)"], ["revisado_nombre", "Revisado por DPDI (nombre)"],
      ["revisado_cargo", "Revisado por (cargo)"], ["revisado_fecha", "Fecha (revisado)"]];
    const fdat = new Form($("#datos-unidad"), [
      { key: "_centro", label: "Centro de costo", type: "calc" }, { key: "_nombre", label: "Unidad orgánica", type: "calc" },
      ...campos.map(([k, l]) => ({ key: k, label: l, type: "text" }))]);
    fdat.set(u); fdat.setCalc({ _centro: u.centro_costo, _nombre: u.nombre_unidad }); fdat.habilitar(ed);
    const bd = document.createElement("button"); bd.className = "btn btn-primario"; bd.textContent = "Guardar datos de la unidad"; bd.disabled = !ed;
    bd.onclick = async () => { await POA.api("PUT", "/api/unidad/datos", fdat.get()); POA.toast("Datos guardados."); };
    $("#datos-unidad").after(bd);
    // códigos
    const tipos = [["INI", "Iniciativas (INI-NN)"], ["FI", "Funcionamiento ingresos (FI-NNN)"], ["FE", "Funcionamiento egresos (FE-NNN)"],
      ["OO", "Otras operativas (OO-NNN)"], ["AP", "Actividades propias (AP-NNN)"], ["INV", "Inversiones (INV-NNN)"]];
    $("#codigos").innerHTML = `<table class="tabla compacta"><tbody>${tipos.map(([k, l]) =>
      `<tr><td>${l}</td><td class="nota">${r.codigos[k].registrados} reg. · sig. ${esc(r.codigos[k].siguiente)}</td><td><input class="num-corto" data-cod="${k}" value="0" inputmode="numeric" ${ed ? "" : "disabled"}></td></tr>`).join("")}</tbody></table>
      <button class="btn btn-primario" id="btn-codigos" ${ed ? "" : "disabled"}>Generar códigos</button><pre id="codigos-res" class="nota"></pre>`;
    $("#btn-codigos").onclick = async () => {
      const q = {}; $$("[data-cod]").forEach((i) => (q[i.dataset.cod] = Number(i.value) || 0));
      const res = await POA.api("POST", "/api/codigos", q);
      $("#codigos-res").textContent = res.creados.map((c) => `${c.tipo.padEnd(4)} ${c.codigo.padEnd(9)} → ${c.codigo_bd}`).join("\n");
      POA.toast(res.creados.length + " códigos generados.");
    };
    // plantilla
    new Grid($("#plantilla"), [{ key: "tipo_flujo", label: "Tipo" }, { key: "partida", label: "Partida / rubro" },
      { key: "base_calculo", label: "Base" }, { key: "_monto", label: "Monto anual (S/)", fmt: "money" }]).setRows(r.plantilla.map((x, i) => ({ id: i, ...x })));
    $("#btn-correr").disabled = !ed;
    $("#btn-correr").onclick = async () => {
      const res = await POA.api("POST", "/api/funcionamiento/correr", { reemplazar: $("#reemplazar").checked });
      $("#correr-info").textContent = `✔ ${res.FE} egresos (S/ ${POA.fmt(res.monto_FE, "money")}) y ${res.FI} ingresos (S/ ${POA.fmt(res.monto_FI, "money")}) prorrateados 1/12. Filas reemplazadas: ${res.eliminadas}.`;
      POA.toast("Plantilla de funcionamiento ejecutada.");
    };
    // techos
    const nombres = { INI: "Iniciativas con impacto estratégico", FE: "Actividades de funcionamiento", OO: "Otras actividades operativas", INV: "Actividades de inversión", AP: "Actividades propias (autofinanciadas)" };
    $("#techos").innerHTML = Object.entries(nombres).map(([k, l]) => techoHtml(r.techos[k], l)).join("");
    // importar ficha
    $("#form-importar").addEventListener("submit", async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      try {
        const rep = await POA.api("POST", "/cargas/ficha", fd);
        POA.modal("Importación completada", `<p>${esc(rep.archivo)}${rep.creada ? " (unidad creada)" : ""}</p><pre>${esc(Object.entries(rep.conteos).map(([k, v]) => k + ": " + v).join("\n"))}</pre>
          ${rep.avisos.length ? "<h4>Avisos</h4><ul>" + rep.avisos.slice(0, 20).map((a) => `<li>${esc(a)}</li>`).join("") + "</ul>" : ""}<button class="btn btn-primario" onclick="location.reload()">Aceptar</button>`);
      } catch (err) { /* */ }
    });
  };

  // ================================================================ página resumen
  POA.paginaResumen = async function () {
    const r = await POA.api("GET", "/api/resumen");
    const nRev = r.directrices.filter((d) => d.resultado === "Revisar").length;
    $("#r-cab").innerHTML = `${esc(r.unidad.display)} · Año ${esc(r.anio)} · Directrices: <strong>${11 - nRev}/11 cumplen</strong> · Resultado S/ ${POA.fmt(r.resultado, "money")}`;
    $("#r-semaforo").innerHTML = `<section class="tarjeta"><h2>11. Verificación automática de directrices DPDI</h2><div class="semaforo">${r.directrices.map((d, i) =>
      `<div class="sem ${d.resultado === "Cumple" ? "ok" : "bad"}"><span class="luz"></span><strong>${esc(d.codigo)}</strong><span>${esc(d.descripcion)}</span><span>${d.n} obs.</span>${d.items.length ? `<button class="btn-mini" data-dir="${i}">Ver detalle</button>` : ""}</div>`).join("")}</div></section>`;
    $("#r-semaforo").addEventListener("click", (e) => {
      const b = e.target.closest("[data-dir]"); if (!b) return;
      const d = r.directrices[Number(b.dataset.dir)];
      POA.modal(`${d.codigo} — ${d.descripcion}`, `<ul>${d.items.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>`);
    });
    $("#r-secciones").innerHTML = r.secciones.filter((s) => s.key !== "s11").map((sec) => {
      const pct = new Set(sec.pct_rows || []), tot = new Set(sec.totals || []);
      const filas = sec.rows.map((row, i) => `<tr class="${tot.has(i) ? "total" : ""}">${row.map((v, j) => {
        let f = sec.formats[j] || "text";
        if (f === "money_or_pct") f = String(row[0]).includes("(%)") ? "pct" : "money";
        if (pct.has(i) && j > 0) f = "pct";
        const num = typeof v === "number";
        return `<td class="${num ? "num" : ""}">${typeof v === "string" && (OK.has(v) || BAD.has(v)) ? POA.fmt(v, "status") : num ? POA.fmt(v, f === "text" ? "money" : f) : esc(v ?? "")}</td>`;
      }).join("")}</tr>`).join("");
      return `<section class="tarjeta"><h2>${esc(sec.title)}</h2><div class="tabla-scroll"><table class="tabla"><thead><tr>${sec.headers.map((h) => `<th>${esc(h)}</th>`).join("")}</tr></thead><tbody>${filas}</tbody></table></div>${sec.note ? `<p class="nota">${esc(sec.note)}</p>` : ""}</section>`;
    }).join("");
  };

  // ============================================================ administración
  POA.paginaAdmin = async function ({ entidad, cfg }) {
    await POA.listas();
    let rid = null;
    const filtro = $("#filtro");
    const url = () => `/api/admin/t/${entidad}` + (filtro ? "?tipo=" + filtro.value : "");
    const urlId = (id) => `/api/admin/t/${entidad}/${id}` + (filtro ? "?tipo=" + filtro.value : "");
    const form = new Form($("#a-form"), cfg.fields, {});
    const grid = new Grid($("#a-tabla"), cfg.columns, {
      buscador: $("#a-buscar"), onSelect: (r) => { rid = r.id; form.set(r); $("#a-modo").textContent = "Editando"; },
    });
    async function recargar(sel) {
      const r = await POA.api("GET", url());
      grid.setRows(r.rows); $("#a-cuenta").textContent = r.rows.length + " registros";
      if (sel && grid.get(sel)) { rid = sel; form.set(grid.select(sel)); } else nuevo();
    }
    function nuevo() {
      rid = null; grid.select(null); form.limpiar(false);
      const def = { unidades: { macrogrupo: "Soporte Académico/Administrativo" }, cuentas: { tipo_flujo: "Egresos", activo: 1 },
        plantilla: { ambito: "GLOBAL", macrogrupo: "*", tipo_flujo: "Egresos", base_calculo: "MONTO_FIJO", porcentaje: "1", aplica_ajuste: "No", activo: 1 },
        usuarios: { rol: "RESPONSABLE", activo: 1, debe_cambiar_clave: 1 } }[entidad];
      if (def) form.set(def);
      $("#a-modo").textContent = "Nuevo";
    }
    if (filtro) filtro.addEventListener("change", () => recargar());
    document.addEventListener("click", async (e) => {
      const a = e.target.closest("[data-accion]"); if (!a) return;
      try {
        if (a.dataset.accion === "a-nuevo") nuevo();
        if (a.dataset.accion === "a-guardar") {
          const r = await POA.api(rid ? "PUT" : "POST", rid ? urlId(rid) : url(), form.get());
          POA.toast("Guardado."); await POA.listas(true); form.refrescarOpciones(); await recargar(r.id);
          if (r.aviso) POA.toast(r.aviso, "aviso");
          if (r.clave_temporal) POA.modal("Usuario creado", `<p>Entregue estos datos a la persona por un canal seguro. Deberá cambiar la clave en su primer ingreso.</p>
            <table class="tabla"><tr><th>Usuario</th><td><strong>${esc(r.usuario)}</strong></td></tr><tr><th>Clave temporal</th><td><code class="clave">${esc(r.clave_temporal)}</code></td></tr></table>
            <p class="nota">Esta clave no se volverá a mostrar.</p><button class="btn btn-primario" data-cerrar>Aceptar</button>`);
        }
        if (a.dataset.accion === "a-eliminar" && rid && POA.confirmar("¿Eliminar el registro seleccionado?")) {
          await POA.api("DELETE", urlId(rid)); POA.toast("Eliminado."); await recargar();
        }
      } catch (err) { /* */ }
    });
    await recargar();
  };

  POA.paginaParametros = async function ({ generales, extra }) {
    const p = await POA.api("GET", "/api/admin/parametros");
    const pct = new Set(["inflacion_proyectada", "reajuste_remunerativo_pct", "tasa_morosidad_pct"]);
    const fp = new Form($("#p-form"), generales.map(([k, l, n]) => ({ key: k, label: l + (n ? " — " + n : ""), type: pct.has(k) ? "pct" : "text" })));
    fp.set(p);
    $("#p-guardar").onclick = async () => { await POA.api("PUT", "/api/admin/parametros", fp.get()); POA.toast("Parámetros guardados."); };
    const ft = new Form($("#t-form"), extra);
    ft.set(p);
    $("#t-guardar").onclick = async () => { await POA.api("PUT", "/api/admin/parametros", ft.get()); POA.toast("Apariencia guardada."); setTimeout(() => location.reload(), 600); };
  };

  POA.paginaBitacora = async function () {
    const grid = new Grid($("#b-tabla"), [{ key: "fecha", label: "Fecha" }, { key: "usuario", label: "Usuario" }, { key: "accion", label: "Acción" },
      { key: "detalle", label: "Detalle" }, { key: "ip", label: "IP" }]);
    const cargar = async () => { const r = await POA.api("GET", "/api/admin/bitacora?q=" + encodeURIComponent($("#b-buscar").value)); grid.setRows(r.rows.map((x, i) => ({ id: i, ...x }))); };
    $("#b-buscar").addEventListener("input", debounce(cargar, 400));
    await cargar();
  };

  POA.paginaTablero = async function () {
    const r = await POA.api("GET", "/api/admin/tablero");
    const cuenta = (e) => r.rows.filter((x) => x.estado === e).length;
    const excede = r.rows.filter((x) => x.uso > 1).length;
    $("#t-kpis").innerHTML = [["Unidades", r.rows.length], ["Con datos", r.rows.filter((x) => x.con_datos).length], ["Enviadas", cuenta("Enviada")],
      ["Aprobadas", cuenta("Aprobada")], ["Observadas", cuenta("Observada")], ["Exceden techo", excede]]
      .map(([l, v]) => `<div class="kpi"><span>${l}</span><strong>${v}</strong></div>`).join("");
    const grid = new Grid($("#t-tabla"), [{ key: "unidad", label: "Unidad (código / centro de costo / unidad)" }, { key: "macrogrupo", label: "Macrogrupo" },
      { key: "responsables", label: "Responsables" }, { key: "estado", label: "Estado", fmt: "status" }, { key: "estado_fecha", label: "Fecha estado" },
      { key: "ingresos", label: "Ingresos OPEX", fmt: "money" }, { key: "egresos", label: "Egresos OPEX*", fmt: "money" }, { key: "techo", label: "Techo OPEX", fmt: "money" },
      { key: "uso", label: "% uso", fmt: "pct" }, { key: "capex", label: "CAPEX", fmt: "money" }, { key: "revisar", label: "Directrices a revisar", fmt: "int" }], {
      buscador: $("#t-buscar"), rowClass: (x) => (x.uso > 1 || x.revisar ? "alerta-fila" : ""),
      onSelect: async (x) => { if (POA.confirmar(`¿Ir a la ficha de ${x.unidad}?`)) { await POA.api("POST", "/api/contexto", { uid: x.id }); location.href = "/resumen"; } },
    });
    grid.setRows(r.rows);
    $("#t-tabla").insertAdjacentHTML("afterend", '<p class="nota">* Egresos OPEX sin actividades propias (estas deben autofinanciarse). Haga clic en una fila para abrir su ficha.</p>');
  };

  // cargas de archivos genéricas (data-carga)
  document.addEventListener("submit", async (e) => {
    const f = e.target.closest("form[data-carga]"); if (!f) return;
    e.preventDefault();
    const boton = f.querySelector("button"), texto = boton ? boton.textContent : "";
    if (boton) { boton.disabled = true; boton.textContent = "Procesando…"; }
    try {
      const rep = await POA.api("POST", f.dataset.carga, new FormData(f));
      const avisos = rep.avisos || [];
      const txt = Object.entries(rep).filter(([k]) => !["avisos", "credenciales"].includes(k)).map(([k, v]) => `${k}: ${typeof v === "object" ? JSON.stringify(v) : v}`).join("\n");
      const cred = rep.credenciales || [];
      const tablaCred = cred.length ? `<h4>Claves temporales (${cred.length})</h4><p class="nota">Descárguelas ahora: no se volverán a mostrar.</p>
        <div class="tabla-scroll alto-bajo"><table class="tabla"><tr><th>Usuario</th><th>Nombre</th><th>Clave temporal</th></tr>${cred.map((c) => `<tr><td>${esc(c.usuario)}</td><td>${esc(c.nombre)}</td><td><code>${esc(c.clave)}</code></td></tr>`).join("")}</table></div>
        <button class="btn" id="btn-cred">⬇ Descargar claves (.csv)</button> ` : "";
      POA.modal("Carga completada", `<pre>${esc(txt)}</pre>${tablaCred}${avisos.length ? "<h4>Avisos</h4><ul>" + avisos.slice(0, 50).map((a) => `<li>${esc(a)}</li>`).join("") + "</ul>" : ""}<button class="btn btn-primario" onclick="location.reload()">Aceptar</button>`);
      if (cred.length) $("#btn-cred").onclick = () => {
        const q = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
        const csv = "\ufeffUsuario;Nombre;Correo;Clave temporal\r\n" + cred.map((c) => [c.usuario, c.nombre, c.correo, c.clave].map(q).join(";")).join("\r\n");
        const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
        a.download = "Claves_temporales_POA_DPDI.csv"; a.click(); URL.revokeObjectURL(a.href);
      };
    } catch (err) { /* el mensaje de error ya se mostró */ } finally {
      if (boton) { boton.disabled = false; boton.textContent = texto; }
    }
  });
})();
