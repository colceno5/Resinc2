/* Presentación PPT (formato ValorA) — réplica del PPT de referencia "KPI_<inicio>_a_<fin>.pptx".
   Uso: PPTG.build(PptxGenJS, D, PPT_BG) -> instancia de PptxGenJS lista para .write()
   D lo arma pptData() en index.html. */
(function () {
  var NAVY = "002336", GR = "1B7F4B", RD = "C62828", OR = "E07026", TX = "404040", MU = "667085", BD = "D9DEE7", GRY = "A0A7B5", ALT = "F8FAFD";
  var n1 = function (v) { return v == null ? "—" : v.toFixed(1).replace(".", ","); };
  var tr = function (s, n) { s = String(s == null ? "" : s); return s.length > n ? s.slice(0, n) + "…" : s; };
  var plural = function (n, s, p) { return n === 1 ? s : p; };
  var DIAS = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], DC = ["Dom", "Lun", "Mar", "Mié", "Jue", "Vie", "Sáb"];

  function build(PptxGenJS, D, BG) {
    var pres = new PptxGenJS();
    pres.layout = "LAYOUT_WIDE";
    pres.title = D.titulo + " · " + D.rango;
    pres.author = "Rendimiento de atenciones · Aranda";
    pres.theme = { headFontFace: "Arial Black", bodyFontFace: "Arial" };
    var meta = D.meta, K = D.K, ok = function (p) { return p != null && p >= meta; };
    var pie = "Rango: " + D.rango + "  ·  Meta SLA " + meta + " %";

    pres.defineSlideMaster({ title: "PORTADA", background: { data: BG.portada },
      objects: [{ placeholder: { options: { name: "title", type: "title", x: 6.6, y: 2.35, w: 6.4, h: 2.1, fontFace: "Arial Black", fontSize: 30, color: "000000", align: "center", valign: "middle", margin: 0 }, text: "" } }] });
    pres.defineSlideMaster({ title: "CONTENIDO", background: { data: BG.contenido },
      objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.3, h: 0.7, fontFace: "Arial Black", fontSize: 24, color: TX, align: "left", valign: "middle", margin: 0 }, text: "" } }] });
    pres.defineSlideMaster({ title: "BLANCO", background: { color: "FFFFFF" } });
    pres.defineSlideMaster({ title: "CIERRE", background: { data: BG.cierre || BG.gracias } });

    function content(title, notes, withPie) {
      var s = pres.addSlide({ masterName: "CONTENIDO" });
      s.addText(title, { placeholder: "title" });
      if (withPie !== false) s.addText(pie, { x: 2.0, y: 6.93, w: 7.3, h: 0.3, fontFace: "Arial", fontSize: 9, color: MU, margin: 0, valign: "middle", isTextBox: true });
      if (notes) s.addNotes(notes);
      return s;
    }
    function box(s, x, y, w, h, line) { s.addShape(pres.ShapeType.rect, { x: x, y: y, w: w, h: h, fill: { color: "FFFFFF" }, line: line === false ? { type: "none" } : { color: BD, width: 0.75 } }); }
    function txt(s, t, x, y, w, h, o) { s.addText(t, Object.assign({ x: x, y: y, w: w, h: h, fontFace: "Arial", margin: 0, isTextBox: true }, o)); }
    function card(s, x, y, w, h, ac, val, vs, vc, lab, note) {
      box(s, x, y, w, h);
      s.addShape(pres.ShapeType.rect, { x: x, y: y, w: 0.07, h: h, fill: { color: ac }, line: { type: "none" } });
      var vh = h > 2 ? 1.06 : 0.7, ly = h > 2 ? 1.22 : 0.81, lh = h > 2 ? 0.47 : 0.31, ny = h > 2 ? 1.69 : 1.12, nh = h > 2 ? 0.52 : 0.34;
      txt(s, val, x + 0.2, y + 0.08, w - 0.3, vh, { fontFace: "Arial Black", fontSize: vs, color: vc, valign: "middle" });
      txt(s, lab, x + 0.2, y + ly, w - 0.3, lh, { fontSize: 11, bold: true, color: TX, valign: "middle" });
      txt(s, note, x + 0.2, y + ny, w - 0.3, nh, { fontSize: 9, color: MU, valign: "top" });
    }
    function lectura(s, x, y, w, h, titulo, items, nobox) {
      if (!nobox) box(s, x, y, w, h);
      s.addShape(pres.ShapeType.rect, { x: x, y: y, w: w, h: 0.07, fill: { color: OR }, line: { type: "none" } });
      txt(s, titulo, x + 0.2, y + 0.17, w - 0.4, 0.4, { fontFace: "Arial Black", fontSize: 13, color: NAVY, valign: "middle" });
      var runs = items.map(function (t, i) { return { text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6 } }; });
      txt(s, runs, x + 0.2, y + 0.65, w - 0.4, h - 0.8, { fontSize: 12, color: TX, valign: "top" });
    }
    function tabla(s, head, rows, x, y, colW, rowH, kpiCol, fs) {
      var b = { type: "solid", pt: 0.75, color: BD }, bb = [b, b, b, b];
      var H = head.map(function (h, i) { return { text: h, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: i ? "center" : "left", valign: "middle", border: bb } }; });
      var R = rows.map(function (r, ri) {
        return r.map(function (c, i) {
          var o = { color: TX, fill: { color: ri % 2 ? ALT : "FFFFFF" }, align: i ? "center" : "left", valign: "middle", border: bb };
          if (c && typeof c === "object") { Object.assign(o, c.o); c = c.t; }
          if (i === kpiCol && !o.keep) { o.bold = true; }
          return { text: String(c), options: o };
        });
      });
      s.addTable([H].concat(R), { x: x, y: y, w: colW.reduce(function (a, c) { return a + c; }, 0), colW: colW, rowH: rowH, fontFace: "Arial", fontSize: fs || 10, margin: [0, 0.08, 0, 0.08] });
    }
    function barras(s, cats, vals, o) {
      var cols = vals.map(function (v, i) { return o.colors ? o.colors[i] : ok(v * 100) ? GR : RD; });
      s.addChart(pres.charts.BAR, [{ name: "KPI", labels: cats, values: vals }], Object.assign({
        barGapWidthPct: 40, chartColors: cols, showLegend: false, showTitle: true, titleFontFace: "Arial", titleFontSize: 11, titleColor: TX,
        valAxisMinVal: 0, valAxisMaxVal: 1.15, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
        catAxisLabelFontFace: "Arial", catAxisLabelFontSize: 9, catAxisLabelColor: TX,
        showValue: true, dataLabelFormatCode: "0.0%", dataLabelPosition: "outEnd", dataLabelFontFace: "Arial", dataLabelFontSize: 9, dataLabelColor: TX, dataLabelFontBold: true
      }, o.opts));
    }
    function ranking(title, cat, rows, label, nota, colTxt, withLect) {
      var s = content(title, nota, true), N = rows.length, mx = rows.slice(0, label === "Especialista" ? 10 : 8), tot = rows.length;
      barras(s, mx.map(function (r) { return tr(r.k, 25); }), mx.map(function (r) { return (r.pct || 0) / 100; }),
        { opts: { x: 0.5, y: 1.25, w: 6.1, h: 5.0, barDir: "bar", catAxisOrientation: "maxMin", title: colTxt } });
      var body = mx.map(function (r) { return [tr(r.k, 29), r.n, r.a, r.f, { t: n1(r.pct) + " %", o: { color: ok(r.pct) ? GR : RD, bold: true } }]; });
      tabla(s, [label, "Evaluados", "Dentro", "Fuera", "KPI"], body, 6.85, 1.35, [2.5, 0.9, 0.8, 0.75, 1.05], 0.36, 4);
      var yEnd = 1.35 + (mx.length + 1) * 0.36;
      if (tot > mx.length) txt(s, "Se muestran " + mx.length + " de " + tot + " (ordenados por tickets evaluados). El detalle completo está en la pestaña Rendimiento del aplicativo.", 6.85, yEnd + 0.15, 6.0, 0.5, { fontSize: 9, color: MU, valign: "top" });
      return { s: s, yEnd: yEnd, list: mx };
    }

    /* 1. Portada */
    var s1 = pres.addSlide({ masterName: "PORTADA" });
    s1.addText(String(D.titulo).toUpperCase(), { placeholder: "title" });
    txt(s1, [{ text: "Informe ", options: { bold: true, fontSize: 20, color: OR } }, { text: "comité", options: { bold: true, fontSize: 20, color: "000000", breakLine: true } },
      { text: D.rango, options: { fontSize: 15, color: "595959" } }], 7.4, 4.85, 5.2, 1.2, { align: "center", valign: "middle" });
    s1.addNotes("Informe generado el " + D.fechaGen + " con los filtros del aplicativo." + (D.alcance ? " Alcance: " + D.alcance + "." : ""));

    /* 2. Resumen ejecutivo */
    var s2 = content("Resumen ejecutivo del KPI", "KPI justificado = tickets evaluados dentro del 100 % de progreso (incluye justificaciones aprobadas en la auditoría) ÷ tickets evaluados. Los cancelados y los tickets sin progreso no se evalúan.");
    var gap = K.pct - meta, kc = ok(K.pct) ? GR : RD, aj = K.a - K.o;
    box(s2, 0.5, 1.3, 4.9, 4.9);
    s2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.3, w: 4.9, h: 0.09, fill: { color: kc }, line: { type: "none" } });
    txt(s2, "KPI JUSTIFICADO", 0.8, 1.55, 4.3, 0.35, { fontSize: 12, bold: true, color: MU, valign: "middle" });
    txt(s2, n1(K.pct) + " %", 0.8, 1.95, 4.3, 1.3, { fontFace: "Arial Black", fontSize: 54, color: kc, valign: "middle" });
    txt(s2, gap >= 0 ? "Meta alcanzada" + (gap >= 0.05 ? " (+" + n1(gap) + " pts sobre la meta)" : "") : "Brecha de " + n1(-gap) + " pts frente a la meta", 0.8, 3.3, 4.3, 0.4, { fontSize: 14, bold: true, color: kc, valign: "middle" });
    s2.addShape(pres.ShapeType.rect, { x: 0.8, y: 4.35, w: 4.3, h: 0.26, fill: { color: "E5E8EE" }, line: { type: "none" } });
    s2.addShape(pres.ShapeType.rect, { x: 0.8, y: 4.35, w: Math.max(0.02, 4.3 * Math.min(100, K.pct) / 100), h: 0.26, fill: { color: kc }, line: { type: "none" } });
    s2.addShape(pres.ShapeType.rect, { x: 0.8 + 4.3 * meta / 100 - 0.015, y: 4.23, w: 0.03, h: 0.5, fill: { color: NAVY }, line: { type: "none" } });
    txt(s2, "Meta " + meta + " %", 3.88, 4.75, 1.1, 0.3, { fontSize: 10, bold: true, color: NAVY, align: "right", valign: "middle" });
    card(s2, 5.7, 1.3, 2.2, 2.35, NAVY, String(K.n), 30, NAVY, "Tickets evaluados", "sin cancelados ni sin progreso");
    card(s2, 8.05, 1.3, 2.2, 2.35, GR, String(K.a), 30, GR, "Dentro del 100 %", aj > 0 ? "incluye " + aj + " " + plural(aj, "justificado aprobado", "justificados aprobados") : "sin justificaciones aprobadas");
    card(s2, 10.4, 1.3, 2.2, 2.35, RD, String(K.f), 30, RD, "Fuera del 100 %", "tickets excedidos");
    card(s2, 5.7, 3.85, 2.2, 2.35, NAVY, n1(K.pO) + " %", 30, NAVY, "KPI original", "sin contar justificaciones");
    box(s2, 8.05, 3.85, 4.55, 2.35);
    s2.addShape(pres.ShapeType.rect, { x: 8.05, y: 3.85, w: 0.07, h: 2.35, fill: { color: NAVY }, line: { type: "none" } });
    txt(s2, meta + " %", 8.3, 4.1, 4.2, 1.1, { fontFace: "Arial Black", fontSize: 48, color: NAVY, align: "center", valign: "middle" });
    txt(s2, "Meta KPI", 8.3, 5.2, 4.2, 0.7, { fontSize: 28, bold: true, color: TX, align: "center", valign: "middle" });

    /* 3. Evolución */
    var s3 = content("Evolución del cumplimiento", "Un punto por " + D.granTxt + " según la fecha de registro del ticket. KPI justificado: incluye justificaciones aprobadas. KPI original: sin justificaciones. Meta " + meta + " %.");
    if (D.evoPng) s3.addImage({ data: String(D.evoPng).replace(/^data:/, ""), x: 0.4, y: 1.25, w: 8.9, h: 8.9 * 860 / 1500, altText: "Evolución del cumplimiento: KPI justificado, KPI original, tickets evaluados y meta" });
    else {
      var sr = D.serie;
      s3.addChart(pres.charts.LINE, [{ name: "KPI justificado", labels: sr.map(function (p) { return p.lab; }), values: sr.map(function (p) { return p.pct == null ? null : p.pct / 100; }) },
        { name: "KPI original", labels: sr.map(function (p) { return p.lab; }), values: sr.map(function (p) { return p.pO == null ? null : p.pO / 100; }) },
        { name: "Meta " + meta + " %", labels: sr.map(function (p) { return p.lab; }), values: sr.map(function () { return meta / 100; }) }],
      { x: 0.5, y: 1.25, w: 8.7, h: 5.0, chartColors: ["2F5FD0", GRY, OR], lineDataSymbol: "circle", valAxisLabelFormatCode: "0%", showLegend: true, legendPos: "t", valAxisMaxVal: 1, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10 });
    }
    var oks = D.serie.filter(function (p) { return ok(p.pct); }).length, best = D.serie.filter(function (p) { return p.pct != null; }).slice().sort(function (a, b) { return b.pct - a.pct; })[0], it3 = [];
    it3.push(oks + " de " + D.serie.length + " " + D.plTxt + " alcanzaron la meta.");
    if (best) it3.push("Mejor " + D.granTxt + ": " + best.labFull + " con " + n1(best.pct) + " %.");
    lectura(s3, 9.35, 1.3, 3.5, 4.9, "Lectura del periodo", it3, true);

    /* 4. Seguimiento diario */
    if (D.dias && D.dias.length) {
      var s4 = content("Seguimiento diario" + (D.semNum ? " semana " + D.semNum : ""), "Cumplimiento (KPI justificado) por día de registro del ticket y tickets evaluados por día. Meta " + meta + " %.");
      var ds = D.dias, labs = ds.map(function (d) { return d.lab; }), mn = Math.min.apply(null, ds.map(function (d) { return d.pct; }).concat([meta])), lo = Math.max(0, Math.floor((mn - 5) / 5) * 5) / 100, mxN = Math.max.apply(null, ds.map(function (d) { return d.n; }));
      s4.addChart([
        { type: pres.charts.BAR, data: [{ name: "Tickets evaluados", labels: labs, values: ds.map(function (d) { return d.n; }) }], options: { chartColors: ["DCE0F5"], barGapWidthPct: 50, secondaryValAxis: true, secondaryCatAxis: true } },
        { type: pres.charts.LINE, data: [{ name: "Cumplimiento", labels: labs, values: ds.map(function (d) { return d.pct / 100; }) }],
          options: { chartColors: ["3F5FC0"], lineSize: 2, lineDataSymbol: "none" } },
        { type: pres.charts.LINE, data: [{ name: "≥ meta", labels: labs, values: ds.map(function (d) { return ok(d.pct) ? d.pct / 100 : null; }) }],
          options: { chartColors: [GR], lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 10 } },
        { type: pres.charts.LINE, data: [{ name: "< meta", labels: labs, values: ds.map(function (d) { return ok(d.pct) ? null : d.pct / 100; }) }],
          options: { chartColors: [RD], lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 10 } },
        { type: pres.charts.LINE, data: [{ name: "Meta " + meta + " %", labels: labs, values: ds.map(function () { return meta / 100; }) }],
          options: { chartColors: [OR], lineSize: 1.5, lineDash: "dash", lineDataSymbol: "none" } }
      ], { x: 0.41, y: 1.13, w: 8.94, h: 5.1, showLegend: true, legendPos: "t", legendFontFace: "Arial", legendFontSize: 10, legendColor: TX,
        valAxes: [{ valAxisMinVal: lo, valAxisMaxVal: 1, valAxisMajorUnit: 0.05, valAxisLabelFormatCode: "0%", valAxisLabelFontSize: 10, valAxisLabelColor: MU, valGridLine: { color: "E3E6EB", size: 0.75 } },
          { valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: Math.max(4, mxN * 4), valGridLine: { style: "none" } }],
        catAxes: [{ catAxisLabelFontSize: 10, catAxisLabelColor: TX }, { catAxisHidden: true }] });
      var bajo = ds.filter(function (d) { return d.pct < 100; }), pleno = ds.filter(function (d) { return d.pct >= 100; }), it4 = [];
      it4.push("Del total de los " + ds.length + " " + plural(ds.length, "día evaluado", "días evaluados") + (bajo.length ? ", " + bajo.map(function (d) { return "el " + DIAS[d.dow] + " " + d.dd + " se tuvo un " + n1(d.pct) + " % dentro del 100 %"; }).join("; ") + "." : " se mantuvo el 100 %."));
      if (pleno.length && bajo.length) it4.push(pleno.map(function (d) { return DC[d.dow]; }).join(", ").replace(/, ([^,]*)$/, " y $1") + " se mantuvo el 100 %.");
      if (D.aud.aprob) it4.push(D.aud.aprob + " ticket(s) vencido(s) cuentan como cumplidos por justificación aprobada.");
      if (D.aud.rech) it4.push(D.aud.rech + " ticket(s) vencido(s) cuentan como no cumplidos por justificación desaprobada" + (D.aud.rech === 1 ? " y es el que quedó bajo el indicador." : " y quedaron bajo el indicador."));
      lectura(s4, 9.35, 1.3, 3.5, 4.9, "Lectura del periodo", it4);
    }

    /* 5-7. Grupo / especialista / categoría */
    var g = D.grupos, gs = g.filter(function (x) { return x.pct != null; }).slice().sort(function (a, b) { return b.pct - a.pct; });
    var r5 = ranking("Cumplimiento por grupo", "grupo", g, "Grupo", "Cumplimiento (KPI justificado) por grupo responsable, ordenado por tickets evaluados.", "KPI por grupo (verde = alcanza la meta)");
    var it5 = [gs.filter(function (x) { return ok(x.pct); }).length + " de " + gs.length + " alcanzan la meta de " + meta + " %."];
    if (gs.length > 1) { it5.push("Mayor cumplimiento: " + gs[0].k + " (" + n1(gs[0].pct) + " %)."); it5.push("Menor cumplimiento: " + gs[gs.length - 1].k + " (" + n1(gs[gs.length - 1].pct) + " %)."); }
    lectura(r5.s, 6.85, r5.yEnd + (g.length > r5.list.length ? 0.75 : 0.2), 6.0, 2.0, "Lectura", it5);
    ranking("Cumplimiento por especialista", "especialista", D.esp, "Especialista", "Cumplimiento (KPI justificado) por especialista, ordenado por tickets evaluados.", "KPI por especialista (verde = alcanza la meta)");
    if (D.cat && D.cat.length) ranking("Cumplimiento por categoría", "categoría", D.cat, "Categoría", "Cumplimiento (KPI justificado) por categoría del ticket, ordenado por tickets evaluados.", "KPI por categoría (verde = alcanza la meta)");

    /* 8. Auditoría */
    var A = D.aud, s8 = content("Auditoría de tickets vencidos", "Vencidos = tickets cerrados o resueltos con más del 100 % de progreso. Al aprobar una justificación completa (fecha real y sustento) el ticket cuenta como cumplido y el KPI se recalcula.");
    [["Vencidos", A.venc, NAVY, "cerrados fuera del 100 %"], ["Sin justificar", A.pend, NAVY, "sin fecha real o sin sustento"], ["Con sustento", A.sust, GR, "fecha real y sustento"], ["Aprobadas", A.aprob, GR, "cuentan como cumplidas"], ["Desaprobadas", A.rech, RD, "siguen fuera del KPI"]]
      .forEach(function (c, i) { card(s8, 0.5 + i * 2.5, 1.3, 2.3, 1.55, c[2], String(c[1]), 28, c[2], c[0], c[3]); });
    txt(s8, "Mayores desvíos sin aprobación", 0.5, 3.05, 7.3, 0.35, { fontFace: "Arial Black", fontSize: 12, color: NAVY, valign: "middle" });
    var tb = A.top.length ? A.top.map(function (t) { return [tr(t.c + " · " + (t.s || ""), 48), tr(t.r, 22), { t: Math.round(t.p) + " %", o: { color: RD, bold: true } }, t.est]; }) : [["Sin tickets vencidos pendientes de aprobación", "", "", ""]];
    tabla(s8, ["Ticket", "Especialista", "Progreso", "Estado"], tb, 0.5, 3.45, [3.0, 1.7, 0.9, 1.7], 0.38, -1, 9);
    barras(s8, ["KPI original", "KPI justificado", "Meta"], [K.pO / 100, K.pct / 100, meta / 100], { colors: [GRY, GR, OR], opts: { x: 8.1, y: 3.0, w: 4.75, h: 3.25, barDir: "col", barGapWidthPct: 50, title: "Efecto de las justificaciones" } });

    /* 9. Carga semanal, últimos N (fondo blanco, título y tabla grandes como en la referencia) */
    if (D.sem && D.sem.length) {
      var s9 = pres.addSlide({ masterName: "BLANCO" });
      s9.addNotes("Evaluados y KPI de la semana seleccionada y las semanas anteriores (por fecha de registro). Meta SLA " + meta + " %.");
      txt(s9, "Carga semanal, últimos " + D.sem.length, 1.0, 0.45, 11.3, 1.1, { fontSize: 48, color: "000000", valign: "middle" });
      var bb = { type: "solid", pt: 0.75, color: "E3E6EB" }, bd9 = [bb, bb, bb, bb], HN = "1E2B6B";
      var H9 = ["Semana", "Evaluados", "KPI orig.", "KPI just."].map(function (h, i) { return { text: h, options: { fontSize: 20, bold: true, color: "FFFFFF", fill: { color: HN }, align: i ? "center" : "left", valign: "middle", border: bd9 } }; });
      var R9 = D.sem.map(function (w, i) {
        var last = i === D.sem.length - 1, f = last ? "E6ECFA" : (i % 2 ? "F5F7FB" : "FFFFFF");
        return [{ text: w.lab + " · " + w.rg, options: { fontSize: 22, bold: last, color: "111111", fill: { color: f }, align: "left", valign: "middle", border: bd9 } },
          { text: String(w.n), options: { fontSize: 22, bold: last, color: "111111", fill: { color: f }, align: "center", valign: "middle", border: bd9 } },
          { text: n1(w.pO) + " %", options: { fontSize: 22, bold: last, color: "6B7280", fill: { color: f }, align: "center", valign: "middle", border: bd9 } },
          { text: n1(w.pct) + " %", options: { fontSize: 22, bold: true, color: ok(w.pct) ? GR : RD, fill: { color: f }, align: "center", valign: "middle", border: bd9 } }];
      });
      s9.addTable([H9].concat(R9), { x: 1.1, y: 2.0, w: 11.1, colW: [4.6, 2.2, 2.1, 2.2], rowH: 0.7, fontFace: "Calibri", fontSize: 20, margin: [0, 0.12, 0, 0.12] });
      var ym = 2.0 + (R9.length + 1) * 0.7 + 0.08;
      txt(s9, "Meta SLA", 1.2, ym, 4.4, 0.5, { fontFace: "Calibri", fontSize: 18, italic: true, color: "A8741A", valign: "middle" });
      txt(s9, meta + " %", 7.0, ym, 2.5, 0.5, { fontFace: "Calibri", fontSize: 18, bold: true, color: "A8741A", align: "center", valign: "middle" });
    }

    /* 10. Cierre */
    var s10 = pres.addSlide({ masterName: "CIERRE" });
    s10.addNotes("Fin de la presentación. Generada el " + D.fechaGen + " desde «Rendimiento de atenciones · Aranda».");
    return pres;
  }
  window.PPTG = { build: build };
  if (typeof module !== "undefined") module.exports = window.PPTG;
})();
