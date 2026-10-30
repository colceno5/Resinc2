/* Presentación para gerencia · KPI de Resolución de Incidencias
   Genera un .pptx con el formato de la plantilla de ValorA (INDUCCIÓN_TI.pptx).
   Uso:  PPTG.build(PptxGenJS, datos, PPT_BG)  →  instancia de PptxGenJS lista para write()/writeFile().
   Los fondos (portada, contenido, gracias) vienen de ppt-plantilla.js. */
(function (root) {
  "use strict";
  const C = { navy: "002336", ora: "E07026", red: "D22C29", gr: "1B7F4B", rd: "C62828", am: "B26A00", tx: "404040", mu: "667085", bd: "D9DEE7", soft: "F3F6FA", sf: "FFFFFF", gy: "A0A7B5", ac: "2F5FD0" };
  const FH = "Arial Black", FB = "Arial";
  const nf = v => v == null ? "—" : v.toFixed(1).replace(".", ",") + " %";
  const sg = v => v == null || Math.abs(v) < .05 ? "±0" : (v > 0 ? "+" : "−") + Math.abs(v).toFixed(1).replace(".", ",");
  const cut = (s, n) => { s = String(s == null ? "" : s).replace(/\s+/g, " ").trim(); return s.length > n ? s.slice(0, n - 1).trimEnd() + "…" : s };
  const d64 = s => String(s || "").replace(/^data:/, "");

  function build(PptxGenJS, D, BG) {
    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_WIDE"; // 13,33" × 7,5" como la plantilla
    pres.title = D.titulo + " · " + D.rango;
    pres.subject = "Informe de KPI para gerencia";
    pres.author = "Mesa de ayuda";
    pres.theme = { headFontFace: FH, bodyFontFace: FB };
    const meta = D.meta, kc = v => v == null ? C.mu : v >= meta ? C.gr : C.rd;
    const gl = D.granTxt; // "semana" | "día" | "mes"

    /* ---------- layouts (marcos de la plantilla) ---------- */
    pres.defineSlideMaster({
      title: "PORTADA", background: { data: d64(BG.portada) },
      objects: [
        { placeholder: { options: { name: "title", type: "title", x: 6.6, y: 2.35, w: 6.4, h: 2.1, fontFace: FH, fontSize: 30, color: "000000", align: "center", valign: "middle", margin: 0 }, text: "" } },
        { placeholder: { options: { name: "sub", type: "body", x: 7.4, y: 4.85, w: 5.2, h: 1.2, fontFace: FB, fontSize: 16, color: C.ora, align: "center", valign: "middle", margin: 0 }, text: "" } }
      ]
    });
    pres.defineSlideMaster({
      title: "CONTENIDO", background: { data: d64(BG.contenido) },
      objects: [
        { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 12.3, h: 0.7, fontFace: FH, fontSize: 24, color: C.tx, align: "left", valign: "middle", margin: 0 }, text: "" } }
      ]
    });
    pres.defineSlideMaster({ title: "GRACIAS", background: { data: d64(BG.gracias) }, objects: [] });

    const foot = "Rango: " + D.rango + "  ·  " + D.alcance + "  ·  Meta SLA " + meta + " %";
    function content(title, notes, section) {
      const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: section });
      s.addText(title, { placeholder: "title" });
      s.addText(foot, { x: 2.0, y: 6.93, w: 7.3, h: 0.3, fontFace: FB, fontSize: 9, color: C.mu, margin: 0, valign: "middle", isTextBox: true, objectName: "Pie" });
      if (notes) s.addNotes(notes);
      return s;
    }
    // tarjeta: valor grande + etiqueta + nota
    function card(s, x, y, w, h, val, label, hint, color, vs) {
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.sf }, line: { color: C.bd, width: 0.75 }, objectName: "Tarjeta " + label });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color: color || C.navy }, line: { type: "none" }, objectName: "Acento " + label });
      s.addText(String(val), { x: x + 0.2, y: y + 0.08, w: w - 0.3, h: h * 0.45, fontFace: FH, fontSize: vs || 26, color: color || C.navy, valign: "middle", margin: 0, fit: "shrink", isTextBox: true, objectName: "Valor " + label });
      s.addText(label, { x: x + 0.2, y: y + h * 0.52, w: w - 0.3, h: h * 0.2, fontFace: FB, fontSize: 11, bold: true, color: C.tx, valign: "middle", margin: 0, isTextBox: true, objectName: "Etiqueta " + label });
      if (hint) s.addText(hint, { x: x + 0.2, y: y + h * 0.72, w: w - 0.3, h: h * 0.22, fontFace: FB, fontSize: 9, color: C.mu, valign: "top", margin: 0, isTextBox: true, objectName: "Nota " + label });
    }
    const axis = (mn) => ({ valAxisMinVal: mn, valAxisMaxVal: 1, valAxisMajorUnit: .1, valAxisLabelFormatCode: "0%" });
    const baseChart = { catAxisLabelFontFace: FB, valAxisLabelFontFace: FB, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10, catAxisLabelColor: C.mu, valAxisLabelColor: C.mu, valGridLine: { color: "E3E6EB", size: 0.75 }, catGridLine: { style: "none" }, legendFontFace: FB, legendFontSize: 11, legendColor: C.tx };
    const panel = (s, x, y, w, h, title, items, fs) => {
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.sf }, line: { color: C.bd, width: 0.75 }, objectName: "Panel " + title });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.07, fill: { color: C.ora }, line: { type: "none" }, objectName: "Franja " + title });
      s.addText(title, { x: x + 0.2, y: y + 0.17, w: w - 0.4, h: 0.4, fontFace: FH, fontSize: 13, color: C.navy, margin: 0, valign: "middle", isTextBox: true, objectName: "Título " + title });
      const runs = []; items.forEach((t, i) => runs.push({ text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: fs ? 12 : 7 } }));
      s.addText(runs.length ? runs : [{ text: "Sin información para mostrar.", options: {} }], { x: x + 0.2, y: y + 0.65, w: w - 0.4, h: h - 0.8, fontFace: FB, fontSize: fs || 12, color: C.tx, valign: "top", margin: 0, fit: "shrink", isTextBox: true, objectName: "Texto " + title });
    };

    /* ---------- 1. Portada ---------- */
    pres.addSection({ title: "Portada" });
    let s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Portada" });
    s.addText(D.titulo.toUpperCase(), { placeholder: "title" });
    s.addText([{ text: "Informe para gerencia", options: { bold: true, fontSize: 20, color: C.ora, breakLine: true } }, { text: D.rango, options: { fontSize: 15, color: "595959" } }], { placeholder: "sub" });
    s.addNotes("Informe generado el " + D.fechaGen + " con los filtros del aplicativo. " + D.alcance + ".");

    /* ---------- 2. Resumen ejecutivo ---------- */
    pres.addSection({ title: "Resumen" });
    const K = D.K, gap = K.pct == null ? null : K.pct - meta;
    s = content("Resumen ejecutivo del KPI", "KPI justificado = tickets evaluados dentro del 100 % de progreso (incluye justificaciones aprobadas en la auditoría) ÷ tickets evaluados. Los cancelados y los tickets sin progreso no se evalúan.", "Resumen");
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.9, h: 4.9, fill: { color: C.sf }, line: { color: C.bd, width: 0.75 }, objectName: "Tarjeta KPI principal" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.9, h: 0.09, fill: { color: kc(K.pct) }, line: { type: "none" }, objectName: "Franja KPI principal" });
    s.addText("KPI JUSTIFICADO", { x: 0.8, y: 1.55, w: 4.3, h: 0.35, fontFace: FB, fontSize: 12, bold: true, color: C.mu, margin: 0, isTextBox: true, objectName: "Etiqueta KPI" });
    s.addText(nf(K.pct), { x: 0.8, y: 1.95, w: 4.3, h: 1.3, fontFace: FH, fontSize: 54, color: kc(K.pct), margin: 0, valign: "middle", fit: "shrink", isTextBox: true, objectName: "Valor KPI" });
    s.addText(gap == null ? "Sin tickets evaluables" : gap >= 0 ? (gap < .05 ? "Meta alcanzada (justo en la meta)" : "Meta alcanzada (" + sg(gap) + " pts sobre la meta)") : "Brecha de " + Math.abs(gap).toFixed(1).replace(".", ",") + " pts frente a la meta", { x: 0.8, y: 3.3, w: 4.3, h: 0.4, fontFace: FB, fontSize: 14, bold: true, color: kc(K.pct), margin: 0, isTextBox: true, objectName: "Estado frente a meta" });
    // barra de avance con marca de meta
    const bx = 0.8, bw = 4.3, by = 4.35;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: 0.26, fill: { color: "E5E8EE" }, line: { type: "none" }, objectName: "Barra fondo" });
    if (K.pct != null && K.pct > 0) s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: Math.max(0.02, bw * Math.min(100, K.pct) / 100), h: 0.26, fill: { color: kc(K.pct) }, line: { type: "none" }, objectName: "Barra KPI" });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw * meta / 100 - 0.015, y: by - 0.12, w: 0.03, h: 0.5, fill: { color: C.navy }, line: { type: "none" }, objectName: "Marca de meta" });
    s.addText("Meta " + meta + " %", { x: bx + bw * meta / 100 - 1.0, y: by + 0.4, w: 1.1, h: 0.3, fontFace: FB, fontSize: 10, bold: true, color: C.navy, align: "right", margin: 0, isTextBox: true, objectName: "Texto meta" });
    s.addText("Tickets dentro del 100 % de progreso ÷ tickets evaluados", { x: 0.8, y: 5.35, w: 4.3, h: 0.6, fontFace: FB, fontSize: 10, italic: true, color: C.mu, margin: 0, valign: "top", isTextBox: true, objectName: "Fórmula" });
    const gx = 5.7, cw = 2.2, ch = 2.35, gp = 0.15, aprob = K.a - K.o;
    const cards = [
      [K.n, "Tickets evaluados", "sin cancelados ni sin progreso", C.navy],
      [K.a, "Dentro del 100 %", aprob > 0 ? "incluye " + aprob + " justificados aprobados" : "atendidos a tiempo", C.gr],
      [K.f, "Fuera del 100 %", "tickets excedidos", K.f ? C.rd : C.navy],
      [nf(K.pO), "KPI original", "sin contar justificaciones", C.navy],
      [D.total, "Registrados en el rango", "todos los estados", C.navy],
      [D.canc + " / " + D.sinp, "Cancelados / sin progreso", "no entran al KPI", C.mu]
    ];
    cards.forEach((c, i) => card(s, gx + (i % 3) * (cw + gp), 1.3 + Math.floor(i / 3) * (ch + 0.2), cw, ch, c[0], c[1], c[2], c[3], String(c[0]).length > 7 ? 20 : 30));

    /* ---------- 3. Evolución ---------- */
    pres.addSection({ title: "Desempeño" });
    const SR = D.serie;
    const labs = SR.map(p => p.lab);
    const all = SR.flatMap(p => [p.pct, p.pO]).filter(v => v != null).concat([meta]);
    const mn = Math.max(0, Math.floor((Math.min.apply(null, all) - 5) / 10) * 10) / 100;
    s = content("Evolución del cumplimiento", "Un punto por " + gl + " según la fecha de registro del ticket. KPI justificado: incluye justificaciones aprobadas. KPI original: sin justificaciones. Meta " + meta + " %.", "Desempeño");
    s.addChart(pres.charts.LINE, [
      { name: "KPI justificado", labels: labs, values: SR.map(p => p.pct == null ? null : p.pct / 100) },
      { name: "KPI original", labels: labs, values: SR.map(p => p.pO == null ? null : p.pO / 100) },
      { name: "Meta " + meta + " %", labels: labs, values: SR.map(() => meta / 100) }
    ], Object.assign({ x: 0.5, y: 1.25, w: 8.6, h: 5.0, chartColors: [C.navy, C.gy, C.ora], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7, showLegend: true, legendPos: "b", showTitle: true, title: "Cumplimiento por " + gl, titleFontFace: FB, titleFontSize: 13, titleColor: C.tx }, baseChart, axis(mn)));
    const okN = SR.filter(p => p.pct != null && p.pct >= meta).length;
    const best = SR.filter(p => p.pct != null).slice().sort((a, b) => b.pct - a.pct || b.n - a.n)[0], worst = SR.filter(p => p.pct != null).slice().sort((a, b) => a.pct - b.pct || b.n - a.n)[0];
    const first = SR.find(p => p.pct != null), last = SR.slice().reverse().find(p => p.pct != null);
    const lect = [];
    if (SR.length) lect.push(okN + " de " + SR.length + " " + (gl === "mes" ? "meses" : gl === "día" ? "días" : "semanas") + " alcanzaron la meta.");
    if (best) lect.push("Mejor " + gl + ": " + best.labFull + " con " + nf(best.pct) + ".");
    if (worst && worst !== best) lect.push("Menor " + gl + ": " + worst.labFull + " con " + nf(worst.pct) + ".");
    if (first && last && first !== last) lect.push("Variación entre el primer y el último " + gl + ": " + sg(last.pct - first.pct) + " pts.");
    panel(s, 9.35, 1.3, 3.5, 4.9, "Lectura del periodo", lect);

    /* ---------- 4. Volumen dentro/fuera ---------- */
    s = content("Tickets dentro y fuera del 100 %", "Tickets evaluados por " + gl + ", separados entre los que quedaron dentro del 100 % de progreso (incluye justificaciones aprobadas) y los que quedaron fuera.", "Desempeño");
    const showV = SR.length <= 16;
    s.addChart(pres.charts.BAR, [
      { name: "Dentro del 100 %", labels: labs, values: SR.map(p => p.a) },
      { name: "Fuera del 100 %", labels: labs, values: SR.map(p => p.f) }
    ], Object.assign({ x: 0.5, y: 1.25, w: 8.6, h: 5.0, barDir: "col", barGrouping: "stacked", barGapWidthPct: 45, chartColors: [C.gr, C.rd], showLegend: true, legendPos: "b", showValue: showV, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 10, dataLabelFontBold: true, dataLabelFormatCode: "0;;;", showTitle: true, title: "Tickets evaluados por " + gl, titleFontFace: FB, titleFontSize: 13, titleColor: C.tx }, baseChart));
    const pf = K.n ? K.f / K.n * 100 : null, mf = SR.slice().sort((a, b) => b.f - a.f)[0];
    const l4 = ["Del total de " + K.n + " tickets evaluados, " + K.a + " quedaron dentro del 100 % y " + K.f + " fuera" + (pf != null ? " (" + nf(pf) + ")." : ".")];
    if (mf && mf.f > 0) l4.push("Mayor cantidad fuera del 100 %: " + mf.labFull + " con " + mf.f + " de " + mf.n + " tickets.");
    if (aprob > 0) l4.push(aprob + " ticket(s) vencido(s) cuentan como cumplidos por justificación aprobada.");
    panel(s, 9.35, 1.3, 3.5, 4.9, "Lectura del periodo", l4);

    /* ---------- 5-7. Rankings (grupo, especialista, categoría) ---------- */
    function ranking(title, header, list, max, notes) {
      if (!list || !list.length) return;
      const top = list.slice(0, max), s2 = content(title, notes, "Desempeño");
      s2.addChart(pres.charts.BAR, [{ name: "KPI", labels: top.map(x => cut(x.k, 26)), values: top.map(x => x.pct == null ? 0 : x.pct / 100) }],
        Object.assign({ x: 0.5, y: 1.25, w: 6.1, h: 5.0, barDir: "bar", catAxisOrientation: "maxMin", barGapWidthPct: 40, chartColors: top.map(x => kc(x.pct)), showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%", dataLabelFontSize: 10, dataLabelFontFace: FB, dataLabelColor: C.tx, showLegend: false, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 1.15, showTitle: true, title: "KPI por " + header.toLowerCase() + " (verde = alcanza la meta)", titleFontFace: FB, titleFontSize: 12, titleColor: C.tx }, baseChart, { valAxisHidden: true, valGridLine: { style: "none" }, catAxisLabelFontSize: 10 }));
      const H = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: C.navy }, align: "center", fontSize: 10 } });
      const rows = [[{ text: header, options: { bold: true, color: "FFFFFF", fill: { color: C.navy }, fontSize: 10 } }, H("Evaluados"), H("Dentro"), H("Fuera"), H("KPI")]];
      top.forEach((x, i) => { const z = i % 2 ? { color: "F8FAFD" } : { color: "FFFFFF" }; rows.push([{ text: cut(x.k, 30), options: { fill: z, fontSize: 10 } }, { text: String(x.n), options: { align: "center", fill: z, fontSize: 10 } }, { text: String(x.a), options: { align: "center", fill: z, fontSize: 10 } }, { text: String(x.f), options: { align: "center", fill: z, fontSize: 10, color: x.f ? C.rd : C.tx } }, { text: nf(x.pct), options: { align: "center", bold: true, color: kc(x.pct), fill: z, fontSize: 10 } }]) });
      s2.addTable(rows, { x: 6.85, y: 1.35, w: 6.0, colW: [2.5, 0.9, 0.8, 0.75, 1.05], rowH: 0.36, fontFace: FB, color: C.tx, border: { type: "solid", color: C.bd, pt: 0.75 }, valign: "middle", objectName: "Tabla " + header });
      const yEnd = 1.35 + 0.36 * (top.length + 1) + (list.length > max ? 0.7 : 0.2), okc = list.filter(x => x.pct != null && x.pct >= meta).length;
      const srt = list.filter(x => x.pct != null).slice().sort((a, b) => b.pct - a.pct || b.n - a.n), lt = [okc + " de " + list.length + " alcanzan la meta de " + meta + " %."];
      if (srt.length > 1) { lt.push("Mayor cumplimiento: " + cut(srt[0].k, 34) + " (" + nf(srt[0].pct) + ")."); lt.push("Menor cumplimiento: " + cut(srt[srt.length - 1].k, 34) + " (" + nf(srt[srt.length - 1].pct) + ")."); }
      if (6.2 - yEnd >= 1.5) panel(s2, 6.85, yEnd, 6.0, Math.min(6.2 - yEnd, 2.0), "Lectura", lt);
      if (list.length > max) s2.addText("Se muestran " + max + " de " + list.length + " (ordenados por tickets evaluados). El detalle completo está en la pestaña Rendimiento del aplicativo.", { x: 6.85, y: 1.35 + 0.36 * (top.length + 1) + 0.15, w: 6.0, h: 0.5, fontFace: FB, fontSize: 9, italic: true, color: C.mu, margin: 0, valign: "top", isTextBox: true, objectName: "Nota ranking" });
    }
    ranking("Cumplimiento por grupo", "Grupo", D.grupos, 8, "Cumplimiento (KPI justificado) por grupo responsable, ordenado por tickets evaluados.");
    ranking("Cumplimiento por especialista", "Especialista", D.esp, 10, "Cumplimiento (KPI justificado) por especialista, ordenado por tickets evaluados.");
    ranking("Cumplimiento por categoría", "Categoría", D.cat, 8, "Cumplimiento (KPI justificado) por categoría del ticket, ordenado por tickets evaluados.");

    /* ---------- 8. Auditoría de vencidos ---------- */
    pres.addSection({ title: "Control" });
    const A = D.aud;
    s = content("Auditoría de tickets vencidos", "Vencidos = tickets cerrados o resueltos con más del 100 % de progreso. Al aprobar una justificación completa (fecha real y sustento) el ticket cuenta como cumplido y el KPI se recalcula.", "Control");
    const aw = 2.3, ag = 0.2;
    [[A.venc, "Vencidos", "cerrados fuera del 100 %", C.navy], [A.pend, "Sin justificar", "sin fecha real o sin sustento", A.pend ? C.am : C.navy], [A.sust, "Con sustento", "fecha real y sustento", C.gr], [A.aprob, "Aprobadas", "cuentan como cumplidas", C.gr], [A.rech, "Desaprobadas", "siguen fuera del KPI", A.rech ? C.rd : C.navy]]
      .forEach((c, i) => card(s, 0.5 + i * (aw + ag), 1.3, aw, 1.55, c[0], c[1], c[2], c[3], 28));
    // tabla de mayores desvíos
    s.addText("Mayores desvíos sin aprobación", { x: 0.5, y: 3.05, w: 7.3, h: 0.35, fontFace: FH, fontSize: 12, color: C.navy, margin: 0, isTextBox: true, objectName: "Título desvíos" });
    const HH = (t, al) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: C.navy }, align: al || "center", fontSize: 10 } });
    const tr = [[HH("Ticket", "left"), HH("Especialista", "left"), HH("Progreso"), HH("Estado")]];
    (A.top || []).forEach((x, i) => { const z = { color: i % 2 ? "F8FAFD" : "FFFFFF" }; tr.push([{ text: cut(x.c + " · " + x.s, 44), options: { fill: z, fontSize: 9 } }, { text: cut(x.r, 22), options: { fill: z, fontSize: 9 } }, { text: x.p.toFixed(0) + " %", options: { align: "center", bold: true, color: C.rd, fill: z, fontSize: 10 } }, { text: x.est, options: { align: "center", fill: z, fontSize: 9 } }]) });
    if (tr.length > 1) s.addTable(tr, { x: 0.5, y: 3.45, w: 7.3, colW: [3.2, 1.9, 1.1, 1.1], rowH: 0.38, fontFace: FB, color: C.tx, border: { type: "solid", color: C.bd, pt: 0.75 }, valign: "middle", objectName: "Tabla desvíos" });
    else s.addText("No hay tickets vencidos sin aprobar en el rango.", { x: 0.5, y: 3.5, w: 7.3, h: 0.5, fontFace: FB, fontSize: 12, color: C.mu, margin: 0, isTextBox: true, objectName: "Sin desvíos" });
    s.addChart(pres.charts.BAR, [{ name: "KPI", labels: ["KPI original", "KPI justificado", "Meta"], values: [(K.pO || 0) / 100, (K.pct || 0) / 100, meta / 100] }],
      Object.assign({ x: 8.1, y: 3.0, w: 4.75, h: 3.25, barDir: "col", barGapWidthPct: 50, chartColors: [C.gy, kc(K.pct), C.ora], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%", dataLabelFontSize: 11, dataLabelFontFace: FB, dataLabelColor: C.tx, showLegend: false, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 1.15, showTitle: true, title: "Efecto de las justificaciones", titleFontFace: FB, titleFontSize: 12, titleColor: C.tx }, baseChart, { valAxisHidden: true, valGridLine: { style: "none" } }));

    /* ---------- 9. Estado y progreso ---------- */
    s = content("Estado de los tickets y avance de los abiertos", "Estados de todos los tickets registrados en el rango (incluye cancelados). Los tickets abiertos son los que están en Nuevo, Asignado o En Proceso; se agrupan por su % de progreso.", "Control");
    const EP = ["002336", "1B7F4B", "E07026", "2F5FD0", "A0A7B5", "D22C29", "8E6BBF", "C9A227"];
    const est = D.estados.filter(x => x.n > 0);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Tickets", labels: est.map(x => x.e), values: est.map(x => x.n) }],
      { x: 0.5, y: 1.25, w: 5.9, h: 5.0, holeSize: 55, chartColors: EP.slice(0, est.length), showLegend: true, legendPos: "r", legendFontFace: FB, legendFontSize: 11, legendColor: C.tx, showValue: true, showPercent: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelFontFace: FB, showTitle: true, title: "Tickets por estado (" + D.total + ")", titleFontFace: FB, titleFontSize: 13, titleColor: C.tx });
    const AB = D.abiertos;
    s.addChart(pres.charts.BAR, [{ name: "Abiertos", labels: AB.B.map(b => b.l), values: AB.B.map(b => b.n) }],
      Object.assign({ x: 6.6, y: 1.25, w: 6.25, h: 3.7, barDir: "col", barGapWidthPct: 40, chartColors: ["1B7F4B", "7FB069", "E8B923", "E07026", "C62828"], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFontFace: FB, dataLabelColor: C.tx, dataLabelFormatCode: "0;;;", showLegend: false, valAxisHidden: true, showTitle: true, title: "Tickets abiertos por % de progreso (" + AB.total + ")", titleFontFace: FB, titleFontSize: 13, titleColor: C.tx }, baseChart, { valAxisHidden: true, valGridLine: { style: "none" } }));
    const gt = AB.B[4] ? AB.B[4].n : 0;
    s.addShape(pres.shapes.RECTANGLE, { x: 6.6, y: 5.15, w: 6.25, h: 1.05, fill: { color: gt ? "FDEAEA" : "E3F5EA" }, line: { type: "none" }, objectName: "Aviso abiertos" });
    s.addText(AB.total ? (gt ? gt + " ticket(s) abierto(s) ya superan el 100 % de progreso y afectarán el KPI si se cierran así." : "Ningún ticket abierto supera el 100 % de progreso.") + (AB.total - AB.conP > 0 ? " " + (AB.total - AB.conP) + " abierto(s) sin progreso registrado." : "") : "No hay tickets abiertos en el rango.", { x: 6.8, y: 5.2, w: 5.85, h: 0.95, fontFace: FB, fontSize: 12, bold: true, color: gt ? C.rd : C.gr, margin: 0, valign: "middle", isTextBox: true, objectName: "Texto aviso abiertos" });

    /* ---------- 10. Planes de acción ---------- */
    const PL = D.planes;
    if (PL && PL.n) {
      pres.addSection({ title: "Acciones" });
      const per = 6, pages = Math.ceil(PL.lista.length / per);
      for (let pg = 0; pg < pages; pg++) {
        s = content("Planes de acción" + (pages > 1 ? " (" + (pg + 1) + "/" + pages + ")" : ""), "Un plan por causa raíz detectada. Alerta: Vencido = fecha compromiso superada; Por vencer = 7 días o menos. El indicador es el último valor semanal de la medición automática del plan.", "Acciones");
        const pw = 1.95, pgp = 0.12, ty = pg === 0 ? 2.5 : 1.3;
        if (pg === 0) [[PL.n, "Planes registrados", C.navy], [PL.ok, "Cumplidos", C.gr], [PL.cu, "En curso", C.ac], [PL.pe, "Pendientes", C.am], [nf(PL.avg).replace(" %", "%"), "Avance promedio", C.navy], [PL.ve, "Con fecha vencida", PL.ve ? C.rd : C.navy]]
          .forEach((c, i) => { const x = 0.5 + i * (pw + pgp); s.addShape(pres.shapes.RECTANGLE, { x, y: 1.25, w: pw, h: 1.0, fill: { color: C.sf }, line: { color: C.bd, width: 0.75 }, objectName: "Resumen " + c[1] }); s.addShape(pres.shapes.RECTANGLE, { x, y: 1.25, w: 0.07, h: 1.0, fill: { color: c[2] }, line: { type: "none" }, objectName: "Acento " + c[1] }); s.addText(String(c[0]), { x: x + 0.18, y: 1.3, w: pw - 0.25, h: 0.55, fontFace: FH, fontSize: 22, color: c[2], margin: 0, valign: "middle", isTextBox: true, objectName: "Valor " + c[1] }); s.addText(c[1], { x: x + 0.18, y: 1.85, w: pw - 0.25, h: 0.3, fontFace: FB, fontSize: 10, bold: true, color: C.tx, margin: 0, valign: "middle", isTextBox: true, objectName: "Etiqueta " + c[1] }) });
        const HP = (t, al) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: C.navy }, align: al || "center", fontSize: 10 } });
        const tb = [[HP("Plan"), HP("Causa detectada", "left"), HP("Responsable", "left"), HP("Compromiso"), HP("Alerta"), HP("Estado"), HP("Avance"), HP("Indicador")]];
        PL.lista.slice(pg * per, pg * per + per).forEach((p, i) => {
          const z = { color: i % 2 ? "F8FAFD" : "FFFFFF" }, ac = p.alerta === "Vencido" ? C.rd : p.alerta === "Por vencer" ? C.am : p.alerta === "Cumplido" || p.alerta === "En plazo" ? C.gr : C.mu;
          const ec = p.est.indexOf("Cumplido") === 0 ? C.gr : p.est === "En curso" ? C.ac : C.am;
          tb.push([{ text: p.id, options: { bold: true, align: "center", fill: z, fontSize: 10 } }, { text: cut(p.cause, 90), options: { fill: z, fontSize: 9.5 } }, { text: cut(p.resp || "—", 24), options: { fill: z, fontSize: 9.5 } }, { text: p.fecha || "—", options: { align: "center", fill: z, fontSize: 9.5 } }, { text: p.alerta, options: { align: "center", bold: true, color: ac, fill: z, fontSize: 9.5 } }, { text: p.est === "Cumplido / validado" ? "Cumplido" : p.est, options: { align: "center", bold: true, color: ec, fill: z, fontSize: 9.5 } }, { text: p.av + " %", options: { align: "center", fill: z, fontSize: 10 } }, { text: p.ind || "—", options: { align: "center", fill: z, fontSize: 9.5 } }]);
        });
        s.addTable(tb, { x: 0.5, y: ty, w: 12.3, colW: [0.8, 3.9, 1.8, 1.2, 1.15, 1.25, 0.9, 1.3], rowH: [0.38].concat(PL.lista.slice(pg * per, pg * per + per).map(() => 0.55)), fontFace: FB, color: C.tx, border: { type: "solid", color: C.bd, pt: 0.75 }, valign: "middle", objectName: "Tabla planes" });
      }
    }

    /* ---------- 11. Conclusiones ---------- */
    s = content("Conclusiones y puntos de atención", "Resumen automático construido con los datos del rango seleccionado.", PL && PL.n ? "Acciones" : "Control");
    panel(s, 0.5, 1.3, 6.05, 4.9, "Lo que muestran los datos", D.notas.lectura, 15);
    panel(s, 6.8, 1.3, 6.05, 4.9, "Requiere atención", D.notas.atencion.length ? D.notas.atencion : ["Sin puntos de atención adicionales en el rango."], 15);

    /* ---------- 12. Cierre ---------- */
    s = pres.addSlide({ masterName: "GRACIAS", sectionTitle: PL && PL.n ? "Acciones" : "Control" });
    s.addNotes("Fin de la presentación. Generada el " + D.fechaGen + " desde «Rendimiento de atenciones · Aranda».");
    return pres;
  }
  const api = { build };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.PPTG = api;
})(typeof window !== "undefined" ? window : globalThis);
