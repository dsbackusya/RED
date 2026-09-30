// OCR con modelos PaddleOCR (PP-OCRv4) sobre onnxruntime-web, sin OpenCV. Detección (DB) + reconocimiento (CTC).
window.Paddle = (() => {
  const ORT = 'https://cdn.jsdelivr.net/npm/onnxruntime-web@1.18.0/dist/';
  const MODELOS = 'https://cdn.jsdelivr.net/npm/@gutenye/ocr-models@1.4.2/assets/';
  let det = null, rec = null, dic = null;

  const cargarScript = src => new Promise((ok, no) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => no(new Error('No se pudo cargar ' + src)); document.head.appendChild(s); });

  async function iniciar(progreso) {
    if (det) return;
    if (!window.ort) await cargarScript(ORT + 'ort.min.js');
    ort.env.wasm.wasmPaths = ORT; ort.env.wasm.numThreads = 1;
    progreso && progreso('Cargando modelo de lectura…');
    const [d, r, k] = await Promise.all([
      fetch(MODELOS + 'ch_PP-OCRv4_det_infer.onnx').then(x => x.arrayBuffer()),
      fetch(MODELOS + 'ch_PP-OCRv4_rec_infer.onnx').then(x => x.arrayBuffer()),
      fetch(MODELOS + 'ppocr_keys_v1.txt').then(x => x.text())]);
    det = await ort.InferenceSession.create(d, { executionProviders: ['wasm'] });
    rec = await ort.InferenceSession.create(r, { executionProviders: ['wasm'] });
    dic = ['<blank>', ...k.split(/\r?\n/).filter(x => x.length), ' '];
  }

  function aCanvas(img, w, h) { const c = document.createElement('canvas'); c.width = w; c.height = h; c.getContext('2d').drawImage(img, 0, 0, w, h); return c; }

  // ---- detección ----
  async function detectar(img, lado) {
    const k = Math.min(1, lado / Math.max(img.width, img.height));
    let w = Math.max(32, Math.round(img.width * k / 32) * 32), h = Math.max(32, Math.round(img.height * k / 32) * 32);
    const d = aCanvas(img, w, h).getContext('2d').getImageData(0, 0, w, h).data, n = w * h, x = new Float32Array(3 * n);
    const M = [0.485, 0.456, 0.406], S = [0.229, 0.224, 0.225];
    for (let i = 0; i < n; i++) for (let c = 0; c < 3; c++) x[c * n + i] = (d[i * 4 + c] / 255 - M[c]) / S[c];
    const out = await det.run({ [det.inputNames[0]]: new ort.Tensor('float32', x, [1, 3, h, w]) });
    return { P: out[det.outputNames[0]].data, w, h };
  }

  // componentes conexas de la mapa de probabilidad -> cajas orientadas (PCA)
  function cajas(P, w, h, sx, sy) {
    const thr = 0.3, lab = new Int32Array(w * h), res = [], pila = new Int32Array(w * h);
    let id = 0;
    for (let s = 0; s < w * h; s++) {
      if (lab[s] || P[s] <= thr) continue;
      id++; let sp = 0; pila[sp++] = s; lab[s] = id;
      const px = []; let suma = 0;
      while (sp) {
        const p = pila[--sp], x = p % w, y = (p / w) | 0; px.push(p); suma += P[p];
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const nx = x + dx, ny = y + dy; if (nx < 0 || ny < 0 || nx >= w || ny >= h) continue;
          const q = ny * w + nx; if (!lab[q] && P[q] > thr) { lab[q] = id; pila[sp++] = q; }
        }
      }
      if (px.length < 25 || suma / px.length < 0.55) continue;
      let mx = 0, my = 0; for (const p of px) { mx += p % w; my += (p / w) | 0; } mx /= px.length; my /= px.length;
      let cxx = 0, cyy = 0, cxy = 0;
      for (const p of px) { const dx = (p % w) - mx, dy = ((p / w) | 0) - my; cxx += dx * dx; cyy += dy * dy; cxy += dx * dy; }
      let ang = 0.5 * Math.atan2(2 * cxy, cxx - cyy);
      const ex = Math.cos(ang), ey = Math.sin(ang);
      let u0 = 1e9, u1 = -1e9, v0 = 1e9, v1 = -1e9;
      for (const p of px) { const dx = (p % w) - mx, dy = ((p / w) | 0) - my, u = dx * ex + dy * ey, v = -dx * ey + dy * ex; if (u < u0) u0 = u; if (u > u1) u1 = u; if (v < v0) v0 = v; if (v > v1) v1 = v; }
      let L = u1 - u0 + 1, H = v1 - v0 + 1;
      if (H > L) { // eje mayor es el vertical: no es una línea horizontal
        ang += Math.PI / 2; [L, H] = [H, L]; const t = u0; u0 = v0; v0 = t; const t2 = u1; u1 = v1; v1 = t2;
        // recalcular extremos con el nuevo eje
        const ex2 = Math.cos(ang), ey2 = Math.sin(ang); u0 = 1e9; u1 = -1e9; v0 = 1e9; v1 = -1e9;
        for (const p of px) { const dx = (p % w) - mx, dy = ((p / w) | 0) - my, u = dx * ex2 + dy * ey2, v = -dx * ey2 + dy * ex2; if (u < u0) u0 = u; if (u > u1) u1 = u; if (v < v0) v0 = v; if (v > v1) v1 = v; }
        L = u1 - u0 + 1; H = v1 - v0 + 1;
      }
      if (Math.abs(ang) > Math.PI / 2) ang -= Math.sign(ang) * Math.PI;
      const dd = (L * H * 1.6) / (2 * (L + H));            // unclip del DB
      L += 2 * dd; H += 2 * dd;
      const ec = Math.cos(ang), es = Math.sin(ang), uc = (u0 + u1) / 2, vc = (v0 + v1) / 2;
      const cx = (mx + uc * ec - vc * es) * sx, cy = (my + uc * es + vc * ec) * sy;
      res.push({ cx, cy, L: L * sx, H: H * sy, ang, sx, sy });
    }
    return res;
  }

  // recorte rotado con muestreo bilineal -> tensor de reconocimiento
  function recortar(datos, W, Hh, b) {
    const L = b.L, H = b.H, outH = 48, outW = Math.max(16, Math.min(1400, Math.round(outH * L / H)));
    const ec = Math.cos(b.ang), es = Math.sin(b.ang), x = new Float32Array(3 * outH * outW), n = outH * outW;
    for (let j = 0; j < outH; j++) for (let i = 0; i < outW; i++) {
      const u = ((i + 0.5) / outW - 0.5) * L, v = ((j + 0.5) / outH - 0.5) * H;
      const sx = b.cx + u * ec - v * es, sy = b.cy + u * es + v * ec;
      const x0 = Math.floor(sx), y0 = Math.floor(sy), fx = sx - x0, fy = sy - y0;
      for (let c = 0; c < 3; c++) {
        const g = (xx, yy) => { xx = Math.min(W - 1, Math.max(0, xx)); yy = Math.min(Hh - 1, Math.max(0, yy)); return datos[(yy * W + xx) * 4 + c]; };
        const val = (g(x0, y0) * (1 - fx) + g(x0 + 1, y0) * fx) * (1 - fy) + (g(x0, y0 + 1) * (1 - fx) + g(x0 + 1, y0 + 1) * fx) * fy;
        x[c * n + j * outW + i] = (val / 255 - 0.5) / 0.5;
      }
    }
    return new ort.Tensor('float32', x, [1, 3, outH, outW]);
  }
  async function reconocer(tensor) {
    const out = await rec.run({ [rec.inputNames[0]]: tensor }), o = out[rec.outputNames[0]], T = o.dims[1], C = o.dims[2], d = o.data;
    let txt = '', prev = 0, conf = 0, cnt = 0;
    for (let t = 0; t < T; t++) {
      let mi = 0, mv = -1; for (let c = 0; c < C; c++) { const v = d[t * C + c]; if (v > mv) { mv = v; mi = c; } }
      if (mi !== 0 && mi !== prev) { txt += dic[mi] || ''; conf += mv; cnt++; }
      prev = mi;
    }
    return { text: txt, conf: cnt ? conf / cnt : 0 };
  }

  async function leer(img, { lado = 1280, progreso } = {}) {
    await iniciar(progreso);
    progreso && progreso('Detectando texto…');
    const { P, w, h } = await detectar(img, lado);
    const bs = cajas(P, w, h, img.width / w, img.height / h);
    const cv = document.createElement('canvas'); cv.width = img.width; cv.height = img.height; cv.getContext('2d').drawImage(img, 0, 0);
    const datos = cv.getContext('2d').getImageData(0, 0, img.width, img.height).data;
    const lineas = [];
    let i = 0;
    for (const b of bs) {
      if (b.H < 6 || b.L < 6) continue;
      const r = await reconocer(recortar(datos, img.width, img.height, b));
      progreso && progreso(`Leyendo líneas ${++i}/${bs.length}`, i / bs.length);
      if (r.text.trim()) lineas.push({ ...r, cx: b.cx, cy: b.cy, L: b.L, H: b.H, ang: b.ang });
    }
    // orden de lectura corrigiendo la inclinación global
    const angs = lineas.map(l => l.ang).sort((a, b) => a - b), sk = angs.length ? angs[angs.length >> 1] : 0, cs = Math.cos(sk), sn = Math.sin(sk);
    lineas.forEach(l => { l.ry = -l.cx * sn + l.cy * cs; l.rx = l.cx * cs + l.cy * sn; });
    lineas.sort((a, b) => a.ry - b.ry);
    const filas = [];
    lineas.forEach(l => { const f = filas.find(f => Math.abs(f.ry - l.ry) < 0.55 * Math.min(f.H, l.H)); if (f) { f.items.push(l); f.ry = (f.ry + l.ry) / 2; } else filas.push({ ry: l.ry, H: l.H, items: [l] }); });
    const texto = filas.sort((a, b) => a.ry - b.ry).map(f => f.items.sort((a, b) => a.rx - b.rx).map(x => x.text).join('  ')).join('\n');
    return { texto, lineas, filas: filas.length };
  }
  return { leer, iniciar };
})();
