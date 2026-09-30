// Lectura gratuita de facturas (fotos): QR SUNAT + OCR (PaddleOCR, respaldo Tesseract) en el navegador. Todo se revisa a mano antes de guardar.
window.OCR = (() => {
  const RUC_DINET = '20427919111';
  const ALIAS_NODOS = { NASCA: 'NAZCA', CANETE: 'CANETE' };
  let worker = null;

  const cargarScript = src => new Promise((ok, fallo) => {
    const s = document.createElement('script'); s.src = src; s.onload = ok;
    s.onerror = () => fallo(new Error('No se pudo cargar ' + src)); document.head.appendChild(s);
  });
  async function libs() {
    if (!window.jsQR) await cargarScript('https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js');
  }
  const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();

  function imagenDe(file) {
    return new Promise((ok, fallo) => {
      const url = URL.createObjectURL(file), i = new Image();
      i.onload = () => { URL.revokeObjectURL(url); ok(i); };
      i.onerror = () => { URL.revokeObjectURL(url); fallo(new Error('No se pudo abrir la imagen')); };
      i.src = url;
    });
  }
  function canvasDe(img, lado, gris, contraste) {
    const k = Math.min(1, lado / Math.max(img.width, img.height));
    const c = document.createElement('canvas'); c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
    const x = c.getContext('2d'); x.drawImage(img, 0, 0, c.width, c.height);
    if (gris) {
      const d = x.getImageData(0, 0, c.width, c.height), p = d.data;
      for (let i = 0; i < p.length; i += 4) {
        let g = 0.3 * p[i] + 0.59 * p[i + 1] + 0.11 * p[i + 2];
        if (contraste) g = Math.max(0, Math.min(255, (g - 110) * 1.6 + 128));
        p[i] = p[i + 1] = p[i + 2] = g;
      }
      x.putImageData(d, 0, 0);
    }
    return c;
  }
  function leerQR(img) {
    for (const lado of [Math.max(img.width, img.height), 1800, 1100]) {
      const c = canvasDe(img, lado, false), d = c.getContext('2d').getImageData(0, 0, c.width, c.height);
      const r = window.jsQR(d.data, c.width, c.height, { inversionAttempts: 'attemptBoth' });
      if (r && r.data) return r.data;
    }
    return null;
  }

  // RUC peruano: dígito verificador módulo 11
  function rucValido(r) {
    if (!/^(10|15|17|20)\d{9}$/.test(r)) return false;
    const w = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]; let s = 0;
    for (let i = 0; i < 10; i++) s += Number(r[i]) * w[i];
    let d = 11 - (s % 11); if (d === 10) d = 0; if (d === 11) d = 1;
    return d === Number(r[10]);
  }
  const dinero = s => Number(String(s).replace(',', '.'));
  const montos = l => (l.match(/\b\d{1,5}[.,]\d{2}\b/g) || []).map(dinero);
  const cerca = (a, b) => Math.abs(a - b) <= 0.03;
  const iso = (d, m, a) => `${a}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;

  function fechaValida(d, m, a, ref) {
    d = +d; m = +m; a = +a;
    if (a < 2025 || a > 2027 || m < 1 || m > 12 || d < 1 || d > 31) return '';
    const f = iso(d, m, a);
    if (ref && Math.abs(new Date(f) - new Date(ref)) > 10 * 86400000) return '';
    return f;
  }

  function parsear(texto, qr, ref, nodos) {
    const R = { factura: '', importe: null, fecha: '', ruc: '', guia: '', destino: '', destinosVistos: [], bultos: null, peso: null, aviso: [], viaQR: false };
    const T = texto.replace(/\r/g, '');
    const lineas = T.split('\n');

    // --- QR SUNAT: RUC|tipo|serie|correlativo|igv|total|fecha|tipoDocCli|docCli
    let q = null;
    if (qr) {
      const p = qr.split('|');
      if (p.length >= 9 && rucValido(p[0]) && /^[A-Z0-9]{4}$/.test(p[2]) && /^\d+$/.test(p[3])) q = p;
    }

    // --- factura
    const reFac = /(?<![0-9])(F[A-Z0-9]{3}|B[A-Z0-9]{3}|E[0-9]{3}|EB[0-9]{2})\s*[-–—]\s*(\d{4,8})(?!\d)/g;
    const cuenta = new Map();
    for (const m of T.matchAll(reFac)) { const k = m[1] + '-' + m[2]; cuenta.set(k, (cuenta.get(k) || 0) + 1); }
    const impresa = [...cuenta].sort((a, b) => b[1] - a[1])[0]?.[0] || '';
    if (q) {
      const serie = q[2], corr = q[3];
      const igual = [...cuenta.keys()].find(k => k.startsWith(serie + '-') && Number(k.split('-')[1]) === Number(corr));
      R.factura = igual || serie + '-' + corr.padStart(8, '0');
      R.importe = dinero(q[5]); R.ruc = q[0]; R.viaQR = true;
      const f = q[6].includes('/') ? q[6].split('/') : q[6].split('-').reverse();
      R.fecha = fechaValida(f[0], f[1], f[2], null); // la fecha del QR es exacta: no se compara con el registro
    } else R.factura = impresa;

    // --- importe por texto (si no hubo QR)
    if (R.importe == null) {
      const info = [], ocurr = new Map();
      lineas.forEach(l0 => {
        const l = norm(l0);
        const esIgv = /IGV|I\.G\.V/.test(l), esBase = /GRAVADA|EXONERAD|INAFECT|GRATUIT|DESCUENTO|SUB\s*-?\s*TOTAL|UNIT|V\.?\s*UNI|PESO|\bKG\b/.test(l);
        const esTotal = /TOTAL|IMPORTE|A\s*PAGAR|\bSON\s*[:;]/.test(l) && !esBase && !esIgv;
        montos(l0).forEach(v => {
          info.push({ v, esIgv, esBase, esTotal });
          if (!esIgv && !esBase) { const o = ocurr.get(v) || { n: 0, t: 0 }; o.n++; if (esTotal) o.t++; ocurr.set(v, o); }
        });
      });
      const bases = [...new Set(info.filter(x => x.esBase && x.v > 0).map(x => x.v))];
      let cand = [...ocurr].filter(([v]) => v > 0).map(([v, o]) => ({ v, consistente: bases.some(b => cerca(b * 1.18, v)), score: o.n + o.t * 3 + (o.n >= 2 ? 2 : 0) }));
      cand.forEach(c => { if (c.consistente) c.score += 5; });
      // descartar montos que son el IGV de otro candidato
      cand = cand.filter(c => !cand.some(d => d !== c && cerca(d.v * 0.18 / 1.18, c.v)));
      cand.sort((a, b) => b.score - a.score || b.v - a.v);
      const top = cand[0];
      if (top && (top.consistente || top.score > 2)) R.importe = top.v;
      else if (bases.length) {
        R.importe = Math.round(Math.max(...bases) * 1.18 * 100) / 100;
        R.aviso.push(`Total estimado desde el subtotal ${Math.max(...bases).toFixed(2)} + IGV: verifícalo en la foto.`);
      } else if (top) R.importe = top.v;
      if (R.importe != null && !R.aviso.length && !info.some(x => x.esTotal && cerca(x.v, R.importe)) && !(top && top.consistente))
        R.aviso.push('El importe no apareció junto a la palabra TOTAL: verifícalo en la foto.');
    }

    // --- fecha (texto)
    if (!R.fecha) {
      for (const m of T.matchAll(/(\d{2})\s*[\/-]\s*(\d{2})\s*[\/-]\s*(20\d{2})|(20\d{2})\s*[\/-]\s*(\d{2})\s*[\/-]\s*(\d{2})/g)) {
        const f = m[1] ? fechaValida(m[1], m[2], m[3], ref) : fechaValida(m[6], m[5], m[4], ref);
        if (f) { R.fecha = f; break; }
      }
    }

    // --- guía de remisión
    const g = T.replace(/EG[O0](\d)/g, 'EG0$1').match(/EG0\d\s*[-–—]\s*(\d{4,8})/);
    if (g) R.guia = 'EG07-' + g[1].padStart(8, '0');

    // --- RUC del emisor
    if (!R.ruc) {
      for (const l of lineas) {
        if (!/RUC/i.test(l) || /CONTRIBUY|REGISTRO|UNICO|ÚNICO/i.test(l)) continue;
        const m = [...l.matchAll(/(?<!\d)(\d{11})(?!\d)/g)].map(x => x[1]).find(r => r !== RUC_DINET && rucValido(r));
        if (m) { R.ruc = m; break; }
      }
    }

    // --- peso y bultos
    const pe = T.match(/PESO(?:\s*ENVIO)?\s*[:\-]?\s*(\d{1,4}(?:[.,]\d{1,2})?)/i) || T.match(/\((\d{1,4}(?:[.,]\d+)?)\s*K[gGoO]\)/);
    if (pe) R.peso = dinero(pe[1]);
    const bu = T.match(/(?:^|\n)\s*(\d{1,3})(?:[.,]00)?\s+(?:CAJAS?|PAQUETES?|BULTOS?)\b/i);
    if (bu) R.bultos = Number(bu[1]);

    // --- destino: nodos del catálogo que aparecen en el texto
    const tn = ' ' + norm(T).replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ') + ' ';
    const despues = tn.match(/ DESTINO ([A-Z ]{3,25})/);
    const puntos = [];
    (nodos || []).forEach(n => {
      const claves = [n, ...Object.entries(ALIAS_NODOS).filter(([, v]) => v === n).map(([k]) => k)];
      let cnt = 0;
      claves.forEach(k => { cnt += tn.split(' ' + k + ' ').length - 1; });
      if (!cnt) return;
      let s = cnt + (despues && claves.some(k => despues[1].trim().startsWith(k)) ? 5 : 0);
      if (n === 'ICA') s -= 1;
      puntos.push([n, s]);
    });
    puntos.sort((a, b) => b[1] - a[1]);
    R.destinosVistos = puntos.map(p => p[0]);
    R.destino = puntos[0]?.[0] || (despues ? despues[1].trim() : '');
    return R;
  }

  async function obtenerWorker(progreso) {
    if (!window.Tesseract) await cargarScript('https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js');
    if (!worker) worker = await window.Tesseract.createWorker('spa', 1, { logger: m => progreso && m.status && progreso(m.status, m.progress) });
    return worker;
  }

  // Lector principal: PaddleOCR (paddle.js). Respaldo: Tesseract si Paddle falla o no encuentra factura e importe.
  // file: imagen de la factura. ref: fecha del registro (YYYY-MM-DD). nodos: lista de nodos del catálogo.
  async function leer(file, { ref, nodos, progreso } = {}) {
    await libs();
    const img = await imagenDe(file);
    progreso && progreso('Buscando código QR…', 0);
    const qr = leerQR(img);
    let texto = '', R = null, motor = '';
    try {
      if (!window.Paddle) await cargarScript('paddle.js?v=3');
      const P = await window.Paddle.leer(img, { progreso });
      texto = P.texto; motor = 'PaddleOCR'; R = parsear(texto, qr, ref, nodos);
    } catch (e) { console.warn('PaddleOCR no disponible, se usa Tesseract:', e); }
    if (!R || !(R.factura && R.importe != null)) {
      const w = await obtenerWorker(progreso);
      const pasadas = [{ lado: 2000, contraste: false, psm: '6' }, { lado: 2600, contraste: true, psm: '4' }];
      for (const p of pasadas) {
        await w.setParameters({ tessedit_pageseg_mode: p.psm });
        const { data } = await w.recognize(canvasDe(img, p.lado, true, p.contraste));
        texto = texto ? texto + '\n' + data.text : data.text;
        motor = motor ? motor + ' + Tesseract' : 'Tesseract';
        R = parsear(texto, qr, ref, nodos);
        if (R.factura && R.importe != null) break;
      }
    }
    R.texto = texto; R.qr = qr; R.motor = motor;
    return R;
  }

  // Comprime una imagen a <= maxBytes (escala de grises, WebP si el navegador lo permite, si no JPEG).
  async function comprimir(file, maxBytes = 20 * 1024) {
    const img = await imagenDe(file);
    const aBlob = (c, tipo, q) => new Promise(ok => c.toBlob(ok, tipo, q));
    const prueba = await aBlob(canvasDe(img, 64, false), 'image/webp', 0.5);
    const tipo = prueba && prueba.type === 'image/webp' ? 'image/webp' : 'image/jpeg';
    for (const lado of [1400, 1200, 1050, 900, 800, 700, 600, 520, 440, 380, 320]) {
      const c = canvasDe(img, lado, true, false);
      let lo = 0.3, hi = 0.85, mejor = null;
      const min = await aBlob(c, tipo, lo);
      if (!min || min.size > maxBytes) continue;
      mejor = { blob: min, calidad: lo };
      for (let i = 0; i < 6; i++) {
        const q = (lo + hi) / 2, b = await aBlob(c, tipo, q);
        if (b && b.size <= maxBytes) { mejor = { blob: b, calidad: q }; lo = q; } else hi = q;
      }
      return { blob: mejor.blob, tipo, ext: tipo === 'image/webp' ? 'webp' : 'jpg', ancho: c.width, alto: c.height, calidad: Math.round(mejor.calidad * 100) };
    }
    throw new Error('No se pudo comprimir la imagen a ' + Math.round(maxBytes / 1024) + ' KB');
  }

  return { leer, parsear, rucValido, comprimir };
})();
