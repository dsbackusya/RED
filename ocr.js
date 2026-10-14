// Lectura gratuita de facturas (fotos): QR SUNAT + OCR Tesseract (modelo liviano, ~3 MB) en el navegador. Todo se revisa a mano antes de guardar.
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
  // Ubica el papel (la zona clara más grande de la foto) y devuelve un canvas recortado. Con fondos cargados (cartón con letras grandes, tableros) el OCR se confunde y no lee el encabezado.
  function recortarPapel(img, fu = 0.5) {
    const k = 240 / Math.max(img.width, img.height), w = Math.max(8, Math.round(img.width * k)), h = Math.max(8, Math.round(img.height * k)), n = w * h;
    const c = document.createElement('canvas'); c.width = w; c.height = h; const cx = c.getContext('2d'); cx.drawImage(img, 0, 0, w, h);
    const p = cx.getImageData(0, 0, w, h).data, g = new Uint8Array(n), hist = new Array(256).fill(0);
    for (let i = 0, q = 0; q < n; i += 4, q++) { const v = Math.round(0.3 * p[i] + 0.59 * p[i + 1] + 0.11 * p[i + 2]); g[q] = v; hist[v]++; }
    let suma = 0; for (let v = 0; v < 256; v++) suma += v * hist[v];
    let sB = 0, wB = 0, mx = 0, umbral = 128, oscuro = 60;   // umbral de Otsu: separa lo claro (papel) de lo oscuro (fondo)
    for (let v = 0; v < 256; v++) { wB += hist[v]; if (!wB) continue; const wF = n - wB; if (!wF) break; sB += v * hist[v]; const d = sB / wB - (suma - sB) / wF, e = wB * wF * d * d; if (e > mx) { mx = e; umbral = v; oscuro = sB / wB; } }
    umbral = Math.round(oscuro + (umbral - oscuro) * fu);   // se baja el umbral hacia el fondo para no perder las zonas del papel con sombra
    const vis = new Uint8Array(n), cola = new Int32Array(n); let mejor = null;
    for (let s = 0; s < n; s++) {
      if (vis[s] || g[s] <= umbral) continue;
      let a = 0, b = 0, x0 = w, x1 = 0, y0 = h, y1 = 0, cnt = 0; cola[b++] = s; vis[s] = 1;
      while (a < b) {
        const q = cola[a++], qx = q % w, qy = (q - qx) / w; cnt++;
        if (qx < x0) x0 = qx; if (qx > x1) x1 = qx; if (qy < y0) y0 = qy; if (qy > y1) y1 = qy;
        for (const nq of [qx > 0 ? q - 1 : -1, qx < w - 1 ? q + 1 : -1, qy > 0 ? q - w : -1, qy < h - 1 ? q + w : -1]) if (nq >= 0 && !vis[nq] && g[nq] > umbral) { vis[nq] = 1; cola[b++] = nq; }
      }
      if (!mejor || cnt > mejor.cnt) mejor = { cnt, x0, x1, y0, y1 };
    }
    if (!mejor) return null;
    const area = (mejor.x1 - mejor.x0 + 1) * (mejor.y1 - mejor.y0 + 1) / n;
    if (area < 0.08 || area > 0.9) return null;   // sin un papel claramente distinto del fondo no se recorta
    const pad = 0.03, sx = Math.max(0, Math.floor((mejor.x0 / w - pad) * img.width)), sy = Math.max(0, Math.floor((mejor.y0 / h - pad) * img.height));
    const ex = Math.min(img.width, Math.ceil(((mejor.x1 + 1) / w + pad) * img.width)), ey = Math.min(img.height, Math.ceil(((mejor.y1 + 1) / h + pad) * img.height));
    const r = document.createElement('canvas'); r.width = ex - sx; r.height = ey - sy; r.getContext('2d').drawImage(img, sx, sy, r.width, r.height, 0, 0, r.width, r.height);
    return r;
  }
  // Corrige la iluminación despareja (sombras, texto térmico más claro arriba): cada punto se compara con el brillo del papel a su alrededor, así el texto tenue y el fuerte quedan igual de oscuros.
  function canvasPlano(img, lado) {
    const k = Math.min(2.5, lado / Math.max(img.width, img.height)), w = Math.round(img.width * k), h = Math.round(img.height * k);   // aquí sí se amplía: las letras chicas se leen mejor más grandes
    const c = document.createElement('canvas'); c.width = w; c.height = h; const x = c.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(img, 0, 0, w, h);
    const d = x.getImageData(0, 0, w, h), p = d.data;
    for (let i = 0; i < p.length; i += 4) p[i] = p[i + 1] = p[i + 2] = 0.3 * p[i] + 0.59 * p[i + 1] + 0.11 * p[i + 2];
    const s = document.createElement('canvas'); s.width = Math.max(8, Math.round(w / 24)); s.height = Math.max(8, Math.round(h / 24));
    const sx = s.getContext('2d'); sx.imageSmoothingQuality = 'high'; sx.drawImage(c, 0, 0, s.width, s.height);
    const b = document.createElement('canvas'); b.width = w; b.height = h; const bx = b.getContext('2d'); bx.imageSmoothingQuality = 'high'; bx.drawImage(s, 0, 0, w, h);
    const bg = bx.getImageData(0, 0, w, h).data;
    for (let i = 0; i < p.length; i += 4) { const r = p[i] / Math.max(30, bg[i]), v = Math.max(0, Math.min(255, (r - 0.45) / 0.5 * 255)); p[i] = p[i + 1] = p[i + 2] = v; }
    x.putImageData(d, 0, 0); return c;
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

  // Número del comprobante de las empresas de transporte (no son facturas SUNAT con serie F/B): cada una imprime el suyo.
  // El OCR confunde O/0, S/5 y B/8, por eso los patrones aceptan ambas lecturas.
  // la serie de Jatsa es siempre 0029; si el OCR la lee con un dígito distinto (por ejemplo 0079), se corrige
  const SERIES_JATSA = ['0029'];
  const jatsaSerie = s => { const d = o => [...s].filter((c, i) => c !== o[i]).length; return SERIES_JATSA.includes(s) ? s : (SERIES_JATSA.find(o => d(o) <= 1) || s); };
  const DOC_TRANSPORTE = [
    [/G[E3]\s?[O0]?2\s*[-–—]\s*(\d{8})(?!\d)/, m => 'GE02-' + m[1]],                                   // Logística Antezana: guía de envío
    [/PEDIDO\s*[:;.]?\s*P\s?[O0]\s?5\s?2\s*[-–—]\s*(\d{8})(?!\d)/, m => 'P052-' + m[1]],               // Movil Bus: N° de pedido
    [/(?<![A-Z0-9])V\s?[O0]\s?1\s?5\s*[-–—]\s*(\d{5})(?!\d)/, m => 'V015-' + m[1]],                    // Logistics Empresarial: guía de remisión transportista
    [/(?<![A-Z0-9])[O0]\s?[S5]\s*[:;.]?\s*(0\s?0\s?\d\s?\d)\s*[-–—]\s*(\d{8})(?!\d)/, m => jatsaSerie(m[1].replace(/\s/g, '')) + '-' + m[2]],   // Jatsa Cargo: OS 0029-xxxxxxxx
    [/(?<![A-Z0-9])[O0][S5]\s?(0\s?[12])\s*[-–—]\s*(\d{8})(?!\d)/, m => 'OS' + m[1].replace(/\s/g, '') + '-' + m[2]],   // Grupo / Express Hugamor: orden de servicio
    [/GUIA\s+DE\s+CONTROL[\s\S]{0,60}?(?<!\d)(\d{3})\s*[-–—]\s*(\d{6})(?!\d)/, m => m[1] + '-' + m[2]]   // Oltursa: guía de control
  ];
  function docTransporte(T) {
    const x = norm(T);
    for (const [re, f] of DOC_TRANSPORTE) { const m = x.match(re); if (m) return f(m); }
    return '';
  }

  // Importe escrito en letras ("SON: CUARENTA Y SEIS CON 50/100 SOLES"): sirve para comprobar el importe leído, porque el OCR confunde 0 y 1 en los números pero casi no falla en las palabras.
  const PAL = { CERO: 0, UN: 1, UNO: 1, UNA: 1, DOS: 2, TRES: 3, CUATRO: 4, CINCO: 5, SEIS: 6, SIETE: 7, OCHO: 8, NUEVE: 9, DIEZ: 10, ONCE: 11, DOCE: 12, TRECE: 13, CATORCE: 14, QUINCE: 15, DIECISEIS: 16, DIECISIETE: 17, DIECIOCHO: 18, DIECINUEVE: 19, VEINTE: 20, VEINTIUN: 21, VEINTIUNO: 21, VEINTIDOS: 22, VEINTITRES: 23, VEINTICUATRO: 24, VEINTICINCO: 25, VEINTISEIS: 26, VEINTISIETE: 27, VEINTIOCHO: 28, VEINTINUEVE: 29, TREINTA: 30, CUARENTA: 40, CINCUENTA: 50, SESENTA: 60, SETENTA: 70, OCHENTA: 80, NOVENTA: 90, CIEN: 100, CIENTO: 100, DOSCIENTOS: 200, TRESCIENTOS: 300, CUATROCIENTOS: 400, QUINIENTOS: 500, SEISCIENTOS: 600, SETECIENTOS: 700, OCHOCIENTOS: 800, NOVECIENTOS: 900 };
  function importeEnLetras(T) {
    const x = norm(T).replace(/\n/g, ' ').replace(/\s+/g, ' ');
    const m = x.match(/\bSON\s*[:;.\-–—]?\s*([A-Z ]{3,90}?)\s+CON\s+(\d{1,2})\s*\/\s*1\s*0\s*0?/);
    if (!m) return null;
    let total = 0, act = 0;
    for (const w of m[1].split(' ')) {
      if (!w || w === 'Y') continue;
      if (w === 'MIL') { total += (act || 1) * 1000; act = 0; continue; }
      if (!(w in PAL)) return null;
      act += PAL[w];
    }
    return Math.round((total + act + Number(m[2]) / 100) * 100) / 100;
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
    } else R.factura = impresa || docTransporte(T);

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

    // --- el total de estas guías es casi siempre un monto redondo (tarifa por bulto): si el OCR lee 40.10 y el detalle dice 40.00, prevalece el detalle
    if (!R.viaQR && R.importe != null && Math.round(R.importe * 100) % 100 !== 0) {
      const redondo = montos(T).find(v => Number.isInteger(v) && v > 0 && Math.abs(v - R.importe) > 0 && Math.abs(v - R.importe) <= 0.15);
      if (redondo != null && !/IGV|I\.G\.V/i.test(T)) { R.importe = redondo; R.aviso.push(`Importe ajustado a ${redondo.toFixed(2)} (monto del detalle): verifícalo en la foto.`); }
    }

    // --- el importe en letras confirma (o corrige) el importe leído
    const enLetras = importeEnLetras(T);
    if (enLetras != null && !R.viaQR && R.importe !== enLetras) {
      const hayEnTexto = montos(T).some(v => cerca(v, enLetras));
      if (R.importe == null || hayEnTexto) { R.importe = enLetras; R.aviso = R.aviso.filter(a => !/importe/i.test(a)); }
      else R.aviso.push(`El importe en letras (${enLetras.toFixed(2)}) no coincide con el leído: verifícalo en la foto.`);
    }

    // --- fecha (texto)
    if (!R.fecha) {
      for (const m of T.matchAll(/(\d{2})\s*[\/-]\s*(\d{2})\s*[\/-]\s*(20\d{2})|(20\d{2})\s*[\/-]\s*(\d{2})\s*[\/-]\s*(\d{2})/g)) {
        const f = m[1] ? fechaValida(m[1], m[2], m[3], ref) : fechaValida(m[6], m[5], m[4], ref);
        if (f) { R.fecha = f; break; }
      }
      if (!R.fecha && ref) {   // el OCR a veces lee mal el año (2026 como 2028 o 2076): si día y mes caen junto a la fecha del registro, se toma el año del registro
        for (const m of T.matchAll(/(?<!\d)(\d{2})\s*[\/-]\s*(\d{2})\s*[\/-]\s*(\d{4})(?!\d)/g)) {
          const f = fechaValida(m[1], m[2], Number(ref.slice(0, 4)), ref);
          if (f) { R.fecha = f; break; }
        }
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
    if (!window.Tesseract) await cargarScript('https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js');
    if (!worker) worker = await window.Tesseract.createWorker('spa', 1, {
      langPath: 'https://cdn.jsdelivr.net/gh/tesseract-ocr/tessdata_fast@main', gzip: false,   // modelo "fast": ~1 MB en vez de ~8 MB
      logger: m => progreso && m.status && progreso(m.status, m.progress) });
    return worker;
  }

  // file: imagen de la factura. ref: fecha del registro (YYYY-MM-DD). nodos: lista de nodos del catálogo.
  async function leer(file, { ref, nodos, progreso } = {}) {
    await libs();
    const img = await imagenDe(file);
    progreso && progreso('Buscando código QR…', 0);
    const qr = leerQR(img);
    const w = await obtenerWorker(progreso);
    let texto = '', R = null;
    const papel = recortarPapel(img), base = papel || img;
    const pasadas = [{ src: base, lado: 2000, contraste: false, psm: '6' }, { src: base, lado: 2600, contraste: true, psm: '4' }, { src: base, lado: 2600, plano: true, psm: '6' }];
    const mitad = (s, a, b) => { const r = document.createElement('canvas'); r.width = s.width; r.height = Math.round(s.height * (b - a)); r.getContext('2d').drawImage(s, 0, Math.round(s.height * a), s.width, r.height, 0, 0, r.width, r.height); return r; };
    pasadas.push({ src: mitad(base, 0, 0.4), lado: 2400, plano: true, psm: '6' });   // encabezado ampliado: ahí están el N° y la fecha, con letra pequeña
    if (papel) pasadas.push({ src: img, lado: 2000, contraste: false, psm: '6' });   // si el recorte no bastó, se prueba también con la foto completa
    for (const p of pasadas) {
      await w.setParameters({ tessedit_pageseg_mode: p.psm });
      const { data } = await w.recognize(p.plano ? canvasPlano(p.src, p.lado) : canvasDe(p.src, p.lado, true, p.contraste));
      texto = texto ? texto + '\n' + data.text : data.text;
      R = parsear(texto, qr, ref, nodos);
      if (R.factura && R.importe != null) break;
    }
    R.texto = texto; R.qr = qr; R.motor = 'Tesseract';
    return R;
  }

  // Versión para guardar: reduce por pasos, pasa a grises, quita el ruido de la cámara y lleva el papel a blanco y la tinta a negro (el ruido es lo que más pesa; sin él sobran bits para el texto).
  function canvasNitido(img, lado) {
    const k = Math.min(1, lado / Math.max(img.width, img.height)), w = Math.round(img.width * k), h = Math.round(img.height * k);
    let c = document.createElement('canvas'); c.width = img.width; c.height = img.height; c.getContext('2d').drawImage(img, 0, 0);
    while (c.width / 2 >= w && c.height / 2 >= h) {   // reducción a la mitad cuantas veces se pueda: conserva más detalle que un solo salto
      const m = document.createElement('canvas'); m.width = Math.round(c.width / 2); m.height = Math.round(c.height / 2);
      const mx = m.getContext('2d'); mx.imageSmoothingQuality = 'high'; mx.drawImage(c, 0, 0, m.width, m.height); c = m;
    }
    const f = document.createElement('canvas'); f.width = w; f.height = h;
    const x = f.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(c, 0, 0, w, h);
    const d = x.getImageData(0, 0, w, h), p = d.data, n = w * h, g0 = new Float32Array(n), g = new Uint8ClampedArray(n), hist = new Uint32Array(256);
    for (let i = 0, q = 0; i < p.length; i += 4, q++) g0[q] = 0.3 * p[i] + 0.59 * p[i + 1] + 0.11 * p[i + 2];
    for (let y = 0; y < h; y++) for (let xx = 0; xx < w; xx++) {   // suavizado muy leve (vecinos en cruz) para quitar el grano sin borrar las letras
      const q = y * w + xx; let s = g0[q] * 4, c = 4;
      if (xx > 0) { s += g0[q - 1] * 2; c += 2; } if (xx < w - 1) { s += g0[q + 1] * 2; c += 2; } if (y > 0) { s += g0[q - w] * 2; c += 2; } if (y < h - 1) { s += g0[q + w] * 2; c += 2; }
      const v = Math.round(s / c); g[q] = v; hist[v]++;
    }
    const pct = f => { let a = 0; for (let v = 0; v < 256; v++) { a += hist[v]; if (a >= n * f) return v; } return 255; };
    const negro = pct(0.02), blanco = Math.max(negro + 60, pct(0.92)), esc = 255 / (blanco - negro);   // niveles automáticos: papel a blanco, tinta a negro
    for (let q = 0; q < n; q++) { const v = Math.max(0, Math.min(255, (g[q] - negro) * esc)); p[q * 4] = p[q * 4 + 1] = p[q * 4 + 2] = v; }
    x.putImageData(d, 0, 0); return f;
  }
  // Comprime una imagen a <= maxBytes (grises, WebP si el navegador lo permite, si no JPEG). Prioriza la resolución: usa la mayor en la que todavía cabe con calidad aceptable, porque las letras se leen por tamaño más que por suavidad.
  async function comprimir(file, maxBytes = 40 * 1024) {
    const img = await imagenDe(file);
    const aBlob = (c, tipo, q) => new Promise(ok => c.toBlob(ok, tipo, q));
    const prueba = await aBlob(canvasDe(img, 64, false), 'image/webp', 0.5);
    const tipo = prueba && prueba.type === 'image/webp' ? 'image/webp' : 'image/jpeg';
    let ultima = null;
    for (const lado of [1500, 1400, 1300, 1200, 1100, 1000, 900, 820, 740, 670, 600, 540, 480, 420, 360, 300]) {
      const c = canvasNitido(img, lado);
      let lo = 0.22, hi = 0.92;
      const min = await aBlob(c, tipo, lo);
      if (!min || min.size > maxBytes) continue;
      let mejor = { blob: min, calidad: lo };
      for (let i = 0; i < 6; i++) {
        const q = (lo + hi) / 2, b = await aBlob(c, tipo, q);
        if (b && b.size <= maxBytes) { mejor = { blob: b, calidad: q }; lo = q; } else hi = q;
      }
      ultima = { blob: mejor.blob, tipo, ext: tipo === 'image/webp' ? 'webp' : 'jpg', ancho: c.width, alto: c.height, calidad: Math.round(mejor.calidad * 100) };
      if (mejor.calidad >= .3) break;   // resolución máxima con calidad aceptable
    }
    if (!ultima) throw new Error('No se pudo comprimir la imagen a ' + Math.round(maxBytes / 1024) + ' KB');
    return ultima;
  }

  return { leer, parsear, rucValido, comprimir, recortarPapel };
})();
