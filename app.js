const $ = s => document.querySelector(s);
const ICO = {
  clip: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.5l-8.6 8.6a5.5 5.5 0 01-7.8-7.8l9-9a3.7 3.7 0 015.2 5.2l-9 9a1.8 1.8 0 01-2.6-2.6l8.3-8.3"/></svg>',
  file: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8zM14 3v5h5M9 13h6M9 17h6"/></svg>',
  refresh: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 11a8 8 0 00-14.9-3M4 4v4h4M4 13a8 8 0 0014.9 3M20 20v-4h-4"/></svg>',
  x: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  check: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
  prev: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 5l-7 7 7 7"/></svg>',
  img: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="1.6"/><path d="M21 16l-5-5-8 8"/></svg>',
  edit: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20h4L19 9l-4-4L4 16zM13.5 6.5l4 4"/></svg>',
  next: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 5l7 7-7 7"/></svg>'
};
const flecha = sube => `<svg class="ar" viewBox="0 0 10 10" width="9" height="9" fill="currentColor" aria-hidden="true"><path d="${sube ? 'M5 1.5l4 7H1z' : 'M5 8.5l-4-7h8z'}"/></svg>`;
// librería de Excel: se descarga solo al importar o exportar
let xlsxPromesa;
const xlsxLib = () => window.XLSX ? Promise.resolve() : (xlsxPromesa ||= new Promise((ok, ko) => { const t = document.createElement('script'); t.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; t.onload = ok; t.onerror = () => { xlsxPromesa = null; ko(new Error('No se pudo cargar la librería de Excel. Verificar la conexión a internet.')); }; document.head.appendChild(t); }));
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g,'').toUpperCase().replace(/\s+/g,' ').trim();
const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
const PAGOS = ['CONTADO','CREDITO'];
const MOTIVOS = ['DESPACHO', 'RECOJO'];
const motTag = m => m === 'RECOJO' ? '<span class="tb rec">Recojo</span>' : '<span class="tb">Despacho</span>';
const num = v => { const n = Number(String(v ?? '').replace(/,/g,'')); return Number.isFinite(n) ? n : null; };
const today = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };   // fecha local (no UTC)
const fdmy = f => f ? String(f).slice(0, 10).split('-').reverse().join('/') : '';
const fdt = f => f ? fdmy(f) + ' ' + String(f).slice(11, 16) : '';

function confirmar(titulo, mensaje, ok = 'Aceptar', peligro = false) {
  return new Promise(res => {
    const d = $('#cfDlg'); $('#cfT').textContent = titulo; $('#cfM').textContent = mensaje || ''; $('#cfM').classList.toggle('hide', !mensaje);
    const b = $('#cfOk'); b.textContent = ok; b.classList.toggle('red', !!peligro);
    d.onclose = () => res(d.returnValue === 'si'); d.returnValue = 'no'; d.showModal();
  });
}
document.addEventListener('change', e => { if (e.target.dataset && 'v' in e.target.dataset) e.target.dataset.v = e.target.value; });
function setPaso(n) {
  const h = $('#cPasos');
  if (h.children.length !== 5) {
    h.innerHTML = ['Archivo', 'Revisar', 'Listo'].map((t, i) => (i ? '<span class="ln"></span>' : '') + `<span class="ps"><span class="pn"></span><span class="pt">${t}</span></span>`).join('');
  }
  [...h.children].forEach((el, q) => {
    if (q % 2) { el.classList.toggle('hecho', q / 2 + 0.5 < n); return; }
    const k = q / 2 + 1; el.className = 'ps ' + (k < n ? 'hecho' : k === n ? 'on' : '');
    const pn = el.firstElementChild, txt = k < n ? 'c' : String(k); if (pn.dataset.s !== txt) { pn.dataset.s = txt; pn.innerHTML = k < n ? ICO.check : k; }
  });
}
function toast(text, type) {
  const t = document.createElement('div'); t.className = 'toast ' + (type === 'err' ? 'err' : 'ok');
  t.innerHTML = `<i>${type === 'err' ? ICO.x : ICO.check}</i><span>${esc(text)}</span>`; $('#toasts').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 300); }, type === 'err' ? 6500 : 3200);
}
function flash(el, text, type) {
  if (type === 'ok') { if (el) el.className = 'msg'; return toast(text, 'ok'); }
  if (el) { el.textContent = text; el.className = 'msg ' + type; }
  toast(text, 'err');
}

const cfg = window.APP_CONFIG || {};
if (!cfg.SUPABASE_URL || !cfg.SUPABASE_ANON_KEY) { $('#setup').classList.remove('hide'); throw new Error('config'); }
const sb = supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);

const zonaEnv = e => e.zona || '(sin zona)';
let agBase = [];   // agencias ya conocidas (para sugerir al escribir)
let previa = [];            // filas de la carga en curso
let hrnCarga = null;        // detalle HRN encontrado en el mismo Excel
let registros = [];
const rtMios = new Map();   // registros que acabo de modificar yo (el eco en tiempo real se ignora)
const marcarMio = id => rtMios.set(id, Date.now()), esMio = id => Date.now() - (rtMios.get(id) || 0) < 4000;

// ---------- acceso y pantalla de carga: red de distribución sobre el mapa real ----------
// destino, nombre visible, lado de la etiqueta, vehículo (c camión / a avión / null) y sentido (1 despacho, -1 devolución)
const AC_RED = [['PIURA', 'Piura', 'l', 'a', -1], ['CHICLAYO', 'Chiclayo', 'l', null, 1], ['TRUJILLO', 'Trujillo', 'l', 'c', 1], ['SANTA', 'Chimbote', 'l', null, 1], ['CAJAMARCA', 'Cajamarca', 'r', 'c', -1], ['MAYNAS', 'Iquitos', 'r', 'a', 1], ['SAN MARTIN', 'Tarapoto', 'r', null, 1], ['CORONEL PORTILLO', 'Pucallpa', 'r', null, 1], ['HUANCAYO', 'Huancayo', 'r', 'c', 1], ['HUAMANGA', 'Ayacucho', 'r', null, 1], ['ICA', 'Ica', 'l', 'c', 1], ['CUSCO', 'Cusco', 'r', 'a', 1], ['AREQUIPA', 'Arequipa', 'l', 'a', -1], ['PUNO', 'Puno', 'r', null, 1], ['TACNA', 'Tacna', 'l', null, 1]];
const AC_FORMA = {
  camion: 'M-8 -3.6h9.6v7.2h-9.6zM2.2 -3.6h3.4l2.6 3.2v4h-6z',
  avion: 'M9.5 0L3.5 -1.4L-.5 -7.6L-3 -7.6L-1 -1.4L-6 -1.4L-7.6 -4.2L-9.5 -4.2L-8.2 0L-9.5 4.2L-7.6 4.2L-6 1.4L-1 1.4L-3 7.6L-.5 7.6L3.5 1.4Z'
};
const AC = { rutas: [], cams: [], vel: .7, dib: false, on: false, w: 0 };
const acQuieto = !!(window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches);
const acSvg = (t, a, p) => { const e = document.createElementNS('http://www.w3.org/2000/svg', t); for (const k in a) e.setAttribute(k, a[k]); if (p) p.appendChild(e); return e; };
function acDibujar() {
  if (AC.dib || !kaGeo) return;
  const svg = $('#lSvg'); AC.dib = true;
  const defs = acSvg('defs', {}, svg);
  acSvg('linearGradient', { id: 'acTierra', x1: 0, y1: 0, x2: 0, y2: 1 }, defs).innerHTML = '<stop offset="0" stop-color="#eaf2fd"/><stop offset="1" stop-color="#d9e6f8"/>';
  acSvg('radialGradient', { id: 'acVelo', cx: 310, cy: 450, r: 520, gradientUnits: 'userSpaceOnUse' }, defs).innerHTML = '<stop offset="0" stop-color="#fff"/><stop offset=".55" stop-color="#fff" stop-opacity=".7"/><stop offset="1" stop-color="#fff" stop-opacity="0"/>';
  acSvg('mask', { id: 'acMasc', maskUnits: 'userSpaceOnUse', x: -500, y: -300, width: 1620, height: 1500 }, defs).innerHTML = '<rect x="-500" y="-300" width="1620" height="1500" fill="url(#acVelo)"/>';
  acSvg('filter', { id: 'acSombra', x: '-10%', y: '-10%', width: '120%', height: '125%' }, defs).innerHTML = '<feDropShadow dx="0" dy="10" stdDeviation="12" flood-color="#0b3d82" flood-opacity=".16"/>';
  // cuadrícula de coordenadas que se desvanece hacia los bordes
  const gG = acSvg('g', { mask: 'url(#acMasc)', stroke: '#0b3d82', 'stroke-opacity': .17, 'stroke-width': .8, fill: 'none' }, svg);
  for (let y = -150; y <= 1100; y += 100) acSvg('path', { d: `M-500 ${y}H1120` }, gG);
  for (let x = -400; x <= 1000; x += 100) acSvg('path', { d: `M${x} -300V1200` }, gG);
  const gP = acSvg('g', { class: 'ac-som' }, svg), gPt = acSvg('g', {}, svg), gR = acSvg('g', {}, svg), gT = acSvg('g', {}, svg), gN = acSvg('g', {}, svg);
  const red = new Set(AC_RED.map(r => r[0]));
  kaGeo.p.forEach(p => {
    const rojo = red.has(norm(p[0]));
    acSvg('path', { d: p[4], class: 'ac-prov' + (rojo ? ' d' : '') }, gP);
    if (!rojo && norm(p[0]) !== 'LIMA') acSvg('circle', { cx: p[2], cy: p[3], r: 1.7, class: 'ac-pto' }, gPt);   // el resto de las provincias, apenas insinuadas
  });
  const pos = n => { const p = kaProv.get(norm(n)); return p ? [p[2], p[3]] : null; };
  const O = pos('LIMA') || [300, 600];
  AC_RED.forEach((r, i) => {
    const d = pos(r[0]); if (!d) return;
    const mx = (O[0] + d[0]) / 2, my = (O[1] + d[1]) / 2, dx = d[0] - O[0], dy = d[1] - O[1], k = .16 * (i % 2 ? 1 : -1), dv = r[4] < 0;
    const path = acSvg('path', { d: `M${O[0]} ${O[1]} Q${mx - dy * k} ${my + dx * k} ${d[0]} ${d[1]}`, class: 'ac-ruta' + (dv ? ' dv' : '') }, gR);
    const g = acSvg('g', { class: 'ac-n' + (dv ? ' dv' : '') }, gN);
    acSvg('circle', { class: 'p', cx: d[0], cy: d[1], r: 6, style: `animation-delay:${(i * .37) % 3}s` }, g); acSvg('circle', { class: 'c', cx: d[0], cy: d[1], r: 5 }, g);
    const w = r[1].length * 6.1 + 16, lx = r[2] === 'r' ? d[0] + 11 : d[0] - 11 - w, ch = acSvg('g', { class: 'ac-chip' }, g);
    acSvg('rect', { x: lx, y: d[1] - 10, width: w, height: 20, rx: 10 }, ch); acSvg('text', { x: lx + w / 2, y: d[1] + 3.6, 'text-anchor': 'middle' }, ch).textContent = r[1];
    AC.rutas.push({ path, g });
    if (r[3]) {   // solo algunos destinos llevan vehículo, para no recargar el mapa
      const e = acSvg('g', { class: 'ac-cam' + (dv ? ' dv' : '') }, gT), est = acSvg('path', { class: 'est' }, e), pt = acSvg('g', {}, e), aire = r[3] === 'a';
      acSvg('path', { d: aire ? AC_FORMA.avion : AC_FORMA.camion, class: aire ? 'av' : 'cm' }, pt);
      if (!aire) { acSvg('circle', { cx: -4.6, cy: 4, r: 1.5, class: 'rd' }, pt); acSvg('circle', { cx: 3.8, cy: 4, r: 1.5, class: 'rd' }, pt); }
      AC.cams.push({ path, e, est, pt, dir: r[4], len: path.getTotalLength(), t: (i * .21 + (dv ? .5 : 0)) % 1, v: aire ? .0026 : .0016 });
    }
  });
  const o = acSvg('g', { class: 'ac-o' }, gN);
  acSvg('circle', { class: 'p', cx: O[0], cy: O[1], r: 8 }, o); acSvg('circle', { class: 'o', cx: O[0], cy: O[1], r: 8 }, o);
  acSvg('rect', { x: O[0] - 86, y: O[1] - 12, width: 68, height: 24, rx: 12 }, o); acSvg('text', { x: O[0] - 52, y: O[1] + 4, 'text-anchor': 'middle' }, o).textContent = 'CTD Lima';
  acConectar(AC.w);
}
function acConectar(w) {   // w: avance de 0 a 1; enciende ese porcentaje de nodos y rutas
  const n = Math.round(AC.rutas.length * w);
  AC.rutas.forEach((r, i) => { r.path.classList.toggle('on', i < n); r.g.classList.toggle('on', i < n); });
}
function acLoop() {
  if (!AC.on) return;
  AC.cams.forEach(c => {
    c.t += c.v * AC.vel * c.dir; if (c.t > 1) c.t -= 1; else if (c.t < 0) c.t += 1;
    const p = c.path.getPointAtLength(c.len * c.t), p2 = c.path.getPointAtLength(c.len * Math.max(0, Math.min(1, c.t + .01 * c.dir)));
    const ang = Math.atan2(p2.y - p.y, p2.x - p.x) * 180 / Math.PI;   // el vehículo mira hacia donde avanza; si va hacia la izquierda se voltea para no quedar al revés
    c.pt.setAttribute('transform', `translate(${p.x} ${p.y}) rotate(${ang}) scale(2.1 ${Math.abs(ang) > 90 ? -2.1 : 2.1})`);
    const a = Math.max(0, Math.min(1, c.t - .08 * c.dir)); let d = '';
    for (let k = 0; k <= 4; k++) { const q = c.path.getPointAtLength(c.len * (a + (c.t - a) * k / 4)); d += (k ? 'L' : 'M') + q.x + ' ' + q.y; }
    c.est.setAttribute('d', d);
  });
  requestAnimationFrame(acLoop);
}
function acMostrar(modo) {
  const ac = $('#acceso'), carga = modo === 'carga';
  ac.classList.remove('hide', 'sale'); ac.classList.toggle('carga', carga);
  $('#login').classList.toggle('hide', carga); $('#cargaIni').classList.toggle('hide', !carga);
  AC.vel = acQuieto ? 0 : carga ? 2.2 : .7; AC.w = carga ? 0 : 1;
  acDibujar(); acConectar(AC.w);
  if (carga) { $('#acBar').style.width = '0'; ['acS1', 'acS2', 'acS3'].forEach(i => { $('#' + i).className = ''; }); }
  if (!AC.on) { AC.on = true; acLoop(); }
}
function acCarga(w, txt, paso) {   // avance real de la carga inicial
  AC.w = w; $('#acBar').style.width = w * 100 + '%'; $('#acPc').textContent = Math.round(w * 100) + ' %'; $('#acMsg').textContent = txt; acConectar(w);
  ['acS1', 'acS2', 'acS3'].forEach((id, k) => { $('#' + id).className = !paso ? 'ok' : k + 1 < paso ? 'ok' : k + 1 === paso ? 'on' : ''; });
}
function acOcultar() {
  const ac = $('#acceso'); ac.classList.add('sale');
  setTimeout(() => { ac.classList.add('hide'); ac.classList.remove('sale'); AC.on = false; }, 520);
}
let kpisPend = null;   // consulta del Dashboard en curso (la pantalla de carga espera a que termine)

// ---------- sesión ----------
async function iniciar() {
  // con REQUIRE_LOGIN en false la app abre sin sesión; agregar ?login a la dirección permite probar el ingreso
  if (cfg.REQUIRE_LOGIN === false && !/[?&]login\b/.test(location.search)) { $('#logout').classList.add('hide'); return mostrarApp({ email: '' }); }
  kaCargarGeo();   // el mapa del acceso es el mismo del Dashboard
  const { data } = await sb.auth.getSession();
  if (data.session) mostrarApp(data.session.user); else acMostrar('login');
}
// ingreso con correo y contraseña; los usuarios los crea el administrador en Supabase
$('#lBtn').onclick = async () => {
  const email = $('#lEmail').value.trim().toLowerCase(), password = $('#lPass').value;
  if (!email || !password) return flash($('#lMsg'), 'Ingresar el correo y la contraseña.', 'err');
  $('#lBtn').disabled = true; $('#lBtn').textContent = 'Verificando…';
  const { data, error } = await sb.auth.signInWithPassword({ email, password });
  $('#lBtn').disabled = false; $('#lBtn').textContent = 'Ingresar';
  if (error) { const t = $('#lCard'); t.classList.remove('tiembla'); void t.offsetWidth; t.classList.add('tiembla'); return flash($('#lMsg'), 'Correo o contraseña incorrectos.', 'err'); }
  mostrarApp(data.user);
};
$('#lEmail').addEventListener('keydown', e => { if (e.key === 'Enter') $('#lPass').focus(); });
$('#lPass').addEventListener('keydown', e => { if (e.key === 'Enter') $('#lBtn').click(); });
$('#lCard').addEventListener('submit', e => e.preventDefault());
$('#lOjo').onclick = () => { const p = $('#lPass'); p.type = p.type === 'password' ? 'text' : 'password'; };
$('#logout').onclick = async () => { await sb.auth.signOut(); location.reload(); };
// si la sesión se cierra desde otra pestaña, vence o el usuario es eliminado, la aplicación vuelve al acceso
sb.auth.onAuthStateChange(ev => { if (ev === 'SIGNED_OUT' && !$('#app').classList.contains('hide')) location.reload(); });
// una falla no prevista no debe pasar inadvertida: se avisa una sola vez cada 8 segundos
let ultAviso = 0;
const avisoGeneral = () => { if ($('#app').classList.contains('hide') || Date.now() - ultAviso < 8000) return; ultAviso = Date.now(); toast('No se pudo completar la operación. Actualizar la página e intentar nuevamente.', 'err'); };
window.addEventListener('unhandledrejection', e => { if (e.reason && e.reason.name === 'AbortError') return; console.error(e.reason); avisoGeneral(); });
window.addEventListener('error', e => { if (/ResizeObserver/.test(e.message || '')) return; console.error(e.error || e.message); avisoGeneral(); });

async function mostrarApp(user) {
  const tIni = Date.now(); acMostrar('carga'); acCarga(.1, 'Verificando sesión', 1);
  $('#who').textContent = user.email;
  $('#app').classList.remove('hide');
  $('#kHasta').value = today(); $('#kMes').value = today().slice(0, 7); rangoDeMes();
  $('#rDesde').value = today().slice(0, 8) + '01'; $('#rHasta').value = today();
  sb.from('lista_agencias').select('agencia').then(({ data }) => { agBase = (data || []).map(x => x.agencia); listaAgencias(); });
  const pReg = buscarRegistros(); iniciarTiempoReal(); acCarga(.3, 'Cargando registros', 2);
  const h0 = (location.hash || '').slice(1); let t0 = h0.split('?')[0]; if (t0 === 'hrn') t0 = 'detalle'; if (!TITULOS[t0]) t0 = 'kpis';   // sin enlace, se abre el Dashboard
  if (t0 === 'kpis' && h0.includes('?')) { poblarFiltros(); kAplicar(h0.split('?')[1]); }
  irA(t0);
  try { await pReg; acCarga(.7, 'Calculando indicadores', 3); if (kpisPend) await Promise.race([kpisPend, new Promise(r => setTimeout(r, 10000))]); } catch (e) {}
  acCarga(1, 'Listo', 0);
  await new Promise(r => setTimeout(r, Math.max(250, 1100 - (Date.now() - tIni))));
  acOcultar();
}
window.addEventListener('hashchange', () => { const t = location.hash.slice(1).split('?')[0].replace(/^hrn$/, 'detalle'); if (TITULOS[t]) irA(t); });

const TITULOS = { cargar: 'Cargar despacho', registros: 'Registros', kpis: 'Dashboard', liquidaciones: 'Liquidaciones', detalle: 'Detalle' };
const irA = tab => document.querySelector(`nav button[data-tab="${tab}"]`).click();
const NAVB = 'nav button, .tabbar button[data-tab]';
document.querySelectorAll(NAVB).forEach(b => b.onclick = () => {
  document.querySelectorAll(NAVB).forEach(x => x.classList.toggle('on', x.dataset.tab === b.dataset.tab));
  ['cargar','registros','liquidaciones','kpis','detalle'].forEach(t => $('#tab-' + t).classList.toggle('hide', t !== b.dataset.tab));
  document.title = TITULOS[b.dataset.tab] + ' – Red Troncal';
  history.replaceState(null, '', '#' + b.dataset.tab); try { localStorage.setItem('tab', b.dataset.tab); } catch (e) {}
  scrollTo(0, 0);
  if (b.dataset.tab === 'liquidaciones') lqAbrir();
  if (b.dataset.tab === 'detalle' && !hrnCargado) buscarHrn();
  if (b.dataset.tab === 'kpis') kpisPend = calcularKpis().catch(() => {});
});
function conexion(ok) { $('#estado').className = 'estado' + (ok ? '' : ' mal'); $('#estado span').textContent = ok ? '' : 'Sin conexión con la base de datos'; }

// ---------- carga ----------
$('#cPlantilla').onclick = async () => {
  await xlsxLib();
  const wb = XLSX.utils.book_new(), cab = ['FECHA', 'NODO', 'PEDIDOS', 'BULTO', 'CAJAS', 'TRANSPORTE', 'PLACA', 'AGENCIA', 'MODO DE PAGO', 'ZONA'];
  [['DESPACHO', cab], ['RECOJOS', cab], ['DETALLE', ['FECHA', ...HRN_COLS.map(c => c[1])]], ['DETALLE RECOJOS', ['FECHA', ...RC_COLS.map(c => c[1])]]].forEach(([n, h]) => {
    const ws = XLSX.utils.aoa_to_sheet([h]); ws['!cols'] = h.map(x => ({ wch: Math.max(12, x.length + 2) })); XLSX.utils.book_append_sheet(wb, ws, n);
  });
  XLSX.writeFile(wb, 'plantilla_despacho.xlsx');
};

// fecha de una celda del Excel: número de serie, fecha, "dd/mm/aaaa" o "aaaa-mm-dd"; '' si no se reconoce
function fechaDe(v) {
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
  if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  const t = String(v ?? '').trim(); let m;
  if ((m = t.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/))) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  if ((m = t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/))) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return '';
}
// lee la primera tabla con encabezados NODO y CAJAS de una hoja; null si la hoja no la tiene, [] si está vacía
function tablaDeHoja(ws) {
  const filas = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  for (let i = 0; i < filas.length; i++) {
    const h = filas[i].map(norm);
    const iN = h.findIndex(x => x === 'NODO' || x === 'NODOS');
    const iC = h.findIndex(x => x.includes('CAJA'));
    if (iN < 0 || iC < 0) continue;
    const iP = h.findIndex(x => x.includes('PEDIDO')), iB = h.findIndex(x => x.includes('BULTO')), iT = h.findIndex(x => x.startsWith('TRANSPORT'));
    const iF = h.findIndex(x => ['FECHA', 'FECHA DESPACHO', 'FECHA DE DESPACHO', 'FECHA RECOJO', 'FECHA DE RECOJO'].includes(x)), iPl = h.findIndex(x => x.startsWith('PLACA')), iAg = h.findIndex(x => x.startsWith('AGENCIA')), iPg = h.findIndex(x => x.startsWith('PAGO') || x.includes('MODO DE PAGO') || x === 'OBSERVACION'), iZn = h.findIndex(x => x === 'ZONA');
    const pagoDe = v => { const o = norm(v); return o.includes('CONTADO') ? 'CONTADO' : o.includes('CREDITO') ? 'CREDITO' : ''; }, txt = v => String(v ?? '').trim();
    const out = [];
    for (let j = i + 1; j < filas.length; j++) {
      const nodo = norm(filas[j][iN]);
      if (!nodo || nodo.startsWith('TOTAL')) break;
      out.push({ fecha: iF >= 0 ? fechaDe(filas[j][iF]) : '', placa: iPl >= 0 ? txt(filas[j][iPl]).toUpperCase() : '', nodo, pedidos: num(filas[j][iP]), bultos: num(filas[j][iB]), cajas: num(filas[j][iC]), transporte: norm(filas[j][iT]),
        agencia: iAg >= 0 ? txt(filas[j][iAg]) : '', pago: iPg >= 0 ? pagoDe(filas[j][iPg]) : '', zona: iZn >= 0 ? txt(filas[j][iZn]).toUpperCase() : '' });
    }
    return out;
  }
  return null;
}
// hojas DESPACHO y RECOJOS; si el archivo es de los anteriores (hoja RESUMEN u otra), todo se toma como despacho
function leerTabla(wb) {
  const tipo = n => ({ DESPACHO: 'DESPACHO', RECOJO: 'RECOJO', RECOJOS: 'RECOJO' })[norm(n)];
  const propias = wb.SheetNames.filter(tipo), out = [];
  if (propias.length) {
    propias.forEach(n => (tablaDeHoja(wb.Sheets[n]) || []).forEach(r => out.push({ ...r, motivo: tipo(n) })));
    return out.length ? out : null;
  }
  const hojas = [...wb.SheetNames].sort((a, b) => (norm(b) === 'RESUMEN') - (norm(a) === 'RESUMEN'));
  for (const n of hojas) { const f = tablaDeHoja(wb.Sheets[n]); if (f && f.length) return f.map(r => ({ ...r, motivo: 'DESPACHO' })); }
  return null;
}
// detalle del gasto automático: "ENVIO DE PEDIDOS X" o "RECOJO DE PEDIDOS X" (se puede editar después en Registros)
const detalleGasto = r => (r.motivo === 'RECOJO' ? 'RECOJO' : 'ENVIO') + ' DE PEDIDOS ' + r.nodo;

let cArchivo = null, cFiltro = '', cSeq = 0, recDetCarga = null;
const cFechas = () => [...new Set(previa.map(r => r.fecha).filter(Boolean))].sort();
const cFechasTodas = () => [...new Set([...previa.map(r => r.fecha), ...(hrnCarga || []).map(r => r.fecha_reporte), ...(recDetCarga || []).map(r => r.fecha_reporte)].filter(Boolean))].sort();
const cSinFecha = () => previa.filter(r => !r.fecha).length + (hrnCarga || []).filter(r => !r.fecha_reporte).length + (recDetCarga || []).filter(r => !r.fecha_reporte).length;
const cHojas = () => ({ despacho: previa.some(r => r.motivo === 'DESPACHO'), recojos: previa.some(r => r.motivo === 'RECOJO'), detalle: !!hrnCarga, recdet: !!recDetCarga });
function cReiniciar() {
  cSeq++; previa = []; hrnCarga = null; recDetCarga = null; cArchivo = null; cFiltro = ''; $('#cFile').value = ''; $('#cHojas').classList.add('hide'); $('#cMsg').className = 'msg';
  ['#cPrev', '#cOk', '#cArch', '#cHojas'].forEach(q => $(q).classList.add('hide')); ['#cS1', '#cDrop', '#cAyuda'].forEach(q => $(q).classList.remove('hide')); $('#cGrid').classList.remove('uno');
  $('#cBarra').hidden = true; $('#tab-cargar').classList.remove('conbarra'); setPaso(1);
}
const cEspera = ms => new Promise(r => setTimeout(r, ms));
const cReducir = () => matchMedia('(prefers-reduced-motion:reduce)').matches;
function cContar(el, hasta, ms) {
  if (cReducir() || hasta < 2) { el.textContent = fmtN(hasta); return; }
  const t0 = performance.now(), paso = t => { const k = Math.min(1, (t - t0) / ms); el.textContent = fmtN(Math.round(hasta * (1 - Math.pow(1 - k, 3)))); if (k < 1) requestAnimationFrame(paso); };
  requestAnimationFrame(paso);
}
// lectura animada: cada hoja se revisa, se marca y muestra cuántas filas trajo; devuelve false si se canceló
async function cLeer(seq) {
  if (cReducir()) return true;
  const n = { despacho: previa.filter(r => r.motivo === 'DESPACHO').length, recojos: previa.filter(r => r.motivo === 'RECOJO').length, detalle: hrnCarga ? hrnCarga.length : 0, recdet: recDetCarga ? recDetCarga.length : 0 };
  const hojas = [['despacho', 'DESPACHO', 'Nodos despachados'], ['recojos', 'RECOJOS', 'Nodos recogidos'], ['detalle', 'DETALLE', 'Bultos con su fecha'], ['recdet', 'DETALLE RECOJOS', 'Pedidos recogidos']];
  $('#cDrop').classList.add('hide'); $('#cAyuda').classList.add('hide'); $('#cGrid').classList.add('uno'); $('#cArch').classList.remove('hide');
  $('#cArchN').textContent = cArchivo.nombre; $('#cArchS').textContent = `${cArchivo.kb} KB, leyendo las hojas`;
  $('#cHojas').innerHTML = hojas.map(([k, t, d]) => `<div class="chj" data-k="${k}"><span class="pt">${ICO.check}</span><div class="t"><b>${t}</b><span>${d}</span></div><div class="num"><b></b></div></div>`).join('');
  $('#cHojas').classList.remove('hide'); await cEspera(120);
  for (const [k] of hojas) {
    const el = document.querySelector(`#cHojas [data-k=${k}]`); if (seq !== cSeq) return false;
    if (!n[k]) { el.classList.add('sin'); el.querySelector('.num').textContent = 'Sin filas'; await cEspera(160); continue; }
    el.classList.add('busca'); await cEspera(520); if (seq !== cSeq) return false;
    el.classList.remove('busca'); el.classList.add('listo'); el.querySelector('.num').innerHTML = '<b></b><small>filas</small>'; cContar(el.querySelector('.num b'), n[k], 650); await cEspera(380);
  }
  if (seq !== cSeq) return false;
  await cEspera(250); $('#cHojas').classList.add('hide'); return seq === cSeq;
}
async function leerArchivo(f) {
  try {
    hrnCarga = null; recDetCarga = null; await xlsxLib(); const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
    const filas = leerTabla(wb); hrnCarga = leerHrn(wb); recDetCarga = leerRecojoDet(wb);
    if (!filas && !hrnCarga && !recDetCarga) return flash($('#cMsg'), 'No se encontraron filas en las hojas DESPACHO, RECOJOS, DETALLE o DETALLE RECOJOS. Descargar la plantilla y completar esas hojas.', 'err');
    previa = (filas || []).map(r => ({ ...r, agencia: r.agencia || '', modo_pago: r.pago || '', zona: r.zona || '', detalle_gasto: detalleGasto(r) }));
    let rellenas = 0;   // la zona vacía se completa con la última que tuvo ese nodo en los registros guardados
    try {
      const conocidas = new Map();
      (await traerTodo(() => sb.from('envios').select('nodo,zona').not('zona', 'is', null).order('fecha', { ascending: false }).order('id'))).forEach(x => { const k = norm(x.nodo); if (!conocidas.has(k)) conocidas.set(k, x.zona); });
      previa.forEach(r => { if (!r.zona && conocidas.has(norm(r.nodo))) { r.zona = conocidas.get(norm(r.nodo)); rellenas++; } });
    } catch (e) { /* sin conexión o sin la columna: la zona se escribe a mano */ }
    cArchivo = { nombre: f.name, kb: Math.max(1, Math.round(f.size / 1024)) }; cFiltro = ''; listaAgencias();
    const seq = ++cSeq; if (!await cLeer(seq)) return;
    $('#cMsg').className = 'msg'; setPaso(2); pintarPrevia(true);
    if (rellenas) toast(`Zona completada en ${rellenas} ${rellenas === 1 ? 'fila' : 'filas'} a partir de registros anteriores`);
  } catch (err) { flash($('#cMsg'), 'No se pudo leer el archivo: ' + err.message, 'err'); }
}
$('#cFile').onchange = e => { const f = e.target.files[0]; if (f) leerArchivo(f); };
{
  const dz = $('#cDrop');
  dz.onclick = () => $('#cFile').click(); dz.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('#cFile').click(); } };
  dz.ondragover = e => { e.preventDefault(); dz.classList.add('on'); }; dz.ondragleave = () => dz.classList.remove('on');
  dz.ondrop = e => { e.preventDefault(); dz.classList.remove('on'); const f = e.dataTransfer.files[0]; if (f) leerArchivo(f); };
  $('#cElegir').onclick = e => e.stopPropagation(); $('#cElegir').addEventListener('click', e => { e.stopPropagation(); $('#cFile').click(); });
  $('#cCambiar').onclick = cReiniciar; $('#cCancelar').onclick = cReiniciar; $('#cOtro').onclick = cReiniciar;
  $('#cIcSube').innerHTML = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 16V5M7 10l5-5 5 5M5 19h14"/></svg>';
  $('#cIcDoc').innerHTML = ICO.file; $('#cOkIc').innerHTML = ICO.check;
  $('#cOkIr').onclick = e => irA(e.currentTarget.dataset.go);
}
const ALERTA = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 9v4m0 4h.01M10.3 3.9L1.8 18a2 2 0 001.7 3h17a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/></svg>';
const cFalta = r => [!r.agencia && 'sin agencia', !r.modo_pago && 'sin pago'].filter(Boolean).join(' y ');
function cAlertar() {
  const sinAg = previa.filter(r => !r.agencia).length, sinPg = previa.filter(r => !r.modo_pago).length, a = $('#cAlerta'), n = previa.filter(r => cFalta(r)).length;
  a.classList.toggle('hide', !n);
  const partes = [sinAg && `${sinAg} ${sinAg === 1 ? 'nodo sin agencia' : 'nodos sin agencia'}`, sinPg && `${sinPg} ${sinPg === 1 ? 'nodo sin pago' : 'nodos sin pago'}`].filter(Boolean).join(' y ');
  a.innerHTML = n ? `${ALERTA}<span><b>${partes}.</b> ${sinPg ? 'El pago es obligatorio: seleccionarlo en esta tabla. ' : ''}${sinAg ? 'La agencia puede completarse en esta tabla o después en Registros.' : ''}</span>${cFiltro === 'falta' ? '<button type="button" data-f="">Ver todos</button>' : '<button type="button" data-f="falta">Ver solo esos</button>'}` : '';
}
function pintarPrevia(anim) {
  $('#cDrop').classList.add('hide'); $('#cArch').classList.remove('hide'); $('#cAyuda').classList.add('hide'); $('#cGrid').classList.add('uno');
  $('#cArchN').textContent = cArchivo.nombre; $('#cArchS').textContent = `${cArchivo.kb} KB`;
  const des = previa.filter(r => r.motivo === 'DESPACHO'), rec = previa.filter(r => r.motivo === 'RECOJO'), suma = (a, k) => a.reduce((x, r) => x + (r[k] || 0), 0);
  const cant = a => `${fmtN(suma(a, 'pedidos'))} pedidos, ${fmtN(suma(a, 'bultos'))} bultos, ${fmtN(suma(a, 'cajas'))} cajas`;
  const nodosDet = hrnCarga ? new Set(hrnCarga.map(r => r.nodo)).size : 0, nodosRec = recDetCarga ? new Set(recDetCarga.map(r => r.nodo)).size : 0;
  $('#cTiles').innerHTML = `
    <div class="ctile${des.length ? '' : ' vacio'}"><span class="t"><i class="dot"></i>Despachos</span><b data-n="${des.length}">${des.length}</b><small>${des.length ? cant(des) : 'Sin filas en este archivo'}</small></div>
    <div class="ctile${rec.length ? '' : ' vacio'}"><span class="t"><i class="dot"></i>Recojos</span><b data-n="${rec.length}">${rec.length}</b><small>${rec.length ? cant(rec) : 'Sin filas en este archivo'}</small></div>
    <div class="ctile${hrnCarga ? '' : ' vacio'}"><span class="t"><i class="dot"></i>Detalle despacho</span><b data-n="${hrnCarga ? hrnCarga.length : 0}">${hrnCarga ? fmtN(hrnCarga.length) : 0}</b><small>${hrnCarga ? `filas de ${nodosDet} ${nodosDet === 1 ? 'nodo' : 'nodos'}, se agregan a las ya existentes` : 'Sin filas en este archivo'}</small></div>
    <div class="ctile${recDetCarga ? '' : ' vacio'}"><span class="t"><i class="dot"></i>Detalle recojos</span><b data-n="${recDetCarga ? recDetCarga.length : 0}">${recDetCarga ? fmtN(recDetCarga.length) : 0}</b><small>${recDetCarga ? `pedidos de ${nodosRec} ${nodosRec === 1 ? 'nodo' : 'nodos'}, se agregan a los ya existentes` : 'Sin filas en este archivo'}</small></div>`;
  cAlertar();
  $('#cTarjeta').classList.toggle('hide', !previa.length);
  const hayAmbos = des.length && rec.length;
  $('#cSeg').classList.toggle('hide', !hayAmbos && cFiltro !== 'falta');
  $('#cSeg').innerHTML = [['', 'Todos', previa.length], ['DESPACHO', 'Despachos', des.length], ['RECOJO', 'Recojos', rec.length]].map(([v, t, n]) => `<button data-f="${v}" class="${cFiltro === v ? 'on' : ''}">${t} <small>${n}</small></button>`).join('');
  const vis = previa.map((r, i) => [r, i]).filter(([r]) => cFiltro === 'falta' ? !!cFalta(r) : !cFiltro || r.motivo === cFiltro);
  $('#cTabla').innerHTML = '<tr><th>Fecha</th><th>Nodo</th><th>Motivo</th><th>Transporte</th><th class="c">Pedidos</th><th class="c">Bultos</th><th class="c">Cajas</th><th>Placa</th><th>Agencia</th><th>Pago</th><th>Zona</th></tr>' +
    vis.map(([r, i]) => `<tr data-i="${i}">
      <td data-c="fecha" data-l="Fecha"><input data-f="fecha" type="date" class="${r.fecha ? '' : 'vacia'}" value="${r.fecha || ''}"></td>
      <td data-c="nodo"><span class="nodo">${esc(r.nodo)}</span>${cFalta(r) ? `<br><span class="cfalta">${mayus(cFalta(r))}</span>` : ''}</td>
      <td data-c="motivo">${motTag(r.motivo)}</td>
      <td data-c="trans" data-l="Transporte"><input data-f="transporte" value="${esc(r.transporte)}"></td>
      <td data-c="n" data-l="Pedidos"><input data-f="pedidos" type="number" value="${r.pedidos ?? ''}"></td>
      <td data-c="n" data-l="Bultos"><input data-f="bultos" type="number" value="${r.bultos ?? ''}"></td>
      <td data-c="n" data-l="Cajas"><input data-f="cajas" type="number" value="${r.cajas ?? ''}"></td>
      <td data-c="placa" data-l="Placa"><input data-f="placa" value="${esc(r.placa)}" placeholder="ABC-123"></td>
      <td data-c="agencia" data-l="Agencia"><input data-f="agencia" list="dlAgencias" class="${r.agencia ? '' : 'vacia'}" value="${esc(r.agencia)}" placeholder="Ingresar agencia"></td>
      <td data-c="pago" data-l="Pago"><select data-f="modo_pago" data-v="${r.modo_pago}" class="${r.modo_pago ? '' : 'vacia'}"><option value=""${r.modo_pago ? '' : ' selected'}>Elegir…</option>${PAGOS.map(p => `<option${p === r.modo_pago ? ' selected' : ''}>${p}</option>`).join('')}</select></td>
      <td data-c="zona" data-l="Zona"><input data-f="zona" value="${esc(r.zona)}" placeholder="SUR, CENTRO…"></td></tr>`).join('');
  $('#cPrev').classList.remove('hide');
  $('#cBarT').textContent = [des.length && `${des.length} ${des.length === 1 ? 'despacho' : 'despachos'}`, rec.length && `${rec.length} ${rec.length === 1 ? 'recojo' : 'recojos'}`, hrnCarga && `${fmtN(hrnCarga.length)} filas de detalle`, recDetCarga && `${fmtN(recDetCarga.length)} pedidos de detalle de recojos`].filter(Boolean).join(', ');
  const fs = cFechasTodas(), sf = cSinFecha(); $('#cBarS').textContent = (fs.length > 1 ? `Se guardarán con ${fs.length} fechas: del ${fdmy(fs[0])} al ${fdmy(fs[fs.length - 1])}` : fs.length ? `Se guardarán con fecha ${fdmy(fs[0])}` : 'Sin fecha en el archivo') + (sf ? ` · ${sf} ${sf === 1 ? 'fila sin FECHA' : 'filas sin FECHA'}` : '');
  $('#cBarra').hidden = false; $('#tab-cargar').classList.add('conbarra');
  if (anim && !cReducir()) {   // solo al llegar a la revisión: las tarjetas y las filas entran, las cifras suben
    document.querySelectorAll('#cTiles b[data-n]').forEach(b => cContar(b, +b.dataset.n, 900)); [$('#cTiles'), $('#cTabla')].forEach(e => e.classList.add('ent')); setTimeout(() => [$('#cTiles'), $('#cTabla')].forEach(e => e.classList.remove('ent')), 1400);
  }
}
$('#cSeg').onclick = e => { const b = e.target.closest('button'); if (b) { cFiltro = b.dataset.f; pintarPrevia(); } };
$('#cAlerta').onclick = e => { const b = e.target.closest('button'); if (b) { cFiltro = b.dataset.f; pintarPrevia(); } };
$('#cTabla').addEventListener('change', e => {
  const tr = e.target.closest('tr'); if (!tr || !e.target.dataset.f) return;
  const f = e.target.dataset.f, v = e.target.value, r = previa[tr.dataset.i];
  r[f] = ['pedidos', 'bultos', 'cajas'].includes(f) ? num(v) : f === 'modo_pago' || f === 'fecha' ? v : ['zona', 'placa'].includes(f) ? v.trim().toUpperCase() : v.trim();
  if (f === 'fecha') return pintarPrevia();
  if (f === 'agencia' || f === 'modo_pago') {   // la marca y el aviso se actualizan sin recargar la tabla
    e.target.classList.toggle('vacia', !r[f]);
    const td = tr.querySelector('td[data-c=nodo]'), tag = td.querySelector('.cfalta'), t = cFalta(r);
    if (!t) { if (tag) { tag.previousElementSibling?.remove(); tag.remove(); } } else if (tag) tag.textContent = mayus(t); else td.insertAdjacentHTML('beforeend', `<br><span class="cfalta">${mayus(t)}</span>`);
    cAlertar();
  }
});
function cMostrarOk(titulo, texto, soloDetalle) {
  setPaso(3); ['#cS1', '#cPrev'].forEach(q => $(q).classList.add('hide')); $('#cBarra').hidden = true; $('#tab-cargar').classList.remove('conbarra');
  $('#cOkT').textContent = titulo; $('#cOkP').textContent = texto;
  $('#cOkIr').textContent = soloDetalle ? 'Ver detalle' : 'Completar importes en Registros'; $('#cOkIr').dataset.go = soloDetalle ? 'detalle' : 'registros';
  $('#cOk').classList.remove('hide'); scrollTo(0, 0);
}
const mayus = t => t.charAt(0).toUpperCase() + t.slice(1);

async function cGuardarDetalles() {
  const t = [];
  if (hrnCarga) t.push('detalle: ' + hrnTexto(await guardarHrn(hrnCarga)));
  if (recDetCarga) t.push('detalle de recojos: ' + hrnTexto(await guardarRecojoDet(recDetCarga)));
  return t.join('; ');
}
$('#cGuardar').onclick = async () => {
  if (cSinFecha()) return flash($('#cMsg'), (previa.some(r => !r.fecha) ? 'Existen filas sin fecha. Completar la fecha en la tabla.' : 'Existen filas del Detalle sin FECHA. Agregar la fecha en el Excel y volver a cargar el archivo.'), 'err');
  if (!previa.length) {   // solo detalle
    $('#cGuardar').disabled = true;
    try { cMostrarOk('Detalle guardado', mayus(await cGuardarDetalles()) + '.', true); previa = []; hrnCarga = null; recDetCarga = null; $('#cFile').value = ''; }
    catch (err) { flash($('#cMsg'), 'No se pudo guardar el detalle: ' + err.message + '. Volver a cargar el archivo: las filas ya registradas no se duplicarán.', 'err'); }
    $('#cGuardar').disabled = false; return;
  }
  const fechas = cFechas();
  const repetidos = new Set(), vistos = new Set(); previa.forEach(r => { const k = r.fecha + '|' + norm(r.nodo) + '|' + r.motivo; if (vistos.has(k)) repetidos.add(r.nodo + (r.motivo === 'RECOJO' ? ' (recojo)' : '')); vistos.add(k); });
  if (repetidos.size) return flash($('#cMsg'), `Existen nodos repetidos con el mismo motivo y fecha: ${[...repetidos].join(', ')}. Mantener una sola fila por nodo, motivo y fecha.`, 'err');
  const sinPago = previa.filter(r => !r.modo_pago);
  if (sinPago.length) { cFiltro = 'falta'; pintarPrevia(); return flash($('#cMsg'), `Falta registrar el modo de pago en ${sinPago.length} ${sinPago.length === 1 ? 'nodo' : 'nodos'}: ${sinPago.slice(0, 4).map(r => r.nodo).join(', ')}${sinPago.length > 4 ? '…' : ''}. Seleccionarlo en la tabla.`, 'err'); }
  if (previa.some(r => !r.agencia) && !await confirmar('Hay nodos sin agencia', 'La agencia puede completarse después en Registros. ¿Desea guardar de todos modos?', 'Guardar')) return;
  $('#cGuardar').disabled = true;
  const { data: ya, error: eYa } = await sb.from('envios').select('fecha,nodo,motivo,liquidacion_id').in('fecha', fechas);
  if (eYa) { $('#cGuardar').disabled = false; return flash($('#cMsg'), /motivo/.test(eYa.message) ? 'Falta ejecutar supabase_motivo.sql en Supabase.' : eYa.message, 'err'); }
  const clave = (f, n, m) => f + '|' + norm(n) + '|' + m, enviados = new Set((ya || []).filter(x => x.liquidacion_id).map(x => clave(String(x.fecha).slice(0, 10), x.nodo, x.motivo))), nuevosN = previa.filter(r => !enviados.has(clave(r.fecha, r.nodo, r.motivo)));
  const txtFecha = fechas.length > 1 ? `de ${fechas.length} fechas (${fdmy(fechas[0])} al ${fdmy(fechas[fechas.length - 1])})` : 'del ' + fdmy(fechas[0]);
  if ((ya || []).length && !await confirmar(`Ya hay ${ya.length} ${ya.length === 1 ? 'registro' : 'registros'} ${txtFecha}`, `Se actualizarán nodo, transporte, cantidades, placa, agencia, pago y zona con los datos de este archivo; el importe y la factura se conservan.` + (enviados.size ? ` ${enviados.size} ${enviados.size === 1 ? 'registro ya está enviado y no se tocará' : 'registros ya están enviados y no se tocarán'}.` : ''), 'Actualizar')) { $('#cGuardar').disabled = false; return; }
  if (!nuevosN.length) { $('#cGuardar').disabled = false; return flash($('#cMsg'), 'Todos los nodos de este archivo ya están enviados en una liquidación.', 'err'); }
  const filas = nuevosN.map(r => ({ fecha: r.fecha, origen: 'LIMA', transporte: r.transporte, nodo: r.nodo, bultos: r.bultos, pedidos: r.pedidos, cajas: r.cajas,
    placa: r.placa || null, agencia: r.agencia || null, zona: r.zona || null, detalle_gasto: r.detalle_gasto, modo_pago: r.modo_pago, motivo: r.motivo }));
  const { error } = await sb.from('envios').upsert(filas, { onConflict: 'fecha,nodo,motivo' });
  $('#cGuardar').disabled = false;
  if (error) return flash($('#cMsg'), /motivo/.test(error.message) ? 'Falta ejecutar supabase_motivo.sql en Supabase.' : error.message, 'err');
  $('#rDesde').value = fechas[0]; $('#rHasta').value = fechas[fechas.length - 1]; buscarRegistros();
  let detalle = '';
  if (hrnCarga || recDetCarga) {
    try { detalle = ' ' + mayus(await cGuardarDetalles()) + '.'; }
    catch (err) { cReiniciar(); return flash($('#cMsg'), `El despacho se guardó, pero el detalle no: ${err.message}. Volver a cargar el mismo archivo: las filas del detalle ya registradas no se duplicarán.`, 'err'); }
  }
  const hayD = filas.some(r => r.motivo === 'DESPACHO'), hayR = filas.some(r => r.motivo === 'RECOJO');
  cMostrarOk(hayD && hayR ? 'Despachos y recojos guardados' : hayR ? (filas.length === 1 ? 'Recojo guardado' : 'Recojos guardados') : (filas.length === 1 ? 'Despacho guardado' : 'Despachos guardados'), `${filas.length} ${filas.length === 1 ? 'registro' : 'registros'} ${txtFecha}.${detalle} Falta completar importe y factura de cada uno.`, false);
  previa = []; hrnCarga = null; recDetCarga = null; $('#cFile').value = '';
};
setPaso(1);

// ---------- adjuntos de factura (bucket "Facturas") ----------
const BUCKET = 'Facturas';
const MAX_ADJUNTO = 20 * 1024;   // todo adjunto se guarda con 20 KB o menos
// enlace temporal firmado: el bucket es privado y la foto solo se ve con sesión iniciada (se reutiliza durante 50 minutos)
const urlsFirmadas = new Map();
async function urlAdjunto(path) {
  const c = urlsFirmadas.get(path); if (c && Date.now() - c.t < 50 * 60 * 1000) return c.u;
  const { data, error } = await sb.storage.from(BUCKET).createSignedUrl(path, 3600);
  if (error || !data) return '';
  urlsFirmadas.set(path, { u: data.signedUrl, t: Date.now() }); return data.signedUrl;
}
const esImagen = path => /\.(png|jpe?g|gif|webp|bmp|heic)$/i.test(path);
let adjId = null;
$('#adjFile').onchange = async e => {
  const f = e.target.files[0]; e.target.value = '';
  if (!f || !adjId) return;
  const r = regBase.find(x => x.id === adjId); if (!r) return;
  if (!/^(image\/|application\/pdf)/.test(f.type)) return flash($('#rMsg'), 'Solo se aceptan imágenes o PDF.', 'err');
  const esImg = f.type.startsWith('image/');
  if (f.size > 30 * 1024 * 1024) return flash($('#rMsg'), 'El archivo supera los 30 MB.', 'err');
  if (!esImg && f.size > MAX_ADJUNTO) return flash($('#rMsg'), 'El PDF supera 20 KB. Cargar una foto o captura de la factura: las imágenes se comprimen automáticamente.', 'err');
  let subir = f, tipo = f.type, ext = 'pdf';
  if (esImg) {
    flash($('#rMsg'), 'Comprimiendo imagen…', 'ok');
    try { const c = await OCR.comprimir(f, MAX_ADJUNTO); subir = c.blob; tipo = c.tipo; ext = c.ext; }
    catch (err) { return flash($('#rMsg'), 'No se pudo comprimir la imagen (¿formato no compatible?): ' + err.message, 'err'); }
  }
  const path = `${r.fecha.slice(0, 7)}/${r.id}-${Date.now()}.${ext}`;
  flash($('#rMsg'), 'Subiendo adjunto…', 'ok');
  const up = await sb.storage.from(BUCKET).upload(path, subir, { contentType: tipo, upsert: false });
  if (up.error) return flash($('#rMsg'), 'No se pudo subir: ' + up.error.message, 'err');
  marcarMio(r.id);
  const { error } = await sb.from('envios').update({ factura_archivo: path }).eq('id', r.id);
  if (error) { await sb.storage.from(BUCKET).remove([path]); return flash($('#rMsg'), error.message, 'err'); }
  if (r.factura_archivo) await sb.storage.from(BUCKET).remove([r.factura_archivo]);
  r.factura_archivo = path; refrescarAdjunto(r);
  flash($('#rMsg'), `Adjunto guardado (${(subir.size / 1024).toFixed(1)} KB)`, 'ok');
  if (esImg) revisarFactura(f, r);
};

// ---------- lectura automática de la factura (OCR gratuito) + revisión ----------
let fReg = null, fUrl = null, fDup = null, fSeq = 0, fT = null;
function cerrarFactura() { $('#factDlg').close(); if (fUrl) { URL.revokeObjectURL(fUrl); fUrl = null; } }
// Una sola línea de aviso bajo cada campo; "Guardar" se bloquea mientras la factura esté repetida o se esté leyendo
function fPintar(leyendo) {
  const r = fReg, fac = $('#fFactura').value.trim(), imp = $('#fImporte').value.trim();
  const mFac = $('#fmFac'), mImp = $('#fmImp');
  $('#fcFac').className = 'fcampo' + (fDup ? ' mal' : !fac && !leyendo ? ' vacio' : '');
  mFac.className = 'fmsg ' + (fDup ? 'bad' : 'warn');
  mFac.textContent = fDup ? `Ya está registrado en ${fDup.nodo} (${fdmy(fDup.fecha)}).` : '';
  $('#fcImp').className = 'fcampo' + (!imp && !leyendo ? ' vacio' : '');
  mImp.className = 'fmsg';
  if (r && r.importe != null && imp !== '' && Math.abs(Number(r.importe) - num(imp)) > 0.01) mImp.innerHTML = `Antes <s>S/ ${Number(r.importe).toFixed(2)}</s>`; else mImp.textContent = '';
  $('#fGuardar').disabled = !!leyendo || !!fDup;
}
async function fChequear() {
  const mi = ++fSeq, v = $('#fFactura').value.trim();
  const o = facturaCambio(v, fReg) ? await facturaRepetida(v, fReg.id) : null;
  if (mi !== fSeq) return;
  fDup = o; fPintar(false);
}
async function revisarFactura(file, r) {
  fReg = r; fDup = null; fSeq++; clearTimeout(fT);
  $('#fFactura').value = r.factura || ''; $('#fImporte').value = r.importe != null ? r.importe : '';
  $('#fSub').textContent = `${r.nodo} · ${fdmy(r.fecha)}`;
  fUrl = URL.createObjectURL(file); $('#fImg').src = fUrl; fZoom(1);
  const est = $('#fEstado'), txt = est.querySelector('span'); est.hidden = false; est.className = 'fd-ley'; txt.textContent = 'Leyendo la factura…';
  fPintar(true);
  $('#factDlg').showModal();
  let fallo = false;
  try {
    const R = await OCR.leer(file, { ref: r.fecha, nodos: [...new Set(regBase.map(r => norm(r.nodo)))],
      progreso: (t, p) => { txt.textContent = 'Leyendo la factura… ' + (p ? Math.round(p * 100) + '%' : t); } });
    if (R.factura) $('#fFactura').value = R.factura;
    if (R.importe != null && R.importe !== '') $('#fImporte').value = R.importe;
    est.hidden = true;
    const vistos = R.destinosVistos || [];
    if (vistos.length && !vistos.includes(norm(r.nodo))) { est.hidden = false; est.className = 'fd-ley err'; txt.textContent = `La factura menciona ${vistos.slice(0, 3).join(', ')} y este registro es ${r.nodo}.`; }
  } catch (err) {
    fallo = true; est.hidden = false; est.className = 'fd-ley err'; txt.textContent = 'No se pudo leer la factura. Ingresar los datos manualmente a partir de la foto o cerrar con "Solo adjunto".';
  }
  fPintar(false);
  if (!fallo && $('#fFactura').value.trim()) fChequear();
}
$('#fFactura').addEventListener('input', () => { fPintar(false); clearTimeout(fT); fT = setTimeout(fChequear, 400); });
$('#fImporte').addEventListener('input', () => fPintar(false));
// Zoom de la foto: botones + / −, Ctrl + rueda, doble clic y arrastre para moverse
let fZ = 1;
function fZoom(z, cx, cy) {
  const sc = $('#fSc'), img = $('#fImg'), n = Math.min(4, Math.max(1, Math.round(z * 100) / 100));
  if (cx == null) { cx = sc.clientWidth / 2; cy = sc.clientHeight / 2; }
  const rx = img.offsetWidth ? (sc.scrollLeft + cx) / img.offsetWidth : 0, ry = img.offsetHeight ? (sc.scrollTop + cy) / img.offsetHeight : 0;
  fZ = n; img.style.width = n > 1 ? n * 100 + '%' : '';
  sc.scrollLeft = rx * img.offsetWidth - cx; sc.scrollTop = ry * img.offsetHeight - cy;
  sc.classList.toggle('z', n > 1); $('#fZp').textContent = Math.round(n * 100) + '%'; $('#fZm').disabled = n <= 1; $('#fZa').disabled = n >= 4;
}
$('#fZa').onclick = () => fZoom(fZ + .5); $('#fZm').onclick = () => fZoom(fZ - .5);
{
  const sc = $('#fSc'); let dr = null;
  const pos = e => { const b = sc.getBoundingClientRect(); return [e.clientX - b.left, e.clientY - b.top]; };
  sc.addEventListener('wheel', e => { if (!e.ctrlKey) return; e.preventDefault(); fZoom(fZ * (e.deltaY < 0 ? 1.2 : 1 / 1.2), ...pos(e)); }, { passive: false });
  sc.addEventListener('dblclick', e => fZoom(fZ > 1 ? 1 : 2.5, ...pos(e)));
  sc.addEventListener('pointerdown', e => { if (fZ <= 1 || e.button || e.pointerType !== 'mouse') return; dr = { x: e.clientX, y: e.clientY, l: sc.scrollLeft, t: sc.scrollTop }; sc.setPointerCapture(e.pointerId); sc.classList.add('arr'); });
  sc.addEventListener('pointermove', e => { if (!dr) return; sc.scrollLeft = dr.l - (e.clientX - dr.x); sc.scrollTop = dr.t - (e.clientY - dr.y); });
  const fin = () => { dr = null; sc.classList.remove('arr'); };
  sc.addEventListener('pointerup', fin); sc.addEventListener('pointercancel', fin);
}
$('#fCerrar').onclick = cerrarFactura;
$('#factDlg').addEventListener('close', () => { if (fUrl) { URL.revokeObjectURL(fUrl); fUrl = null; } });
$('#fGuardar').onclick = async () => {
  if (!fReg) return;
  const v = id => $('#' + id).value.trim(), nn = id => { const x = v(id); return x === '' ? null : num(x); };
  const cambios = { factura: v('fFactura') || null, importe: nn('fImporte') };
  const est = $('#fEstado'), txt = est.querySelector('span'), avisar = t => { est.hidden = false; est.className = 'fd-ley err'; txt.textContent = t; };
  if (facturaCambio(cambios.factura, fReg)) {
    $('#fGuardar').disabled = true; const o = await facturaRepetida(cambios.factura, fReg.id);
    fDup = o; fPintar(false); if (o) return;
  }
  $('#fGuardar').disabled = true;
  marcarMio(fReg.id);
  const { error } = await sb.from('envios').update(cambios).eq('id', fReg.id);
  $('#fGuardar').disabled = false;
  if (error) return avisar('No se pudo guardar: ' + error.message);
  Object.assign(fReg, cambios); cerrarFactura(); flash($('#rMsg'), 'Datos de la factura guardados', 'ok');
  repintarFila(fReg); contarEstados(); pintarResumenRegistros(); rmSync(fReg);
};
function refrescarAdjunto(r) {
  repintarFila(r); contarEstados();
  if ($('#regDlg').open && rm.id === r.id) { rmFoto(r); rmVivo(); }
}

// ---------- registros ----------
let pag = { n: 1, size: 25 };
const PIN = '<svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" clip-rule="evenodd" d="M5.05 4.05a7 7 0 119.9 9.9L10 18.9l-4.95-4.95a7 7 0 010-9.9zM10 11a2 2 0 100-4 2 2 0 000 4z"/></svg>';
const TRASH = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>';
const TIC = {
  doc: 'M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z',
  money: 'M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z',
  box: 'M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4',
  pin: 'M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0zM15 11a3 3 0 11-6 0 3 3 0 016 0z' };
const tile = (k, color) => `<i class="tile ${color}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="${TIC[k]}"/></svg></i>`;
const TCOLORES = { ELIAS: 'blue', LLANA: 'emerald', RELUFI: 'violet' }, TPAL = ['amber', 'rose', 'cyan', 'blue', 'emerald', 'violet'];
const tbadge = t => !t ? '' : `<span class="tb ${TCOLORES[norm(t)] || TPAL[[...norm(t)].reduce((a, c) => a + c.charCodeAt(0), 0) % TPAL.length]}">${esc(t)}</span>`;

// el resumen cuenta todos los registros del rango (sin el filtro de estado) y cada pestaña filtra por su estado
function pintarResumenRegistros() {
  const d = regBase, imp = a => a.reduce((s, r) => s + (Number(r.importe) || 0), 0), conImp = r => Number(r.importe) > 0;
  const L = d.filter(regTerminado), E = d.filter(regEnviado), F = d.filter(r => !regListo(r) && !regEnviado(r));
  $('#rSeg').innerHTML = [['', 'Todos', 'Todos', d.length], ['term', 'Listos para liquidar', 'Listos', L.length], ['pend', 'Faltan datos', 'Faltan', F.length], ['env', 'Enviados', 'Enviados', E.length]]
    .map(([e, t, c, n]) => `<button type="button" role="tab" data-e="${e}" aria-selected="${estadoReg === e}" class="${estadoReg === e ? 'on' : ''}"><span class="rx-l">${t}</span><span class="rx-s">${c}</span> <b>${n}</b>${e === 'pend' && n ? '<i></i>' : ''}</button>`).join('');
  const R = registros, n = R.length, per = `<span class="mper"> · ${esc(textoPeriodo())}</span>`;
  let t;
  if (!n && estadoReg === 'env') t = 'Ningún registro enviado en este rango';
  else if (estadoReg === 'term') t = `<b>${n}</b> registros · <b>${fmt(imp(R))}</b> · listos para liquidar`;
  else if (estadoReg === 'pend') {
    const sI = R.filter(r => !conImp(r)).length, sF = R.filter(r => conImp(r) && !r.factura).length, sP = R.filter(r => conImp(r) && r.factura && !r.factura_archivo).length;
    t = `<b>${n}</b> registros · <b>${fmt(imp(R))}</b> · <span class="w">${[sI && `${sI} sin importe`, sF && `${sF} sin N° de factura`, sP && `${sP} sin foto`].filter(Boolean).join(' · ') || 'todo completo'}</span>`;
  } else if (estadoReg === 'env') t = `<b>${n}</b> registros · <b>${fmt(imp(R))}</b> · en liquidaciones`;
  else t = `<b>${n}</b> ${n === 1 ? 'registro' : 'registros'} · <b>${fmt(imp(R))}</b> · ${fmtN(R.reduce((s, r) => s + (Number(r.pedidos) || 0), 0))} pedidos`;
  $('#rTotal').innerHTML = t + per;
  pintarFrase();
}
const lqCod = n => 'LIQ-' + String(n).padStart(4, '0');
// avance del registro: importe, factura, foto y enviado
const PASOS = ['Importe', 'Factura', 'Foto', 'Enviado'];
const regPasos = r => [Number(r.importe) > 0, !!r.factura, !!r.factura_archivo, regEnviado(r)];
const regFalta = r => regEnviado(r) ? ['e', 'Enviado' + (lqNumeros.has(r.liquidacion_id) ? ' ' + lqCod(lqNumeros.get(r.liquidacion_id)) : '')] : !(Number(r.importe) > 0) ? ['w', 'Falta importe'] : !r.factura ? ['w', 'Falta factura'] : !r.factura_archivo ? ['w', 'Falta foto'] : regNoLiq(r) ? ['c', r.modo_pago === 'CREDITO' ? 'Completo · crédito' : 'Completo'] : ['ok', 'Lista para liquidar'];
function stepper(r, grande) {
  const p = regPasos(r), sig = p.findIndex((x, i) => !x && !(i === 3 && regNoLiq(r)));
  return '<div class="stp' + (grande ? ' g' : '') + '">' + p.map((h, i) => {
    const punto = `<span class="pt ${h ? (i === 3 ? 'env' : 'hecho') : i === sig ? 'sig' : ''}" title="${PASOS[i]}">${ICO.check}</span>`;
    return (grande ? `<div class="it${h ? ' hecho' : ''}">${punto}<span>${PASOS[i]}</span></div>` : punto) + (i < 3 ? `<span class="ln${h ? ' hecho' : ''}"></span>` : '');
  }).join('') + '</div>';
}
function celdaEstado(r) {
  const [k, t] = regFalta(r), p = regPasos(r), sig = p.findIndex((x, i) => !x && !(i === 3 && regNoLiq(r)));
  const bar = p.map((h, i) => `<i class="${h ? (i === 3 ? 'v' : 'h') : i === sig && k === 'w' ? 'n' : ''}"></i>`).join('');
  const tip = p.map((h, i) => `<div><span>${PASOS[i]}</span><span class="${h ? 'y' : 'n'}">${h ? 'Listo' : i === 3 ? (regNoLiq(r) ? 'No aplica' : 'Pendiente') : 'Falta'}</span></div>`).join('');
  return `<div class="rest"><div class="lb ${k}">${t}</div><div class="bar">${bar}</div><div class="tip">${tip}</div></div>`;
}
function filaRegistro(r) {
  const [k] = regFalta(r), imp = Number(r.importe) > 0;
  const meta = [r.transporte, r.placa].filter(Boolean).map(esc).join('<i>·</i>');
  return `<tr data-id="${r.id}" class="${k}">
      <td data-c="fecha">${fdmy(r.fecha)}</td>
      <td data-c="nodo"><div class="nodo" title="${esc(r.nodo)}">${esc(r.nodo)}${r.motivo === 'RECOJO' ? '<span class="tb rec">Recojo</span>' : ''}</div>${meta ? `<div class="meta">${meta}</div>` : ''}</td>
      <td data-c="carga"><div class="carga"><div><b>${r.pedidos ?? '—'}</b><span>Pedidos</span></div><div><b>${r.bultos ?? '—'}</b><span>Bultos</span></div><div><b>${r.cajas ?? '—'}</b><span>Cajas</span></div></div></td>
      <td data-c="agencia" data-l="Agencia"><div class="ag" title="${esc(r.agencia)}">${r.agencia ? esc(r.agencia) : '<span class="sm">Sin agencia asignada</span>'}</div><div class="zn${r.zona ? '' : ' vacia'}">${r.zona ? esc(r.zona) : 'Sin zona'}</div></td>
      <td class="c" data-c="pago" data-l="Pago">${r.modo_pago ? `<span class="pgo ${esc(r.modo_pago)}">${esc(r.modo_pago)}</span>` : '<span class="sm">—</span>'}</td>
      <td class="num" data-c="importe">${imp ? `<div class="imp">S/ ${Number(r.importe).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}${r.factura ? `<small>${esc(r.factura)}</small>` : ''}</div>` : '<div class="imp falta">Falta importe</div>'}</td>
      <td data-c="prog">${celdaEstado(r)}</td>
      <td data-c="acc"><div class="acc"><button class="b edit" data-act="editar">${ICO.edit}<span>${regEnviado(r) ? 'Ver' : 'Editar'}</span></button><button type="button" class="rib" data-act="menu" aria-label="Más acciones" title="Más acciones"><svg viewBox="0 0 24 24" fill="currentColor"><circle cx="5" cy="12" r="1.8"/><circle cx="12" cy="12" r="1.8"/><circle cx="19" cy="12" r="1.8"/></svg></button></div></td></tr>`;
}
const repintarFila = r => { const tr = document.querySelector(`#rTabla tr[data-id="${r.id}"]`); if (tr) tr.outerHTML = filaRegistro(r); };
function pintarTablaRegistros() {
  const N = registros.length, paginas = Math.max(1, Math.ceil(N / pag.size));
  pag.n = Math.min(Math.max(1, pag.n), paginas);
  const ini = (pag.n - 1) * pag.size, vista = registros.slice(ini, ini + pag.size);
  $('#rTabla').innerHTML = '<tr><th data-c="fecha">Fecha</th><th data-c="nodo">Nodo</th><th data-c="carga">Carga</th><th data-c="agencia">Agencia y zona</th><th data-c="pago" class="c">Pago</th><th data-c="importe" class="num">Importe</th><th data-c="prog">Estado</th><th data-c="acc" aria-label="Acciones"></th></tr>' + (vista.map(filaRegistro).join('') || '<tr><td colspan="8"><div class="empty"><b>Sin registros en este rango</b><span>Ajustar el rango de fechas o cargar un despacho en «Cargar despacho».</span></div></td></tr>');
  const nums = []; for (let i = 1; i <= paginas; i++) if (i === 1 || i === paginas || Math.abs(i - pag.n) <= 1) nums.push(i); else if (nums[nums.length - 1] !== '…') nums.push('…');
  $('#rPag').innerHTML = `<div class="l"><span>Filas por página:</span><select id="rSize" aria-label="Filas por página">${[15, 25, 50, 100].map(v => `<option${v === pag.size ? ' selected' : ''}>${v}</option>`).join('')}</select><span>Mostrando <b>${N ? ini + 1 : 0}-${Math.min(N, ini + pag.size)}</b> de <b>${N}</b> registros</span></div>
    <div class="r"><button data-p="prev"${pag.n <= 1 ? ' disabled' : ''} aria-label="Anterior">${ICO.prev}</button>${nums.map(n => n === '…' ? '<span>…</span>' : `<button data-p="${n}" class="${n === pag.n ? 'on' : ''}">${n}</button>`).join('')}<button data-p="next"${pag.n >= paginas ? ' disabled' : ''} aria-label="Siguiente">${ICO.next}</button></div>`;
}
$('#rPag').addEventListener('click', e => {
  const b = e.target.closest('button[data-p]'); if (!b || b.disabled) return;
  pag.n = b.dataset.p === 'prev' ? pag.n - 1 : b.dataset.p === 'next' ? pag.n + 1 : Number(b.dataset.p); pintarTablaRegistros();
});
$('#rPag').addEventListener('change', e => { if (e.target.id === 'rSize') { pag.size = Number(e.target.value); pag.n = 1; pintarTablaRegistros(); } });

let rSeq = 0;
async function buscarRegistros() {
  const crear = () => {
    let q = sb.from('envios').select('*').order('fecha', { ascending: false }).order('transporte').order('nodo').order('id');
    if ($('#rDesde').value) q = q.gte('fecha', $('#rDesde').value);
    if ($('#rHasta').value) q = q.lte('fecha', $('#rHasta').value);
    if ($('#rTrans').value) q = q.eq('transporte', $('#rTrans').value);
    if ($('#rPago').value) q = q.eq('modo_pago', $('#rPago').value);
    if ($('#rMot').value) q = q.eq('motivo', $('#rMot').value);
    const t = $('#rNodo').value.replace(/[,()*%\\]/g, ' ').trim();   // nodo, agencia, placa o factura
    if (t) q = q.or(['nodo', 'agencia', 'placa', 'factura'].map(c => `${c}.ilike.*${t}*`).join(','));
    return q;
  };
  const mi = ++rSeq; let data; try { data = await traerTodo(crear); } catch (error) { conexion(false); return flash($('#rMsg'), error.message, 'err'); }
  if (mi !== rSeq) return;   // llegó una búsqueda más nueva
  conexion(true); rtUltimo = Date.now();
  regBase = data; listaAgencias();
  const { data: lqs } = await sb.from('liquidaciones').select('id,numero'); if (mi !== rSeq) return; lqNumeros = new Map((lqs || []).map(x => [x.id, x.numero]));
  const ts = [...new Set(data.map(r => r.transporte).filter(Boolean))].sort(), sel = $('#rTrans').value;
  if (!sel) $('#rTrans').innerHTML = '<option value="">Todos</option>' + ts.map(t => `<option>${esc(t)}</option>`).join('');
  aplicarEstado();
}
// sugerencias de agencia al escribir: las ya usadas en los registros y la lista guardada
function listaAgencias() { $('#dlAgencias').innerHTML = [...new Set([...agBase, ...regBase.map(r => r.agencia), ...previa.map(r => r.agencia)].filter(Boolean))].sort().map(n => `<option value="${esc(n)}">`).join(''); }
// Un registro está terminado cuando tiene importe, N° de factura y la factura adjunta
let regBase = [], estadoReg = '';
const regEnviado = r => !!r.liquidacion_id;
const regListo = r => Number(r.importe) > 0 && !!r.factura && !!r.factura_archivo;
// Solo se liquidan los registros al contado: los de crédito nunca pasan a "listos para liquidar"
const regNoLiq = r => r.modo_pago !== 'CONTADO';
const regTerminado = r => regListo(r) && !regEnviado(r) && !regNoLiq(r);
let lqNumeros = new Map();
const contarEstados = () => pintarResumenRegistros();
function aplicarEstado() {
  registros = estadoReg === 'term' ? regBase.filter(regTerminado) : estadoReg === 'env' ? regBase.filter(regEnviado) : estadoReg === 'pend' ? regBase.filter(r => !regListo(r) && !regEnviado(r)) : regBase;
  pintarResumenRegistros(); pintarTablaRegistros();
}
$('#rSeg').addEventListener('click', e => {
  const b = e.target.closest('button[data-e]'); if (!b) return;
  estadoReg = b.dataset.e; pag.n = 1; aplicarEstado();
});
$('#rRefrescar').onclick = () => buscarRegistros();
$('#rNuevo').onclick = () => irA('cargar');
// detalle de pasos al pasar el mouse por el estado: flota fuera de la tabla para no crear barras de desplazamiento
const rTip = $('#rTip'), ocultarTip = () => { rTip.hidden = true; };
$('#rTabla').addEventListener('mouseover', e => {
  const r = e.target.closest('.rest'); if (!r) return ocultarTip();
  rTip.innerHTML = r.querySelector('.tip').innerHTML; rTip.hidden = false;
  const b = r.getBoundingClientRect(), h = rTip.offsetHeight; let top = b.bottom + 6; if (top + h > innerHeight - 8) top = b.top - h - 6;
  rTip.style.top = Math.max(8, top) + 'px'; rTip.style.left = Math.max(8, Math.min(b.left, innerWidth - 198)) + 'px';
});
$('#rTabla').addEventListener('mouseleave', ocultarTip); window.addEventListener('scroll', ocultarTip, true);
const rMenu = $('#rMenu'), cerrarMenu = () => { rMenu.hidden = true; rMenu.dataset.id = ''; };
async function borrarRegistro(id) {
  const r = regBase.find(x => x.id === id);
  if (!await confirmar('¿Eliminar este registro?', 'No se puede deshacer.', 'Eliminar', true)) return;
  const { error } = await sb.from('envios').delete().eq('id', id);
  if (error) return flash($('#rMsg'), error.message, 'err');
  if (r?.factura_archivo) await sb.storage.from(BUCKET).remove([r.factura_archivo]);
  buscarRegistros();
}
$('#rTabla').addEventListener('click', e => {
  const ed = e.target.closest('[data-act=editar]'); if (ed) return abrirReg(ed.closest('tr').dataset.id, registros.map(x => x.id));
  const mn = e.target.closest('[data-act=menu]'); if (!mn) return;
  const id = mn.closest('tr').dataset.id, r = regBase.find(x => x.id === id);
  if (!rMenu.hidden && rMenu.dataset.id === id) return cerrarMenu();
  const b = rMenu.querySelector('button'); b.disabled = !r || regEnviado(r); b.title = b.disabled ? 'Pertenece a una liquidación. Anular la liquidación para poder eliminarlo.' : '';
  rMenu.dataset.id = id; rMenu.hidden = false;
  const rc = mn.getBoundingClientRect();
  rMenu.style.top = Math.max(8, Math.min(rc.bottom + 4, innerHeight - rMenu.offsetHeight - 8)) + 'px'; rMenu.style.right = Math.max(8, innerWidth - rc.right) + 'px';
});
rMenu.addEventListener('click', async e => { const b = e.target.closest('button[data-del]'); if (!b || b.disabled) return; const id = rMenu.dataset.id; cerrarMenu(); await borrarRegistro(id); });
document.addEventListener('click', e => { if (!rMenu.hidden && !e.target.closest('#rMenu, [data-act=menu]')) cerrarMenu(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !rMenu.hidden) cerrarMenu(); });
window.addEventListener('resize', cerrarMenu); window.addEventListener('scroll', cerrarMenu, true);

// ---------- filtros de Registros: se leen como una frase y cada parte se edita ----------
// Los datos viven en los campos de siempre (#rDesde, #rHasta, #rTrans, #rMot, #rPago, #rNodo); esto es solo la forma de verlos y cambiarlos.
const RFIL = {
  trans: { sel: '#rTrans', def: 'todos los transportes', txt: v => v },
  mot: { sel: '#rMot', def: 'despachos y recojos', txt: v => v === 'DESPACHO' ? 'solo despachos' : 'solo recojos', ops: [['', 'Despachos y recojos'], ['DESPACHO', 'Solo despachos'], ['RECOJO', 'Solo recojos']] },
  pago: { sel: '#rPago', def: 'todos los métodos de pago', txt: v => v === 'CONTADO' ? 'al contado' : 'a crédito', ops: [['', 'Todos los métodos de pago'], ['CONTADO', 'Al contado'], ['CREDITO', 'A crédito']] }
};
const RX_CHECK = '<svg viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 6.5L5 9.5 10 3"/></svg>';
const MESES_C = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'set', 'oct', 'nov', 'dic'];
const diaC = f => `${Number(f.slice(8, 10))} ${MESES_C[Number(f.slice(5, 7)) - 1]}`;
function textoPeriodo() {
  const d = $('#rDesde').value, h = $('#rHasta').value;
  if (!d && !h) return 'todas las fechas';
  if (!h) return `desde ${diaC(d)} ${d.slice(0, 4)}`; if (!d) return `hasta ${diaC(h)} ${h.slice(0, 4)}`;
  if (d === h) return `${diaC(d)} ${d.slice(0, 4)}`;
  if (d.slice(0, 4) !== h.slice(0, 4)) return `${diaC(d)} ${d.slice(0, 4)} – ${diaC(h)} ${h.slice(0, 4)}`;
  return d.slice(0, 7) === h.slice(0, 7) ? `${Number(d.slice(8, 10))} – ${diaC(h)} ${h.slice(0, 4)}` : `${diaC(d)} – ${diaC(h)} ${h.slice(0, 4)}`;
}
const opcionesDe = k => k === 'trans' ? [['', 'Todos los transportes'], ...[...$('#rTrans').options].filter(o => o.value).map(o => [o.value, o.value])] : RFIL[k].ops;
function pintarFrase() {
  $('#tkRango').textContent = textoPeriodo();
  let n = 0;
  Object.keys(RFIL).forEach(k => { const v = $(RFIL[k].sel).value, tk = document.querySelector(`.rx-tk[data-k="${k}"]`); tk.textContent = v ? RFIL[k].txt(v) : RFIL[k].def; tk.classList.toggle('set', !!v); if (v) n++; });
  if ($('#rNodo').value.trim()) n++;
  $('#rLimp').hidden = !n; $('#rfN').hidden = !n; $('#rfN').textContent = n; $('#rfBtn').classList.toggle('set', n > 0);
  $('#rfVer').textContent = `Ver ${registros.length} ${registros.length === 1 ? 'registro' : 'registros'}`;
  if ($('#rFilDlg').open) pintarHoja();
}
const rBuscarYa = () => { pag.n = 1; buscarRegistros(); };
function periodoPreset(p) {
  const iso = x => `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, '0')}-${String(x.getDate()).padStart(2, '0')}`, h = new Date(); let d = new Date(h);
  if (p === 'ayer') { d.setDate(d.getDate() - 1); h.setTime(d.getTime()); }
  else if (p === '7') d.setDate(d.getDate() - 6);
  else if (p === 'mes') d = new Date(h.getFullYear(), h.getMonth(), 1);
  else if (p === 'mp') { d = new Date(h.getFullYear(), h.getMonth() - 1, 1); h.setTime(new Date(h.getFullYear(), h.getMonth(), 0).getTime()); }
  $('#rDesde').value = iso(d); $('#rHasta').value = iso(h);
}
document.addEventListener('click', e => {
  if (!$('#tab-registros') || $('#tab-registros').classList.contains('hide')) return;
  const dentro = e.target.closest('.rx-fb'), tk = e.target.closest('.rx-tk[data-pop]');
  document.querySelectorAll('.rx-pop.on').forEach(p => { if (p.closest('.rx-fb') !== dentro) p.classList.remove('on'); });
  if (tk) {
    const pop = $('#' + tk.dataset.pop);
    if (tk.dataset.k) pop.innerHTML = opcionesDe(tk.dataset.k).map(([v, t]) => `<button type="button" class="rx-op${$(RFIL[tk.dataset.k].sel).value === v ? ' on' : ''}" data-k="${tk.dataset.k}" data-v="${esc(v)}">${esc(t)}${RX_CHECK}</button>`).join('');
    pop.classList.toggle('on'); return;
  }
  const op = e.target.closest('.rx-op');
  if (op) { $(RFIL[op.dataset.k].sel).value = op.dataset.v; op.closest('.rx-pop').classList.remove('on'); rBuscarYa(); return; }
  const pre = e.target.closest('#rxPre button');
  if (pre) { periodoPreset(pre.dataset.p); $('#rxp-rango').classList.remove('on'); rBuscarYa(); }
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') document.querySelectorAll('.rx-pop.on').forEach(p => p.classList.remove('on')); });
$('#rDesde').addEventListener('change', rBuscarYa); $('#rHasta').addEventListener('change', rBuscarYa);
let rTimer = null;
$('#rNodo').addEventListener('input', () => { clearTimeout(rTimer); rTimer = setTimeout(rBuscarYa, 350); pintarFrase(); });
$('#rNodo').addEventListener('keydown', e => { if (e.key === 'Enter') { clearTimeout(rTimer); rBuscarYa(); } });
const quitarFiltros = () => { Object.values(RFIL).forEach(f => { $(f.sel).value = ''; }); $('#rNodo').value = ''; rBuscarYa(); };
$('#rLimp').onclick = quitarFiltros;
// en el celular los filtros se editan en una hoja inferior
function pintarHoja() {
  ['trans', 'mot', 'pago'].forEach(k => { $('#rf' + k[0].toUpperCase() + k.slice(1)).innerHTML = opcionesDe(k).map(([v, t]) => `<button type="button" data-k="${k}" data-v="${esc(v)}" class="${$(RFIL[k].sel).value === v ? 'on' : ''}">${esc(t.replace(/^Todos los (transportes|métodos de pago)$/, 'Todos'))}</button>`).join(''); });
  $('#rfDesde').value = $('#rDesde').value; $('#rfHasta').value = $('#rHasta').value;
}
$('#rfBtn').onclick = () => { pintarHoja(); $('#rFilDlg').showModal(); };
$('#rFilDlg').addEventListener('click', e => {
  if (e.target === $('#rFilDlg')) return $('#rFilDlg').close();
  const b = e.target.closest('button[data-k]'); if (b) { $(RFIL[b.dataset.k].sel).value = b.dataset.v; return rBuscarYa(); }
  const p = e.target.closest('#rfPer button'); if (p) { periodoPreset(p.dataset.p); return rBuscarYa(); }
});
$('#rfDesde').addEventListener('change', e => { $('#rDesde').value = e.target.value; rBuscarYa(); });
$('#rfHasta').addEventListener('change', e => { $('#rHasta').value = e.target.value; rBuscarYa(); });
$('#rfQuitar').onclick = quitarFiltros; $('#rfVer').onclick = () => $('#rFilDlg').close();

// ---------- ventana para editar un registro ----------
let rm = { id: null, orden: [] };
const rmReg = () => regBase.find(x => x.id === rm.id);
const rmLeer = () => ({ placa: $('#rmPlaca').value.trim() || null, importe: $('#rmImp').value === '' ? null : num($('#rmImp').value), factura: $('#rmFac').value.trim() || null, detalle_gasto: $('#rmDet').value.trim() || null });
const rmSucio = () => { const r = rmReg(), v = rmLeer(); return !!r && !regEnviado(r) && Object.keys(v).some(k => (v[k] ?? '') != (r[k] ?? '')); };
function rmNav() {
  const i = rm.orden.indexOf(rm.id), sin = i < 0 || i >= rm.orden.length - 1;
  $('#rmPrev').disabled = i <= 0; $('#rmNext').disabled = sin; $('#rmSig').disabled = sin;
}
function abrirReg(id, orden) {
  const r = regBase.find(x => x.id === id); if (!r) return;
  rm.id = id; if (orden) rm.orden = orden;
  const bloq = regEnviado(r);
  $('#rmSub').textContent = fdmy(r.fecha); $('#rmNodo').textContent = r.nodo;
  $('#rmMeta').innerHTML = motTag(r.motivo) + tbadge(r.transporte) + (r.agencia ? `<span class="tb">${esc(r.agencia)}</span>` : '') + (r.zona ? `<span class="tb">Zona ${esc(r.zona)}</span>` : '') + (r.cajas != null ? `<span class="tb">${r.cajas} ${r.cajas === 1 ? 'caja' : 'cajas'}</span>` : '');
  $('#rmPlaca').value = r.placa ?? ''; $('#rmImp').value = r.importe == null ? '' : Number(r.importe).toFixed(2); $('#rmFac').value = r.factura ?? ''; $('#rmDet').value = r.detalle_gasto ?? '';
  ['#rmPlaca', '#rmImp', '#rmFac', '#rmDet'].forEach(q => $(q).disabled = bloq);
  const av = $('#rmAv'); av.hidden = !bloq; av.textContent = bloq ? `Está en ${lqNumeros.has(r.liquidacion_id) ? lqCod(lqNumeros.get(r.liquidacion_id)) : 'una liquidación'}. Anular la liquidación para poder editarlo.` : '';
  $('#rmOk').hidden = $('#rmSig').hidden = bloq; $('#rmCancel').textContent = bloq ? 'Cerrar' : 'Cancelar';
  rm.dup = null; clearTimeout(rmDupT); rmDupSeq++; rmMostrarDup(); rmFoto(r); rmVivo();
  if (!$('#regDlg').open) $('#regDlg').showModal();
}
function rmVivo() {
  const r = rmReg(); if (!r) return;
  const vista = { ...r, ...rmLeer() };
  $('#rmSt').innerHTML = stepper(vista, true); $('#rmHint').textContent = regFalta(vista)[1];
}
function rmFoto(r) {
  const f = r.factura_archivo, bloq = regEnviado(r);
  $('#rmTh').innerHTML = f ? (esImagen(f) ? '<img alt="Factura" style="visibility:hidden">' : ICO.file) : ICO.img;
  $('#rmTt').textContent = f ? 'Factura adjunta' : 'Sin foto de la factura';
  $('#rmTs').textContent = f ? 'Se guarda con 20 KB o menos' : (bloq ? '' : 'Tomar o seleccionar una foto: el N° y el importe se leen automáticamente');
  $('#rmAc').innerHTML = (f ? '<a class="b sec hide" id="rmVer" href="#" target="_blank" rel="noopener">Ver</a>' : '') + (bloq ? '' : f ? '<button type="button" class="b sec" data-f="cambiar">Cambiar</button><button type="button" class="b sec" data-f="quitar">Quitar</button>' : '<button type="button" class="b" data-f="cambiar">Adjuntar foto</button>');
  if (f) urlAdjunto(f).then(u => {   // el enlace firmado llega un instante después; se ignora si ya se cambió de registro o de foto
    if (!u || r.factura_archivo !== f || !$('#rmVer')) return;
    const img = $('#rmTh img'); if (img) { img.onload = () => { img.style.visibility = ''; }; img.src = u; }
    $('#rmVer').href = u; $('#rmVer').classList.remove('hide');
  });
}
// después de leer la factura con el OCR, el importe y el N° se reflejan en la ventana
function rmSync(r) {
  if (!$('#regDlg').open || rm.id !== r.id) return;
  $('#rmImp').value = r.importe == null ? '' : Number(r.importe).toFixed(2); $('#rmFac').value = r.factura ?? ''; rmFoto(r); rmVivo();
}
['#rmPlaca', '#rmImp', '#rmFac', '#rmDet'].forEach(q => $(q).addEventListener('input', rmVivo));
$('#rmPrev').innerHTML = ICO.prev; $('#rmNext').innerHTML = ICO.next; $('#rmX').innerHTML = ICO.x;
$('#rmAc').addEventListener('click', async e => {
  const b = e.target.closest('button[data-f]'), r = rmReg(); if (!b || !r) return;
  if (b.dataset.f === 'cambiar') { adjId = r.id; $('#adjFile').click(); return; }
  if (!await confirmar('¿Quitar el adjunto?', 'La foto o el PDF se borrará de este registro.', 'Quitar', true)) return;
  marcarMio(r.id);
  const { error } = await sb.from('envios').update({ factura_archivo: null }).eq('id', r.id);
  if (error) return flash($('#rMsg'), error.message, 'err');
  await sb.storage.from(BUCKET).remove([r.factura_archivo]);
  r.factura_archivo = null; refrescarAdjunto(r); flash($('#rMsg'), 'Adjunto quitado', 'ok');
});
// Un N° de factura no puede repetirse: devuelve el otro registro que ya lo usa (o null si está libre)
async function facturaRepetida(factura, id) {
  const f = String(factura || '').trim(); if (!f) return null;
  const { data, error } = await sb.from('envios').select('id,nodo,fecha').ilike('factura', f.replace(/[\\%_]/g, '\\$&')).neq('id', id).limit(1);
  return !error && data && data[0] || null;
}
const facturaCambio = (nueva, r) => !!nueva && String(nueva).trim().toUpperCase() !== String(r.factura || '').trim().toUpperCase();
const textoRepetida = (f, o) => `El N° de factura ${f} ya está registrado en ${o.nodo} (${fdmy(o.fecha)}). No se puede repetir.`;
let rmDupSeq = 0, rmDupT = null;
function rmMostrarDup() {
  const o = rm.dup; $('#rmDup').hidden = !o; $('#rmFac').classList.toggle('dup', !!o);
  $('#rmDup').textContent = o ? textoRepetida($('#rmFac').value.trim(), o) : '';
  if (o) $('#rmOk').disabled = $('#rmSig').disabled = true; else { $('#rmOk').disabled = false; rmNav(); }
}
async function rmChequearFactura() {
  const r = rmReg(); if (!r) return; const v = $('#rmFac').value.trim(), mi = ++rmDupSeq;
  const otro = facturaCambio(v, r) ? await facturaRepetida(v, r.id) : null;
  if (mi !== rmDupSeq) return;
  rm.dup = otro; rmMostrarDup();
}
$('#rmFac').addEventListener('input', () => { clearTimeout(rmDupT); rmDupT = setTimeout(rmChequearFactura, 400); });
async function guardarReg(siguiente) {
  const r = rmReg(); if (!r || regEnviado(r)) return;
  const v = rmLeer();
  if (v.importe != null && v.importe < 0) return toast('El importe no puede ser negativo', 'err');
  if (facturaCambio(v.factura, r)) {
    $('#rmOk').disabled = $('#rmSig').disabled = true; const o = await facturaRepetida(v.factura, r.id); rm.dup = o; rmMostrarDup();
    if (o) { $('#rmFac').focus(); return toast(textoRepetida(v.factura, o), 'err'); }
  }
  if (rmSucio()) {
    $('#rmOk').disabled = $('#rmSig').disabled = true;
    marcarMio(r.id);
    const { error } = await sb.from('envios').update(v).eq('id', r.id);
    $('#rmOk').disabled = false; rmNav();
    if (error) return toast(error.message, 'err');
    Object.assign(r, v); repintarFila(r); contarEstados(); pintarResumenRegistros(); toast('Guardado');
  }
  if (siguiente) abrirReg(rm.orden[rm.orden.indexOf(rm.id) + 1]); else $('#regDlg').close();
}
async function rmIr(delta) {
  if (rmSucio() && !await confirmar('Hay cambios sin guardar', 'Al cambiar de registro se perderán.', 'Descartar', true)) return;
  abrirReg(rm.orden[rm.orden.indexOf(rm.id) + delta]);
}
$('#rmF').onsubmit = e => { e.preventDefault(); guardarReg(false); };
$('#rmSig').onclick = () => guardarReg(true);
$('#rmPrev').onclick = () => rmIr(-1); $('#rmNext').onclick = () => rmIr(1);
$('#rmCancel').onclick = $('#rmX').onclick = async () => { if (rmSucio() && !await confirmar('Hay cambios sin guardar', 'Al cerrar se perderán.', 'Descartar', true)) return; $('#regDlg').close(); };

$('#rExport').onclick = async () => {
  await xlsxLib();
  const serial = f => Math.round((Date.parse(f + 'T00:00:00Z') - Date.parse('1899-12-30T00:00:00Z')) / 86400000);
  const aoa = [['MES','FECHA','ORIGEN','TRANSPORTE','NODOS','Cant. Bultos','Cant. Pedidos','Cant -CAJAS','PLACAS','AGENCIAS','DETALLE DEL GASTO','MODO DE PAGO','IMPORTE','OBSERVACIÓN','CONC','MOTIVO','ZONA']]
    .concat(registros.map(r => [MESES[Number(r.fecha.slice(5, 7)) - 1], serial(r.fecha), r.origen, r.transporte, r.nodo, r.bultos, r.pedidos, r.cajas,
      r.placa, r.agencia, r.detalle_gasto, r.modo_pago, r.importe, r.factura, serial(r.fecha) + r.nodo, r.motivo === 'RECOJO' ? 'RECOJO' : 'DESPACHO', r.zona]));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = aoa[0].map(() => ({ wch: 16 }));
  aoa.slice(1).forEach((_, i) => { const c = ws['B' + (i + 2)]; if (c) c.z = 'dd-mmm'; });
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'ACTUAL');
  XLSX.writeFile(wb, `red_troncal_${$('#rDesde').value || 'inicio'}_${$('#rHasta').value || 'hoy'}.xlsx`);
};

// ---------- detalle HRN ----------
const HRN_COLS = [['pedido_cliente','Pedido Cliente','t'],['nro_doc_referencia','Nro. Doc. Referencia','t'],['tipo_documento','Tipo Documento','t'],['lpn','LPN','t'],['fecha_recepcion','Fecha de Recepcion','d'],['ubicacion_recepcion','Ubicacion Recepción','t'],['cuenta_procedencia','Cuenta Procedencia','t'],['descripcion_cuenta','Descripcion Cuenta','t'],['nodo','NODO','t'],['bultos','BULTOS','n'],['observacion','OBSERVACION','t'],['fecha_ruta','Fecha Ruta','d'],['flag_carga_stage','Flag Carga Stage','n'],['fecha_carga_stage','Fecha Carga Stage','d'],['flag_terminar_carga','Flag Terminar Carga','n'],['fecha_terminar_carga','Fecha Terminar Carga','d'],['fecha_despacho','Fecha Despacho','d'],['flag_reprogramado','Flag Reprogramado','n'],['contador_reprogramado','Contador Reprogramado','n'],['motivo_pedido','Motivo Pedido','t'],['alto','Alto','n'],['ancho','Ancho','n'],['largo','Largo','n'],['volumen','Volumen','n'],['peso','Peso','n'],['roll_contenedor','Roll Contenedor','t'],['codigo_zona','Código Zona','t'],['zona','Zona','t'],['orden_clasificacion','Orden Clasificación','t'],['estado_contenedor','Estado Contenedor','t'],['placa_real','Placa Real','t'],['fecha_pedido','Fecha de Pedido','d'],['fecha_creacion','Fecha de Creación','d'],['ubicacion_movimiento','Ubicación Movimiento','t'],['fecha_movimiento','Fecha Movimiento','d'],['usuario_movimiento','Usuario Movimiento','t'],['proveedor','PROVEEDOR','t'],['nodo2','NODO2','t'],['cant_pedidos','CANT_PEDIDOS','t']];
const hrnPorHeader = new Map(HRN_COLS.map(c => [norm(c[1]), c]));

function serialAFecha(n) {
  const d = new Date(Math.round((n - 25569) * 86400) * 1000);
  return d.toISOString().slice(0, 19).replace('T', ' ');
}
function valorHrn(v, tipo) {
  if (v === '' || v == null) return null;
  if (tipo === 'n') return num(v);
  if (tipo === 'd') return typeof v === 'number' ? serialAFecha(v) : (String(v).trim() || null);
  return String(v).trim();
}
function leerHrn(wb) {
  const hojas = [...wb.SheetNames].sort((a, b) => /DETALLE|HRN/.test(norm(b)) - /DETALLE|HRN/.test(norm(a)));
  for (const nombre of hojas) {
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, defval: '', raw: true });
    const hi = filas.findIndex(f => f.some(c => norm(c) === 'PEDIDO CLIENTE'));
    if (hi < 0) continue;
    const mapa = filas[hi].map(h => hrnPorHeader.get(norm(h)) || null);
    if (!mapa.some(c => c && c[0] === 'nodo')) continue;
    const iFecha = filas[hi].findIndex(h => ['FECHA', 'FECHA REPORTE', 'FECHA DE REPORTE'].includes(norm(h)));   // fecha con la que se guarda cada fila
    const out = [];
    for (const f of filas.slice(hi + 1)) {
      if (!mapa.some((c, i) => c && f[i] !== '' && f[i] != null)) continue;
      const r = {};
      HRN_COLS.forEach(c => r[c[0]] = null);
      mapa.forEach((c, i) => { if (c) r[c[0]] = valorHrn(f[i], c[2]); });
      r.fecha_reporte = iFecha >= 0 ? fechaDe(f[iFecha]) : '';
      if (r.nodo) r.nodo = norm(r.nodo);
      if (r.pedido_cliente || r.lpn) out.push(r);
    }
    if (out.length) return out;
  }
  return null;
}

// ---------- detalle de recojos (devoluciones) ----------
// columnas del reporte; antes de ellas el Excel lleva FECHA y NOMBRE CUENTA
const RC_COLS = [['nombre_cuenta','NOMBRE CUENTA','t'],['cuenta','CUENTA','t'],['nro_pedido','N° DE PEDIDO','t'],['nro_referencia','NÚMERO DE REFERENCIA','t'],['cliente_final','NOMBRE CLIENTE FINAL','t'],['fecha_pedido','FECHA DE PEDIDO','d'],['bultos','N° DE BULTOS','n'],['motivo_devolucion','MOTIVO DE DEVOLUCIÓN','t'],['nodo','NODO','t'],['fecha_solicitud_cx','FECHA SOLICITUD CX','d'],['fecha_solicitud_nodo','FECHA SOLICITUD NODO','d'],['fecha_salida_nodo','FECHA SALIDA NODO','d'],['fecha_recojo_agencia','FECHA RECOJO AGENCIA (LIMA)','d'],['fecha_llegada_ctd','FECHA LLEGADA CTD','d'],['lt_retorno','LT RETORNO','n'],['medio','MEDIO','t'],['proveedor','PROVEEDOR','t'],['voucher','VOUCHER','t'],['guia','GUÍA','t'],['clave','CLAVE','t'],['estatus','ESTATUS','t']];
const rcClave = s => norm(s).replace(/[^A-Z0-9]/g, '');
const rcPorHeader = new Map(RC_COLS.map(c => [rcClave(c[1]), c]));
function leerRecojoDet(wb) {
  for (const nombre of wb.SheetNames) {
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, defval: '', raw: true });
    const hi = filas.findIndex(f => { const k = f.map(rcClave); return k.includes('NDEPEDIDO') && k.includes('NODO') && k.includes('FECHASOLICITUDCX'); });
    if (hi < 0) continue;
    const mapa = filas[hi].map(h => rcPorHeader.get(rcClave(h)) || null);
    const iFecha = filas[hi].findIndex(h => ['FECHA', 'FECHA REPORTE', 'FECHA DE REPORTE'].includes(norm(h)));   // fecha con la que se guarda cada fila
    const out = [];
    for (const f of filas.slice(hi + 1)) {
      if (!mapa.some((c, i) => c && f[i] !== '' && f[i] != null)) continue;
      const r = {}; RC_COLS.forEach(c => r[c[0]] = null);
      mapa.forEach((c, i) => { if (!c) return; const v = f[i]; r[c[0]] = c[2] === 'd' ? (fechaDe(v) || null) : valorHrn(v, c[2]); });
      r.fecha_reporte = iFecha >= 0 ? fechaDe(f[iFecha]) : '';
      if (r.nodo) r.nodo = norm(r.nodo);
      if (r.nro_pedido || r.nro_referencia) out.push(r);
    }
    if (out.length) return out;
  }
  return null;
}
// acumulable como el detalle de despacho: solo se agregan las filas que aún no están en esa fecha
async function guardarRecojoDet(filas) {
  const clave = r => [r.fecha_reporte, r.cuenta, r.nro_pedido, r.nro_referencia].map(x => x ?? '').join('|');
  const fechas = [...new Set(filas.map(r => r.fecha_reporte))].sort();
  let ya = []; for (let i = 0; i < fechas.length; i += 30) ya = ya.concat(await traerTodo(() => sb.from('recojo_detalle').select('fecha_reporte,cuenta,nro_pedido,nro_referencia').in('fecha_reporte', fechas.slice(i, i + 30)).order('id')).catch(e => { throw new Error(/recojo_detalle/.test(e.message) ? 'Falta ejecutar supabase_recojo_detalle.sql en Supabase' : e.message); }));
  const existentes = new Set(ya.map(r => clave({ ...r, fecha_reporte: String(r.fecha_reporte).slice(0, 10) }))), vistos = new Set(), nuevas = [];
  filas.forEach(r => { const k = clave(r); if (!existentes.has(k) && !vistos.has(k)) { vistos.add(k); nuevas.push(r); } });
  for (let i = 0; i < nuevas.length; i += 400) {
    const { error } = await sb.from('recojo_detalle').insert(nuevas.slice(i, i + 400)); if (error) throw error;
  }
  hrnCargado = false;
  return { nuevas: nuevas.length, repetidas: filas.length - nuevas.length, fechas };
}

// acumulable: nunca se borra lo anterior; solo se agregan las filas que aún no están en esa fecha
async function guardarHrn(filas) {
  const clave = r => r.fecha_reporte + '|' + (r.lpn ? 'L|' + r.lpn : 'P|' + [r.pedido_cliente, r.nodo, r.roll_contenedor, r.bultos].join('|'));
  const fechas = [...new Set(filas.map(r => r.fecha_reporte))].sort();
  let ya = []; for (let i = 0; i < fechas.length; i += 30) ya = ya.concat(await traerTodo(() => sb.from('hrn_detalle').select('fecha_reporte,lpn,pedido_cliente,nodo,roll_contenedor,bultos').in('fecha_reporte', fechas.slice(i, i + 30)).order('id')));
  const existentes = new Set(ya.map(r => clave({ ...r, fecha_reporte: String(r.fecha_reporte).slice(0, 10) }))), nuevas = filas.filter(r => !existentes.has(clave(r)));
  for (let i = 0; i < nuevas.length; i += 400) {
    const lote = nuevas.slice(i, i + 400);
    const { error } = await sb.from('hrn_detalle').insert(lote); memoLimpiar(); if (error) throw error;
  }
  hrnCargado = false;
  return { nuevas: nuevas.length, repetidas: filas.length - nuevas.length, fechas };
}
const hrnTexto = x => {
  const donde = x.fechas && x.fechas.length > 1 ? `en ${x.fechas.length} fechas (${fdmy(x.fechas[0])} al ${fdmy(x.fechas[x.fechas.length - 1])})` : `al ${fdmy(x.fechas[0])}`;
  return x.nuevas ? `${x.nuevas === 1 ? 'se agregó 1 fila' : 'se agregaron ' + x.nuevas + ' filas'} ${donde}` + (x.repetidas ? `; ${x.repetidas} ya ${x.repetidas === 1 ? 'estaba cargada' : 'estaban cargadas'}` : '') : `no hay filas nuevas: las ${x.repetidas} ya estaban cargadas ${donde.replace(/^al /, 'en el ')}`;
};
let hrnCargado = false, hrnVista = [], hrnMostrar = 100;
const REC_VIS = [['fecha_reporte', 'Fecha reporte'], ['nombre_cuenta', 'Nombre cuenta'], ['cuenta', 'Cuenta'], ['nro_pedido', 'Pedido'], ['cliente_final', 'Cliente final'], ['bultos', 'Bultos'], ['motivo_devolucion', 'Motivo'], ['nodo', 'Nodo'], ['fecha_recojo_agencia', 'Recojo agencia'], ['fecha_llegada_ctd', 'Llegada CTD'], ['proveedor', 'Proveedor'], ['guia', 'Guía'], ['estatus', 'Estatus']];
const detRec = () => $('#qTipo .on').dataset.t === 'rec';
const HRN_VIS = ['fecha_reporte', 'pedido_cliente', 'lpn', 'cuenta_procedencia', 'descripcion_cuenta', 'nodo', 'bultos', 'roll_contenedor', 'peso', 'volumen', 'fecha_recepcion'];
// consulta del Detalle con los filtros de la pantalla (columnas visibles para ver; todas para exportar)
function hrnConsultaQ(cols, conteo) {
  const rec = detRec();
  let q = sb.from(rec ? 'recojo_detalle' : 'hrn_detalle').select(cols, conteo ? { count: 'exact' } : undefined).order('fecha_reporte', { ascending: false }).order('nodo').order('id');
  if ($('#qDesde').value) q = q.gte('fecha_reporte', $('#qDesde').value);
  if ($('#qHasta').value) q = q.lte('fecha_reporte', $('#qHasta').value);
  if ($('#qNodo').value.trim()) q = q.ilike('nodo', '%' + $('#qNodo').value.trim() + '%');
  const t = $('#qPed').value.trim().replace(/[,()*%\\]/g, ' ').trim();
  if (t) q = q.or(rec ? `nro_pedido.ilike.%${t}%,nro_referencia.ilike.%${t}%,guia.ilike.%${t}%` : `pedido_cliente.ilike.%${t}%,lpn.ilike.%${t}%`);
  return q;
}
async function buscarHrn() {
  hrnCargado = true;
  const { data, error, count } = await hrnConsultaQ((detRec() ? REC_VIS.map(c => c[0]) : HRN_VIS).join(','), true).limit(1000);
  if (error) return flash($('#hMsg'), /recojo_detalle/.test(error.message) ? 'Falta ejecutar supabase_recojo_detalle.sql en Supabase.' : error.message, 'err');
  hrnVista = data; hrnMostrar = 100;
  $('#qInfo').textContent = `${count} filas` + (count > data.length ? ` (se muestran las primeras ${data.length}; ajusta los filtros. Al exportar salen todas)` : '');
  pintarHrn();
}
function pintarHrn() {
  const vis = hrnVista.slice(0, hrnMostrar), cols = detRec() ? REC_VIS : HRN_VIS.map((c, k) => [c, ['Fecha reporte', 'Pedido', 'LPN', 'Cuenta', 'Descripción cuenta', 'Nodo', 'Bultos', 'Roll', 'Peso', 'Volumen', 'Recepción'][k]]);
  const celda = (r, c) => c === 'fecha_recepcion' ? fdt(r[c]) : c.startsWith('fecha') ? fdmy(r[c]) : r[c];
  $('#qTabla').innerHTML = '<tr>' + cols.map(c => `<th>${c[1]}</th>`).join('') + '</tr>' +
    vis.map(r => '<tr>' + cols.map(c => `<td>${esc(celda(r, c[0]))}</td>`).join('') + '</tr>').join('') +
    (hrnVista.length > vis.length ? `<tr><td colspan="${cols.length}" style="text-align:center"><button class="lnk" id="qMas">Mostrar más (${vis.length} de ${hrnVista.length})</button></td></tr>` : '');
}
$('#qTabla').addEventListener('click', e => { if (e.target.id === 'qMas') { hrnMostrar += 200; pintarHrn(); } });
$('#qBuscar').onclick = buscarHrn;
$('#qTipo').onclick = e => {
  const b = e.target.closest('button'); if (!b) return;
  $('#qTipo').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); $('#qPedL').textContent = b.dataset.t === 'rec' ? 'Pedido / guía' : 'Pedido / LPN';
  hrnVista = []; $('#qTabla').innerHTML = ''; buscarHrn();
};
$('#qExport').onclick = async () => {
  await xlsxLib(); let todo; try { todo = await traerTodo(() => hrnConsultaQ('*', false)); } catch (e) { return flash($('#hMsg'), e.message, 'err'); }
  const rec = detRec(), cs = rec ? RC_COLS : HRN_COLS;
  const aoa = [['Fecha reporte', ...cs.map(c => c[1])]].concat(todo.map(r => [r.fecha_reporte, ...cs.map(c => r[c[0]])]));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), rec ? 'DETALLE RECOJOS' : 'DETALLE');
  XLSX.writeFile(wb, `${rec ? 'recojos' : 'hrn'}_${$('#qDesde').value || 'inicio'}_${$('#qHasta').value || 'hoy'}.xlsx`);
};

// ---------- KPIs ----------
const fmt = n => 'S/ ' + (n || 0).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtN = n => (n || 0).toLocaleString('es-PE');
const SIN_HRN = '(SIN DETALLE)', RECOJO_CTA = '(RECOJOS)';
let kpiResultado = null;

async function traerTodo(crear) {
  const pag = async n => { const { data, error, status } = await crear().range(n * 1000, n * 1000 + 999); if (error && status !== 416) throw error; return data || []; };
  const out = await pag(0);
  if (out.length < 1000) return out;   // lo normal: una sola consulta
  for (let n = 1; ; n += 4) {   // si hay más, las páginas siguientes salen de a cuatro a la vez
    const bloque = await Promise.all([0, 1, 2, 3].map(i => pag(n + i)));
    for (const d of bloque) { for (const x of d) out.push(x); if (d.length < 1000) return out; }
  }
}

// Reparte el importe de cada nodo/día entre las cuentas del detalle en proporción a "base".
function calcularKpi(envios, hrn, base) {
  const grupos = new Map();
  hrn.forEach(r => {
    const k = r.fecha_reporte + '|' + norm(r.nodo);
    const g = grupos.get(k) || new Map();
    const c = g.get(r.descripcion_cuenta || '(SIN CUENTA)') || { bultos: 0, peso: 0, volumen: 0, pedidos: 0, peds: new Set() };
    c.bultos += Number(r.bultos) || 0; c.peso += Number(r.peso) || 0; c.volumen += Number(r.volumen) || 0;
    if (r.pedido_cliente != null) c.peds.add(r.pedido_cliente); else c.pedidos += Number(r.pedidos) || 0;   // fila de detalle o fila ya agregada (vista hrn_grupos)
    g.set(r.descripcion_cuenta || '(SIN CUENTA)', c); grupos.set(k, g);
  });
  const nuevo = () => ({ costo: 0, pedidos: 0, bultos: 0 });
  const R = { recojo: 0, total: 0, pedidos: 0, bultos: 0, cajas: 0, conImporte: 0, pendientes: 0, cuentas: new Map(), nodos: new Map(), transportes: new Map(), agencias: new Map(), zonas: new Map(), cuentaNodos: new Map(), sinHrn: 0, sinHrnNodos: new Set(), sinHrnFechas: new Set() };
  const ext = (mapa, clave, e, imp, ped, bul) => {
    const a = mapa.get(clave) || { costo: 0, pedidos: 0, bultos: 0, cajas: 0, contado: 0, credito: 0, n: 0, nodos: new Set(), fac: 0 };
    a.costo += imp; a.pedidos += ped; a.bultos += bul; a.cajas += Number(e.cajas) || 0;
    if (e.modo_pago === 'CONTADO') a.contado += imp; else a.credito += imp;
    a.n++; a.nodos.add(e.nodo); if (e.factura) a.fac++; mapa.set(clave, a);
  };
  const porNodo = (cuenta, nodo, c) => { const m = R.cuentaNodos.get(cuenta) || new Map(); m.set(nodo, (m.get(nodo) || 0) + c); R.cuentaNodos.set(cuenta, m); };
  const suma = (mapa, clave, costo, pedidos, bultos) => {
    const a = mapa.get(clave) || nuevo(); a.costo += costo; a.pedidos += pedidos; a.bultos += bultos; mapa.set(clave, a);
  };
  envios.forEach(e => {
    const imp = Number(e.importe) || 0;
    if (imp <= 0) { R.pendientes++; return; }
    const ped = Number(e.pedidos) || 0, bul = Number(e.bultos) || 0;
    R.conImporte++; R.total += imp; R.pedidos += ped; R.bultos += bul; R.cajas += Number(e.cajas) || 0;
    ext(R.nodos, e.nodo, e, imp, ped, bul);
    ext(R.transportes, e.transporte || '(SIN TRANSPORTE)', e, imp, ped, bul);
    ext(R.agencias, e.agencia || '(sin agencia)', e, imp, ped, bul);
    ext(R.zonas, zonaEnv(e), e, imp, ped, bul);
    if (e.motivo === 'RECOJO') { R.recojo += imp; suma(R.cuentas, RECOJO_CTA, imp, ped, bul); porNodo(RECOJO_CTA, e.nodo, imp); return; }   // los recojos no tienen detalle por cuenta
    const g = grupos.get(e.fecha + '|' + norm(e.nodo));
    if (!g) { R.sinHrn += imp; R.sinHrnNodos.add(e.fecha + '|' + norm(e.nodo)); R.sinHrnFechas.add(e.fecha); suma(R.cuentas, SIN_HRN, imp, ped, bul); porNodo(SIN_HRN, e.nodo, imp); return; }
    const np = c => c.peds.size || c.pedidos;
    const w = c => base === 'pedidos' ? np(c) : c[base];
    let tot = [...g.values()].reduce((x, c) => x + w(c), 0);
    const peso = tot > 0 ? w : np;
    if (tot <= 0) tot = [...g.values()].reduce((x, c) => x + np(c), 0);
    g.forEach((c, cuenta) => { const cc = imp * peso(c) / tot; suma(R.cuentas, cuenta, cc, np(c), c.bultos); porNodo(cuenta, e.nodo, cc); });
  });
  return R;
}

let diasResultado = [];
function agruparDias(envios) {
  const por = new Map();
  envios.forEach(e => {
    const d = por.get(e.fecha) || { fecha: e.fecha, registros: 0, pedidos: 0, bultos: 0, cajas: 0, contado: 0, credito: 0, costo: 0, pedConImporte: 0, sinImporte: 0 };
    const imp = Number(e.importe) || 0;
    d.registros++; d.pedidos += Number(e.pedidos) || 0; d.bultos += Number(e.bultos) || 0; d.cajas += Number(e.cajas) || 0;
    if (imp > 0) {
      d.costo += imp; d.pedConImporte += Number(e.pedidos) || 0;
      if (e.modo_pago === 'CONTADO') d.contado += imp; else d.credito += imp;
    } else d.sinImporte++;
    por.set(e.fecha, d);
  });
  return [...por.values()].sort((a, b) => a.fecha.localeCompare(b.fecha));
}
function pintarDias(dias) {
  const T = dias.reduce((a, d) => { for (const k in d) if (k !== 'fecha') a[k] = (a[k] || 0) + d[k]; return a; }, {});
  const max = Math.max(1, ...dias.map(d => d.costo));
  const dia = f => new Date(f + 'T00:00:00').toLocaleDateString('es-PE', { weekday: 'short', day: '2-digit', month: '2-digit' });
  $('#kpDias').innerHTML = '<tr><th>Día</th><th class="num">Registros</th><th class="num">Pedidos</th><th class="num">Bultos</th><th class="num">Cajas</th><th class="num">Contado</th><th class="num">Crédito</th><th class="num">Costo total</th><th class="num">Costo por pedido</th><th class="num">Sin importe</th></tr>' +
    dias.map(d => `<tr data-d="${d.fecha}"><td><b>${esc(dia(d.fecha))}</b></td><td class="num">${d.registros}</td><td class="num">${fmtN(d.pedidos)}</td><td class="num">${fmtN(d.bultos)}</td><td class="num">${fmtN(d.cajas)}</td>
      <td class="num">${fmt(d.contado)}</td><td class="num">${fmt(d.credito)}</td>
      <td class="num"><b>${fmt(d.costo)}</b></td>
      <td class="num">${d.pedConImporte ? fmt(d.costo / d.pedConImporte) : '-'}</td><td class="num">${d.sinImporte || ''}</td></tr>`).join('') +
    (dias.length ? '' : '<tr><td colspan="10"><div class="empty"><b>Sin envíos en este rango</b></div></td></tr>');
  $('#nDias').textContent = `(${dias.length})`;
}
// ---------- pestañas de KPIs, nodo y transporte ----------
const KT = ['res', 'cuenta', 'nodo', 'age', 'zona', 'cal'];
function kTab(v) { v = { dia: 'res', trans: 'zona' }[v] || v; document.querySelectorAll('#kpTabs button').forEach(x => x.classList.toggle('on', x.dataset.v === v)); KT.forEach(t => $('#kp-' + t).classList.toggle('hide', t !== v)); if (KA_TABS.includes(v)) kaPintar(v); kHash(); }
$('#kpTabs').onclick = e => { const b = e.target.closest('button'); if (b) kTab(b.dataset.v); };
document.addEventListener('mouseover', e => { const el = e.target.closest && e.target.closest('#tab-kpis [data-d]'); if (el) document.querySelectorAll(`#tab-kpis [data-d="${el.dataset.d}"]`).forEach(x => x.classList.add('hi')); });
document.addEventListener('mouseout', e => { const el = e.target.closest && e.target.closest('#tab-kpis [data-d]'); if (el) document.querySelectorAll(`#tab-kpis [data-d="${el.dataset.d}"]`).forEach(x => x.classList.remove('hi')); });
const mediana = a => { const s = [...a].sort((x, y) => x - y); return s.length ? (s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2) : 0; };
// ---------- análisis por cuenta, nodo, agencia, transporte y zona ----------
const KA_TABS = ['cuenta', 'nodo', 'age', 'zona'], KA_PAL = ['#0b3d82', '#4f93e6', '#34a853', '#e8a317', '#8e8e93'], KA_COL = { g: '#248a3d', a: '#e8a317', r: '#d70015' };
const KA = {
  cuenta: { mapa: 'cuentas', unit: 'pedidos', titulo: 'Ranking de cuentas', uName: 'Pedidos', rName: 'Costo por pedido', refName: 'promedio' },
  nodo: { mapa: 'nodos', unit: 'cajas', titulo: 'Ranking de nodos', uName: 'Cajas', rName: 'Tarifa por caja', refName: 'mediana', dim: e => e.nodo },
  age: { mapa: 'agencias', unit: 'cajas', titulo: 'Ranking de agencias', uName: 'Cajas', rName: 'Tarifa por caja', refName: 'mediana', dim: e => e.agencia || '(sin agencia)' },
  zona: { mapa: 'zonas', unit: 'cajas', titulo: 'Costo por zona', uName: 'Cajas', rName: 'Tarifa por caja', refName: 'mediana', dim: zonaEnv },
  trans: { mapa: 'transportes', unit: 'cajas', uName: 'Cajas', rName: 'Tarifa por caja', refName: 'mediana', dim: e => e.transporte || '(SIN TRANSPORTE)' }
};
const KS = { cuenta: { sel: null, ord: 'costo', q: '' }, nodo: { sel: null, ord: 'costo', q: '' }, age: { sel: null, ord: 'costo', q: '' }, zona: { sel: null, ord: 'costo', q: '' } };
let kA = null, kaPaint = [];
const KA_INFO = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>';
const kaF0 = n => 'S/ ' + Math.round(n || 0).toLocaleString('es-PE');
const kaSin = k => k.startsWith('(');
const kaNom = k => k === RECOJO_CTA ? 'Recojos' : kaSin(k) ? k.charAt(1).toUpperCase() + k.slice(2, -1).toLowerCase() : k;
const kaSg = d => d > 999 ? '>999' : (d > 0 ? '+' : d < 0 ? '−' : '') + Math.abs(d).toFixed(0);
const kaVeces = (r, ref) => r >= ref * 3 ? `${(r / ref).toFixed(0)} veces la mediana` : `${((r / ref - 1) * 100).toFixed(0)} % sobre la mediana`;
const kaClase = (r, ref) => !r || !ref || r <= ref * 1.3 ? 'g' : r <= ref * 2 ? 'a' : 'r';
const kaTf = (a, b) => `<div class="f"><span>${a}</span><span>${b}</span></div>`;
const kaTipA = h => `data-tip="${esc(h)}"`;
const kaHall = t => `<div class="ka-hall">${KA_INFO}<span>${t}</span></div>`;
const kaKpis = a => `<div class="ka-card ka-kpis">${a.map(k => `<div class="ka-kp"><small>${k[0]}</small><b>${k[1]}</b><em>${k[2]}</em></div>`).join('')}</div>`;
const kaVacio = (t, d) => `<div class="ka-card ka-vacio"><b>${t}</b>${d}</div>`;
function kaDl(cur, prev) {
  if (!kA.prev.total) return '';   // sin datos del período anterior no hay con qué comparar
  if (!prev) return '<span class="ka-dl n">nuevo</span>';
  const d = (cur - prev) / prev * 100; if (Math.abs(d) < 2) return '<span class="ka-dl n">sin cambio</span>';
  return `<span class="ka-dl ${d > 0 ? 'up' : 'dn'}">${kaSg(d)} %</span>`;
}
// el rango se divide en tramos (días, o semanas si pasa de 62 días) para las tendencias
function kaTramos(desde, hasta) {
  const d0 = new Date(desde + 'T00:00:00'), d1 = new Date(hasta + 'T00:00:00');
  if (isNaN(d0) || isNaN(d1) || d1 < d0) return { n: 1, sem: false, pos: () => 0, fecha: () => desde };
  const nd = Math.round((d1 - d0) / 864e5) + 1, sem = nd > 62, t0 = new Date(d0); if (sem) t0.setDate(t0.getDate() - ((t0.getDay() + 6) % 7));
  const paso = sem ? 7 : 1, n = Math.ceil((Math.round((d1 - t0) / 864e5) + 1) / paso);
  return { n, sem, pos: f => Math.min(n - 1, Math.max(0, Math.floor(Math.round((new Date(f + 'T00:00:00') - t0) / 864e5) / paso))), fecha: i => { const t = new Date(t0); t.setDate(t.getDate() + i * paso); return isoLocal(t); } };
}
function kaSpark(b) {
  if (!b) return ''; const k = Math.min(10, b.length); if (k < 2) return '';
  const a = Array(k).fill(0); b.forEach((v, i) => a[Math.floor(i * k / b.length)] += v); if (!a.some(v => v > 0)) return '';
  const W = 78, H = 24, mx = Math.max(...a, 1), p = a.map((v, i) => [3 + i * (W - 6) / (k - 1), H - 3 - v / mx * (H - 7)]);
  return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><polyline points="${p.map(q => q.join(',')).join(' ')}" fill="none" stroke="#1a5db8" stroke-width="1.6" stroke-linejoin="round" stroke-linecap="round" opacity=".75"/><circle cx="${p.at(-1)[0]}" cy="${p.at(-1)[1]}" r="2.4" fill="#0b3d82"/></svg>`;
}
function kaSeries(dim) {
  const m = new Map();
  kA.envios.forEach(e => { const imp = Number(e.importe) || 0; if (imp <= 0) return; const k = dim(e); let a = m.get(k); if (!a) { a = Array(kA.tr.n).fill(0); m.set(k, a); } a[kA.tr.pos(e.fecha)] += imp; });
  return m;
}
function kaItems(id) {
  const c = KA[id], R = kA.R, P = kA.prev.R, pm = P && P[c.mapa], ser = c.dim ? kaSeries(c.dim) : null;
  return [...R[c.mapa]].map(([name, v]) => { const u = v[c.unit] || 0, p = pm && pm.get(name);
    return { name, costo: v.costo, contado: v.contado || 0, credito: v.credito || 0, u, pedidos: v.pedidos, cajas: v.cajas || 0, n: v.n || 0, nodos: v.nodos, r: u ? v.costo / u : 0, prev: p ? p.costo : 0, b: ser ? ser.get(name) : null }; }).sort((a, b) => b.costo - a.costo);
}
const kaRef = (id, it) => id === 'cuenta' ? (kA.R.pedidos ? kA.R.total / kA.R.pedidos : 0) : mediana(it.filter(x => x.u).map(x => x.r));
function kaDatosVista(id) { if (!kA.it[id]) { const items = kaItems(id); kA.it[id] = { items, ref: kaRef(id, items) }; } return kA.it[id]; }
function kaSub(id, x) {
  const rg = n => `${n} ${n === 1 ? 'registro' : 'registros'}`;
  if (id === 'cuenta') { const n = (kA.R.cuentaNodos.get(x.name) || new Map()).size; return `${n} ${n === 1 ? 'nodo' : 'nodos'}`; }
  if (id === 'nodo') { const z = kA.info.get(x.name)?.z; return `${z && z !== '(sin zona)' ? 'Zona ' + z : 'Sin zona'} · ${rg(x.n)}`; }
  if (id === 'age') return `${x.nodos.size} ${x.nodos.size === 1 ? 'nodo' : 'nodos'} · ${rg(x.n)}`;
  return rg(x.n);
}

// ---------- ranking en tabla ----------
function kaFilas(id, items, ref, S) {
  const pago = id !== 'cuenta', conSp = kA.act > 1 && id !== 'cuenta', cp = !!kA.prev.total, max = Math.max(...items.map(x => x.costo), 1), tot = kA.R.total || 1;
  let L = items; if (S.q) { const q = norm(S.q); L = L.filter(x => norm(kaNom(x.name)).includes(q)); }
  const f = { costo: x => x.costo, r: x => x.r, u: x => x.u, p: x => x.pedidos, d: x => x.prev ? (x.costo - x.prev) / x.prev : -9 }[S.ord];
  L = [...L].sort((a, b) => S.ord === 'nombre' ? kaNom(a.name).localeCompare(kaNom(b.name)) : f(b) - f(a));
  return L.map((x, i) => { const cl = kaClase(x.r, ref), dv = ref && x.r ? (x.r / ref - 1) * 100 : 0;
    return `<tr data-n="${esc(x.name)}" class="${id !== 'zona' ? 'sel-ok' : ''} ${id !== 'zona' && S.sel === x.name ? 'sel' : ''}"><td class="ka-ix">${i + 1}</td>
      <td class="ka-nm"><b>${esc(kaNom(x.name))}</b><small>${esc(kaSub(id, x))}</small></td>
      <td class="ka-cc"><b>${kaF0(x.costo)}</b><div class="ka-mb">${pago ? `<i class="c" style="width:${x.contado / max * 100}%"></i><i class="r" style="width:${x.credito / max * 100}%"></i>` : `<i class="c" style="width:${x.costo / max * 100}%"></i>`}</div></td>
      <td class="ka-num ka-oc">${(x.costo / tot * 100).toFixed(1)} %</td>${cp ? `<td class="ka-num ka-oc">${kaDl(x.costo, x.prev)}</td>` : ''}${id !== 'cuenta' ? `<td class="ka-num ka-oc">${fmtN(x.pedidos)}</td>` : ''}<td class="ka-num ka-oc">${fmtN(x.u)}</td>
      <td class="ka-num">${x.u ? `<span class="ka-rp ${cl}">${fmt(x.r)}</span><small class="ka-vs">${kaSg(dv)} % vs ${KA[id].refName}</small>` : '-'}</td>${conSp ? `<td class="ka-oc">${kaSpark(x.b)}</td>` : ''}</tr>`; }).join('') || `<tr><td colspan="8" style="color:var(--mut);padding:24px 8px">Sin resultados.</td></tr>`;
}
function kaRanking(id, items, ref, buscar) {
  const c = KA[id], S = KS[id], T = kA.R.total, conSp = kA.act > 1 && id !== 'cuenta', cp = !!kA.prev.total, uT = items.reduce((s, x) => s + x.u, 0);
  const th = (k, t, cls) => `<th data-s="${k}" class="${cls || ''} ${S.ord === k ? 'on' : ''}"${k === 'd' ? ' title="Variación frente al período anterior"' : ''}>${t}</th>`;
  return `<div class="ka-card"><div class="ka-cab"><div><h3>${c.titulo}</h3><p class="ka-hint">El color de la última columna compara cada fila con ${c.refName === 'promedio' ? 'el promedio' : 'la mediana'} del conjunto (${fmt(ref)}${id === 'cuenta' ? ' por pedido' : ' por caja'}).</p></div>
    ${buscar ? `<input class="ka-q" placeholder="Buscar…" value="${esc(S.q)}" aria-label="Buscar">` : ''}</div>
    <div class="ka-tw"><table class="ka-rt"><thead><tr><th></th>${th('nombre', 'Nombre')}${th('costo', 'Costo')}<th class="ka-num ka-oc">% del total</th>${cp ? th('d', 'Variación', 'ka-num ka-oc') : ''}${id !== 'cuenta' ? th('p', 'Pedidos', 'ka-num ka-oc') : ''}${th('u', c.uName, 'ka-num ka-oc')}${th('r', c.rName, 'ka-num')}${conSp ? '<th class="ka-oc">Tendencia</th>' : ''}</tr></thead>
    <tbody>${kaFilas(id, items, ref, S)}</tbody>
    <tfoot><tr><td></td><td class="l">Total</td><td>${kaF0(T)}</td><td class="ka-num ka-oc">100 %</td>${cp ? `<td class="ka-num ka-oc">${kaDl(T, kA.prev.total)}</td>` : ''}${id !== 'cuenta' ? `<td class="ka-num ka-oc">${fmtN(items.reduce((s, x) => s + x.pedidos, 0))}</td>` : ''}<td class="ka-num ka-oc">${fmtN(uT)}</td><td class="ka-num">${uT ? fmt(T / uT) : '-'}</td>${conSp ? '<td class="ka-oc"></td>' : ''}</tr></tfoot></table></div>
    <div class="ka-ley">${id !== 'cuenta' ? '<span class="c">Contado</span><span class="r">Crédito</span>' : ''}<span class="g">en rango</span><span class="a">alta (hasta el doble ${c.refName === 'promedio' ? 'del promedio' : 'de la mediana'})</span><span class="x">muy alta (más del doble)</span></div></div>`;
}

// ---------- ficha de lo seleccionado ----------
function kaFicha(id, items, ref, plano) {
  const c = KA[id], S = KS[id], it = items.find(x => x.name === S.sel) || items[0]; S.sel = it.name;
  const tot = kA.R.total || 1, dv = ref && it.r ? (it.r / ref - 1) * 100 : 0, pago = id !== 'cuenta';
  const L = kA.envios.filter(e => Number(e.importe) > 0 && c.dim && c.dim(e) === it.name);
  const top = fn => { const m = new Map(); L.forEach(e => m.set(fn(e), (m.get(fn(e)) || 0) + Number(e.importe))); return [...m].sort((a, b) => b[1] - a[1]).slice(0, 4); };
  const blq = (tit, pares) => pares.length ? `<div class="ka-sec">${tit}</div><div class="ka-bl">${pares.map(([n, v]) => `<div><span>${esc(kaNom(n))}</span><em>${kaF0(v)} · ${(v / it.costo * 100).toFixed(0)} %</em><i style="--w:${v / pares[0][1] * 100}%"></i></div>`).join('')}</div>` : '';
  let det = '';
  if (id === 'cuenta') det = blq('Nodos principales', [...(kA.R.cuentaNodos.get(it.name) || [])].sort((a, b) => b[1] - a[1]).slice(0, 4));
  else if (id === 'nodo') det = blq('Cuentas principales', [...kA.R.cuentaNodos].map(([cu, m]) => [cu, m.get(it.name) || 0]).filter(x => x[1] > 0).sort((a, b) => b[1] - a[1]).slice(0, 4)) + blq('Transporte', top(e => e.transporte || '(SIN TRANSPORTE)')) + blq('Agencias', top(e => e.agencia || '(sin agencia)'));
  else det = blq('Nodos que atiende', top(e => e.nodo)) + blq('Transporte', top(e => e.transporte || '(SIN TRANSPORTE)'));
  const dias = c.dim && kA.act > 1;
  return `<div class="${plano ? '' : 'ka-card '}ka-ficha" id="kaFicha"><span class="ka-eye">Detalle de la selección</span><div class="ka-cab" style="margin-bottom:0"><div><h3>${esc(kaNom(it.name))}</h3><p class="ka-hint">${it.p ? `Provincia de ${esc(it.p[0])}, departamento de ${esc(it.p[1])} · ` : ''}N.º ${items.indexOf(it) + 1} de ${items.length} por costo${it.n ? ` · ${it.n} ${it.n === 1 ? 'registro' : 'registros'}` : ''}</p></div>${kaDl(it.costo, it.prev)}</div>
    <div class="ka-fk"><div><small>Costo del período</small><b>${kaF0(it.costo)}</b></div><div><small>% del total</small><b>${(it.costo / tot * 100).toFixed(1)} %</b></div>
    ${id !== 'cuenta' ? `<div><small>Pedidos</small><b>${fmtN(it.pedidos)}</b></div><div><small>Costo por pedido</small><b>${it.pedidos ? fmt(it.costo / it.pedidos) : '-'}</b></div>` : ''}<div><small>${c.uName}</small><b>${fmtN(it.u)}</b></div><div><small>${c.rName}</small><b>${it.u ? fmt(it.r) : '-'}</b>${it.u ? `<em>${kaSg(dv)} % vs ${c.refName}</em>` : ''}</div></div>
    ${dias ? `<div class="ka-sec">Costo por ${kA.tr.sem ? 'semana' : 'día'}</div><div class="ka-ch" id="kaDias"></div>${pago ? `<div class="ka-ley" style="margin-top:4px"><span class="c">Contado ${kaF0(it.contado)}</span><span class="r">Crédito ${kaF0(it.credito)}</span></div>` : ''}` : ''}${det}</div>`;
}
function kaDiasDibujar(id) {
  const el = $('#kaDias'); if (!el || !el.clientWidth || !kA) return; const c = KA[id], S = KS[id], n = kA.tr.n;
  const d = Array.from({ length: n }, () => ({ c: 0, r: 0 })); kA.envios.forEach(e => { const imp = Number(e.importe) || 0; if (imp > 0 && c.dim(e) === S.sel) d[kA.tr.pos(e.fecha)][e.modo_pago === 'CONTADO' ? 'c' : 'r'] += imp; });
  const W = Math.max(240, el.clientWidth), H = 120, t = 6, b = 18, mx = Math.max(...d.map(x => x.c + x.r), 1) * 1.05, st = W / n, bw = Math.max(3, Math.min(26, st * .66)), Y = v => H - b - v / mx * (H - t - b);
  let g = `<line x1="0" x2="${W}" y1="${H - b}" y2="${H - b}" stroke="rgba(0,0,0,.12)"/>`;
  d.forEach((x, i) => { const X = i * st + (st - bw) / 2, hc = (H - b) - Y(x.c), hr = (H - b) - Y(x.r), f = kA.tr.fecha(i), ti = kA.tr.sem ? 'Semana del ' + fdmy(f) : fdmy(f);
    g += `<g class="ka-hv" opacity="1" ${kaTipA(`<b>${ti}</b>${kaTf('Contado', fmt(x.c))}${kaTf('Crédito', fmt(x.r))}`)}><rect x="${i * st}" y="${t}" width="${st}" height="${H - t - b}" fill="transparent"/>
      ${x.c > 0 ? `<rect x="${X}" y="${Y(x.c)}" width="${bw}" height="${hc}" fill="#1a5db8" rx="1.5"/>` : ''}${x.r > 0 ? `<rect x="${X}" y="${Y(x.c) - hr}" width="${bw}" height="${hr}" fill="#9dbdf0" rx="1.5"/>` : ''}</g>`; });
  const et = n === 1 ? [0] : n === 2 ? [0, 1] : [0, Math.floor((n - 1) / 2), n - 1];
  et.forEach((i, k) => g += `<text x="${k === 0 && n > 1 ? 0 : k === et.length - 1 && n > 1 ? W : i * st + st / 2}" y="${H - 4}" font-size="10.5" fill="#8e8e93" text-anchor="${k === 0 && n > 1 ? 'start' : k === et.length - 1 && n > 1 ? 'end' : 'middle'}">${fdmy(kA.tr.fecha(i)).slice(0, 5)}</text>`);
  el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">${g}</svg>`;
}

// ---------- gráficos propios de cada pestaña ----------
function kaPareto(todas) {
  const el = $('#kaPareto'); if (!el || !el.clientWidth) return; const tot = todas.reduce((s, x) => s + x.costo, 0);
  const items = todas.length > 12 ? [...todas.slice(0, 11), { name: `Otras ${todas.length - 11}`, costo: todas.slice(11).reduce((s, x) => s + x.costo, 0) }] : todas, n = items.length;
  const W = Math.max(300, el.clientWidth), H = 262, m = { l: 52, r: 36, t: 12, b: 64 }, step = (W - m.l - m.r) / n, bw = Math.min(38, step * .62), mx = items[0].costo * 1.1;
  const Y = v => H - m.b - v / mx * (H - m.t - m.b), Yp = p => H - m.b - p * (H - m.t - m.b);
  let g = '', cum = 0; const pts = [];
  [0, .5, 1].forEach(p => g += `<line x1="${m.l}" x2="${W - m.r}" y1="${Yp(p)}" y2="${Yp(p)}" stroke="rgba(0,0,0,${p ? .06 : .12})"/><text x="${W - m.r + 6}" y="${Yp(p) + 4}" font-size="10.5" fill="#8e8e93">${p * 100} %</text>`);
  g += `<line x1="${m.l}" x2="${W - m.r}" y1="${Yp(.8)}" y2="${Yp(.8)}" stroke="#a15c00" stroke-dasharray="4 4"/><text x="${W - m.r + 6}" y="${Yp(.8) + 4}" font-size="10.5" font-weight="600" fill="#a15c00">80 %</text>`;
  items.forEach((x, i) => { const antes = cum / tot; cum += x.costo; const cx = m.l + i * step + step / 2, p = cum / tot, nm = kaNom(x.name); pts.push([cx, Yp(p)]);
    g += `<g class="ka-hv" opacity="1" ${kaTipA(`<b>${esc(nm)}</b>${kaTf('Costo', fmt(x.costo))}${kaTf('Aporta', (x.costo / tot * 100).toFixed(1) + ' %')}${kaTf('Acumulado', (p * 100).toFixed(0) + ' %')}`)}><rect x="${m.l + i * step}" y="${m.t}" width="${step}" height="${H - m.t - m.b}" fill="transparent"/>
      <rect x="${cx - bw / 2}" y="${Y(x.costo)}" width="${bw}" height="${H - m.b - Y(x.costo)}" rx="6" fill="${antes < .8 ? '#1a5db8' : '#9dbdf0'}"/>
      <text transform="translate(${cx + 4},${H - m.b + 10}) rotate(-38)" font-size="11" fill="#6e6e73" text-anchor="end">${esc(nm.length > 16 ? nm.slice(0, 15) + '…' : nm)}</text></g>`; });
  g += `<polyline points="${pts.map(p => p.join(',')).join(' ')}" fill="none" stroke="#e8a317" stroke-width="2" stroke-linejoin="round" pointer-events="none"/>` + pts.map(p => `<circle cx="${p[0]}" cy="${p[1]}" r="3.4" fill="#fff" stroke="#e8a317" stroke-width="2" pointer-events="none"/>`).join('');
  el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">${g}</svg>`;
}
function kaDispersion(P, ref) {
  const el = $('#kaDisp'); if (!el || !el.clientWidth) return;
  const W = Math.max(320, el.clientWidth), H = W < 560 ? 320 : 360, m = { l: 54, r: 18, t: 16, b: 44 }, mx = Math.max(...P.map(x => x.u)) * 1.12, my = Math.max(...P.map(x => x.r), ref) * 1.18, mc = Math.max(...P.map(x => x.costo)), mu = mediana(P.map(x => x.u));
  const X = v => m.l + v / mx * (W - m.l - m.r), Y = v => H - m.b - v / my * (H - m.t - m.b);
  let g = ''; for (let i = 0; i <= 4; i++) { const v = my * i / 4; g += `<line x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}" stroke="rgba(0,0,0,${i ? .06 : .12})"/><text x="${m.l - 8}" y="${Y(v) + 4}" font-size="10.5" fill="#8e8e93" text-anchor="end">${v.toFixed(0)}</text>`; }
  for (let i = 0; i <= 4; i++) { const v = mx * i / 4; g += `<text x="${X(v)}" y="${H - 24}" font-size="10.5" fill="#8e8e93" text-anchor="middle">${Math.round(v)}</text>`; }
  g += `<rect x="${m.l}" y="${m.t}" width="${Math.max(0, X(mu) - m.l)}" height="${Math.max(0, Y(ref) - m.t)}" fill="rgba(215,0,21,.045)"/><text x="${m.l + 10}" y="${m.t + 16}" font-size="11.5" font-weight="600" fill="#d70015" opacity=".8">Bajo volumen y costo alto por pedido: revisar</text>`;
  g += `<rect x="${X(mu)}" y="${Y(ref)}" width="${Math.max(0, W - m.r - X(mu))}" height="${Math.max(0, H - m.b - Y(ref))}" fill="rgba(36,138,61,.045)"/><text x="${W - m.r - 10}" y="${H - m.b - 10}" font-size="11.5" font-weight="600" fill="#248a3d" opacity=".85" text-anchor="end">Alto volumen y bajo costo por pedido: referencia</text>`;
  g += `<line x1="${m.l}" x2="${W - m.r}" y1="${Y(ref)}" y2="${Y(ref)}" stroke="#6e6e73" stroke-dasharray="5 4"/><text x="${W - m.r}" y="${Y(ref) + 15}" font-size="11" fill="#6e6e73" text-anchor="end">promedio ${fmt(ref)}</text><line x1="${X(mu)}" x2="${X(mu)}" y1="${m.t}" y2="${H - m.b}" stroke="#6e6e73" stroke-dasharray="5 4" opacity=".6"/>`;
  g += `<text x="${(W + m.l) / 2}" y="${H - 5}" font-size="11.5" fill="#6e6e73" text-anchor="middle">Pedidos en el período</text><text transform="translate(13,${H / 2}) rotate(-90)" font-size="11.5" fill="#6e6e73" text-anchor="middle">Costo por pedido (S/)</text>`;
  // solo se rotulan las cuentas de mayor costo que no se pisen; el resto se ve al pasar el cursor
  const rad = x => 6 + Math.sqrt(x.costo / mc) * 17, rects = [], rotulo = new Set();
  [...P].sort((a, b) => b.costo - a.costo).forEach(x => { if (rotulo.size >= 8) return; const cx = X(x.u), cy = Y(x.r), w = kaNom(x.name).length * 6.4, x0 = cx > W - 170 ? cx - rad(x) - 6 - w : cx + rad(x) + 6, r = [x0, cy - 8, x0 + w, cy + 8];
    if (rects.some(o => r[0] < o[2] && r[2] > o[0] && r[1] < o[3] && r[3] > o[1])) return; rects.push(r); rotulo.add(x.name); });
  [...P].sort((a, b) => b.costo - a.costo).forEach(x => { const cl = kaClase(x.r, ref), rd = rad(x), cx = X(x.u), cy = Y(x.r), izq = cx > W - 170, nm = kaNom(x.name);
    g += `<g class="ka-hv" opacity="1" ${kaTipA(`<b>${esc(nm)}</b>${kaTf('Costo', fmt(x.costo))}${kaTf('Pedidos', fmtN(x.u))}${kaTf('Costo por pedido', fmt(x.r))}`)}><circle cx="${cx}" cy="${cy}" r="${rd}" fill="${KA_COL[cl]}" fill-opacity=".5" stroke="${KA_COL[cl]}" stroke-width="1.5"/>
      ${rotulo.has(x.name) ? `<text x="${izq ? cx - rd - 6 : cx + rd + 6}" y="${cy + 4}" font-size="11.5" fill="#1d1d1f" text-anchor="${izq ? 'end' : 'start'}">${esc(nm)}</text>` : ''}</g>`; });
  el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">${g}</svg>`;
}
function kaDesvio(items, ref, c, tit, hint) {
  const D = items.filter(x => x.u).map(x => ({ ...x, d: (x.r / ref - 1) * 100 })).sort((a, b) => b.d - a.d), mx = Math.max(...D.map(x => Math.abs(x.d)), 10);
  return `<div class="ka-card"><div class="ka-cab" style="margin-bottom:6px"><div><h3>${tit}</h3><p class="ka-hint">${hint}</p></div></div>
    ${D.map(x => { const w = Math.abs(x.d) / mx * 50, cl = x.d <= 30 ? 'g' : x.d <= 100 ? 'a' : 'r';
      return `<div class="ka-dv" ${kaTipA(`<b>${esc(kaNom(x.name))}</b>${kaTf(c.rName, fmt(x.r))}${kaTf('Mediana', fmt(ref))}${kaTf('Sobrecosto estimado', fmt(Math.max(0, x.u * (x.r - ref))))}`)}><span class="n">${esc(kaNom(x.name))}</span><div class="tr"><i class="bar ${cl}" style="left:${x.d >= 0 ? 50 : 50 - w}%;width:${w}%"></i></div><span class="v">${fmt(x.r)}<em>${kaSg(x.d)} %</em></span></div>`; }).join('')}
    <div class="ka-ley"><span class="g">por debajo o hasta 30 % sobre la mediana</span><span class="a">hasta el doble</span><span class="x">más del doble</span></div></div>`;
}
function kaLineas(items) {
  const el = $('#kaLin'); if (!el || !el.clientWidth) return; const T = items.slice(0, 5), n = kA.tr.n, top = Math.max(...T.flatMap(x => x.b)) * 1.05, raw = top / 4, pw = Math.pow(10, Math.floor(Math.log10(raw))), fr = raw / pw, paso = (fr <= 1 ? 1 : fr <= 2 ? 2 : fr <= 2.5 ? 2.5 : fr <= 5 ? 5 : 10) * pw, mx = paso * Math.ceil(top / paso);
  const W = Math.max(300, el.clientWidth), H = 250, m = { l: 48, r: 14, t: 12, b: 28 }, X = i => m.l + i * (W - m.l - m.r) / (n - 1), Y = v => H - m.b - v / mx * (H - m.t - m.b);
  let g = ''; for (let v = 0; v <= mx + 1e-6; v += paso) g += `<line x1="${m.l}" x2="${W - m.r}" y1="${Y(v)}" y2="${Y(v)}" stroke="rgba(0,0,0,${v ? .06 : .12})"/><text x="${m.l - 8}" y="${Y(v) + 4}" font-size="10.5" fill="#8e8e93" text-anchor="end">${Math.round(v).toLocaleString('es-PE')}</text>`;
  [0, Math.floor((n - 1) / 2), n - 1].filter((v, i, a) => a.indexOf(v) === i).forEach((i, k, a) => g += `<text x="${X(i)}" y="${H - 8}" font-size="10.5" fill="#8e8e93" text-anchor="${i === 0 ? 'start' : i === n - 1 ? 'end' : 'middle'}">${fdmy(kA.tr.fecha(i)).slice(0, 5)}</text>`);
  T.forEach((x, k) => { g += `<polyline points="${x.b.map((v, i) => X(i) + ',' + Y(v)).join(' ')}" fill="none" stroke="${KA_PAL[k]}" stroke-width="2.2" stroke-linejoin="round" stroke-linecap="round"/>`;
    g += x.b.map((v, i) => `<circle class="ka-hv" opacity="0" cx="${X(i)}" cy="${Y(v)}" r="6" fill="${KA_PAL[k]}" ${kaTipA(`<b>${esc(kaNom(x.name))}</b>${kaTf(fdmy(kA.tr.fecha(i)), fmt(v))}`)}/>`).join(''); });
  el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">${g}</svg>`;
}
function kaMix(items) {
  const L = [...items].filter(x => x.costo > 0).sort((a, b) => b.credito - a.credito);
  return `<div class="ka-card ka-full"><div class="ka-cab"><div><h3>Contado y crédito por agencia</h3><p class="ka-hint">Distribución del costo según modalidad de pago. Las agencias con mayor crédito concentran facturas pendientes de pago o conciliación. Ordenado por monto a crédito.</p></div></div>
    ${L.map(x => { const pc = x.contado / x.costo * 100, pr = 100 - pc;
      return `<div class="ka-mx"><div class="n">${esc(kaNom(x.name))}<small>${x.n} ${x.n === 1 ? 'registro' : 'registros'}</small></div><div class="b"><i class="ka-c-bg" style="width:${pc}%">${pc > 14 ? pc.toFixed(0) + ' %' : ''}</i><i class="r" style="width:${pr}%">${pr > 14 ? pr.toFixed(0) + ' %' : ''}</i></div><div class="t"><b>${kaF0(x.credito)}</b><small>a crédito de ${kaF0(x.costo)}</small></div></div>`; }).join('')}
    <div class="ka-ley"><span class="c">Contado</span><span class="r">Crédito</span></div></div>`;
}

// ---------- vistas ----------
function kaVistaRanking(id) {
  const c = KA[id], { items, ref } = kaDatosVista(id), R = kA.R, T = R.total, nom = items.filter(x => !kaSin(x.name)), P = kA.prev.R, hayP = !!kA.prev.total;
  let hall, kp, side = '', bottom = '', mapa = '';
  kaPaint = [() => kaDiasDibujar(id)];
  if (id === 'cuenta') {
    const t3 = nom.slice(0, 3), p3 = nom.length ? t3.reduce((s, x) => s + x.costo, 0) / T * 100 : 0, peor = [...nom.filter(x => x.u)].sort((a, b) => b.r - a.r)[0], sinD = items.filter(x => kaSin(x.name)).reduce((s, x) => s + x.costo, 0);
    const Tn = nom.reduce((s, x) => s + x.costo, 0); let ac = 0, n80 = 0; for (const x of nom) { ac += x.costo; n80++; if (Tn && ac / Tn >= .8) break; }
    hall = kaHall(`${t3.length === 1 ? 'La primera cuenta concentra' : `Las ${t3.length} primeras cuentas concentran`} el <b>${p3.toFixed(0)} %</b> del costo.${peor && nom.length >= 3 && peor.r > ref * 1.3 ? ` <b>${esc(kaNom(peor.name))}</b> registra <b>${fmt(peor.r)}</b> por pedido, <b>${((peor.r / ref - 1) * 100).toFixed(0)} %</b> sobre el promedio, con ${fmtN(peor.u)} ${peor.u === 1 ? 'pedido' : 'pedidos'}.` : ''}${sinD / T >= .1 ? ` El <b>${(sinD / T * 100).toFixed(0)} %</b> del costo aún no tiene cuenta asignada.` : ''}`);
    if (!nom.length) hall = kaHall('No se cuenta con el detalle por cuenta para este rango: el costo se muestra como <b>Sin detalle</b> hasta cargar el detalle de los nodos.');
    kp = [['Costo del período', fmt(T), kaDl(T, kA.prev.total) + (hayP ? ' frente al período anterior' : '')], ['Cuentas con costo', nom.length, Tn ? `${n80} ${n80 === 1 ? 'explica' : 'explican'} el 80 % del costo` : ''], ['Concentración en las primeras 3', p3.toFixed(0) + ' %', esc(t3.map(x => kaNom(x.name)).join(', ') || '-')],
      ['Costo por pedido', R.pedidos ? fmt(ref) : '-', (P && P.pedidos ? kaDl(ref, kA.prev.total / P.pedidos) + ' frente al período anterior' : '')]];
    if (nom.length >= 3) side = `<div class="ka-card"><div class="ka-cab" style="margin-bottom:0"><div><h3>Concentración del costo</h3><p class="ka-hint">Las barras oscuras corresponden a las cuentas que acumulan el 80 % del costo; la línea indica el porcentaje acumulado.</p></div></div><div class="ka-ch" id="kaPareto"></div></div>`;
    const PD = nom.filter(x => x.u);
    if (PD.length >= 3) bottom = `<div class="ka-card ka-full"><div class="ka-cab"><div><h3>Volumen frente a costo por pedido</h3><p class="ka-hint">Cada burbuja representa una cuenta y su tamaño, el costo total. En la zona superior izquierda se ubican las cuentas con pocos pedidos y costo unitario alto, que suelen corresponder a envíos aislados o entregas fraccionadas.</p></div></div><div class="ka-ch" id="kaDisp"></div></div>`;
    kaPaint.push(() => nom.length >= 3 && kaPareto(nom), () => PD.length >= 3 && kaDispersion(PD, ref));
  } else if (id === 'nodo') {
    const cj = items.filter(x => x.u), caro = [...cj].sort((a, b) => b.r - a.r)[0], sob = cj.reduce((s, x) => s + Math.max(0, x.u * (x.r - ref)), 0), zN = new Set([...kA.info.values()].map(x => x.z).filter(z => z && z !== '(sin zona)')).size;
    const medP = P ? mediana([...P.nodos.values()].filter(x => x.cajas).map(x => x.costo / x.cajas)) : 0;
    hall = kaHall(cj.length >= 3 && caro.r > ref * 1.3 ? `<b>${esc(kaNom(caro.name))}</b> registra una tarifa de <b>${fmt(caro.r)}</b> por caja, <b>${kaVeces(caro.r, ref)}</b>. Los nodos por encima de la mediana acumulan <b>${kaF0(sob)}</b> de sobrecosto (${(sob / T * 100).toFixed(0)} % del costo).` : cj.length ? `La tarifa mediana por caja es <b>${fmt(ref)}</b>; ningún nodo la supera en más de 30 %.` : 'No hay cajas registradas para calcular la tarifa por caja.');
    kp = [['Costo del período', fmt(T), kaDl(T, kA.prev.total) + (hayP ? ' frente al período anterior' : '')], ['Nodos con envíos', items.length, zN ? `${zN} ${zN === 1 ? 'zona' : 'zonas'}` : 'sin zona asignada'], ['Tarifa mediana por caja', ref ? fmt(ref) : '-', (medP && ref ? kaDl(ref, medP) + ' frente al período anterior' : '')], ['Sobrecosto sobre la mediana', kaF0(sob), `${(sob / T * 100).toFixed(0)} % del costo del período`]];
    if (cj.length >= 2 && ref) side = kaDesvio(items, ref, c, 'Tarifa por caja frente a la mediana', 'Diferencia porcentual entre la tarifa por caja de cada nodo y la mediana. Al pasar el cursor se muestra el sobrecosto estimado.');
    mapa = kaMapaHTML(items); kaPaint.push(() => kaMapaDibujar());
    if (kA.act > 2 && items.length >= 2) bottom = `<div class="ka-card ka-full"><div class="ka-cab"><div><h3>Evolución de los ${Math.min(5, items.length)} nodos de mayor costo</h3><p class="ka-hint">Costo por ${kA.tr.sem ? 'semana' : 'día'}. Permite distinguir un nivel de costo sostenido de un pico puntual.</p></div></div>
      <div class="ka-lg">${items.slice(0, 5).map((x, k) => `<span style="--c:${KA_PAL[k]}">${esc(kaNom(x.name))}</span>`).join('')}</div><div class="ka-ch" id="kaLin"></div></div>`;
    kaPaint.push(() => kaLineas(items));
  } else {
    const g = items[0], cr = items.reduce((s, x) => s + x.credito, 0), may = [...items].filter(x => x.costo > 0).sort((a, b) => b.credito / b.costo - a.credito / a.costo)[0];
    hall = kaHall(`<b>${esc(kaNom(g.name))}</b> concentra el <b>${(g.costo / T * 100).toFixed(0)} %</b> del costo. ${cr > 0 ? `<b>${kaF0(cr)}</b> (${(cr / T * 100).toFixed(0)} %) corresponde a operaciones a crédito${items.length > 1 ? `; <b>${esc(kaNom(may.name))}</b> presenta la mayor proporción a crédito (${(may.credito / may.costo * 100).toFixed(0)} %)` : ''}.` : 'La totalidad del costo del período corresponde a pagos al contado.'}`);
    kp = [['Costo del período', fmt(T), kaDl(T, kA.prev.total) + (hayP ? ' frente al período anterior' : '')], ['Agencias', items.length, `${R.nodos.size} ${R.nodos.size === 1 ? 'nodo atendido' : 'nodos atendidos'}`], ['Participación de la principal', (g.costo / T * 100).toFixed(0) + ' %', esc(kaNom(g.name))], ['Operaciones a crédito', kaF0(cr), `${(cr / T * 100).toFixed(0)} % del costo`]];
    if (items.filter(x => x.u).length >= 2 && ref) side = kaDesvio(items, ref, c, 'Tarifa por caja frente a la mediana', 'Compara la tarifa por caja de cada agencia con la mediana. Una diferencia significativa puede indicar una negociación pendiente o una ruta de mayor complejidad.');
    if (cr > 0) bottom = kaMix(items);
  }
  if (id === 'nodo' && mapa.includes('id="kaLz"')) return `${hall}${kaKpis(kp)}${mapa}${kaRanking(id, items, ref, true)}${side || bottom ? `<div class="ka-dos">${side}${bottom}</div>` : ''}`;
  return `${hall}${kaKpis(kp)}${mapa}<div class="ka-grid"><div class="ka-stack">${kaRanking(id, items, ref, true)}${side}</div><div class="ka-fix">${kaFicha(id, items, ref)}</div></div>${bottom}`;
}
const KA_MIN = 30;   // cajas mínimas para comparar una ruta
function kaAhorro() {
  const cel = {}; kA.envios.forEach(e => { const imp = Number(e.importe) || 0; if (imp <= 0) return; const k = (e.transporte || '(SIN TRANSPORTE)') + '|' + zonaEnv(e); (cel[k] ??= { c: 0, k: 0 }); cel[k].c += Number(e.cajas) || 0; cel[k].k += imp; });
  const trs = [...kA.R.transportes.keys()].sort(), zs = [...kA.R.zonas.keys()].sort((a, b) => kaSin(a) - kaSin(b) || a.localeCompare(b)), det = []; let tot = 0;
  zs.filter(z => !kaSin(z)).forEach(z => { const cs = trs.map(t => ({ t, ...(cel[t + '|' + z] || { c: 0, k: 0 }) })).filter(x => x.c >= KA_MIN && !kaSin(x.t)); if (cs.length < 2) return; cs.forEach(x => x.tar = x.k / x.c);
    const best = cs.reduce((a, b) => b.tar < a.tar ? b : a), worst = cs.reduce((a, b) => b.tar > a.tar ? b : a); let a = 0; cs.forEach(x => a += x.c * (x.tar - best.tar)); det.push({ z, best, worst, a }); tot += a; });
  return { cel, trs, zs, det, tot };
}
function kaVistaZona() {
  const R = kA.R, T = R.total, tr = kaItems('trans'), A = kaAhorro(), refT = R.cajas ? T / R.cajas : 0, P = kA.prev.R, hayP = !!kA.prev.total;
  const conTar = tr.filter(x => x.cajas && !kaSin(x.name)), barato = [...conTar].sort((a, b) => a.r - b.r)[0];
  const { items: zi, ref: zref } = kaDatosVista('zona'), pctA = T ? A.tot / T * 100 : 0;
  const hall = kaHall(A.det.length ? `Si cada zona operara con el transporte de menor tarifa en su ruta, el ahorro estimado sería de <b>${kaF0(A.tot)}</b> (<b>${pctA.toFixed(0)} %</b> del costo del período).${barato ? ` <b>${esc(kaNom(barato.name))}</b> presenta la menor tarifa por caja (${fmt(barato.r)}).` : ''}` : `No existen zonas con al menos dos transportes y ${KA_MIN} cajas cada uno que permitan estimar el ahorro.${barato ? ` <b>${esc(kaNom(barato.name))}</b> presenta la menor tarifa por caja (${fmt(barato.r)}).` : ''}`);
  const kp = kaKpis([['Costo del período', fmt(T), kaDl(T, kA.prev.total) + (hayP ? ' frente al período anterior' : '')], ['Tarifa promedio por caja', refT ? fmt(refT) : '-', (P && P.cajas && refT ? kaDl(refT, kA.prev.total / P.cajas) + ' frente al período anterior' : '')], ['Transporte de menor tarifa', barato ? esc(kaNom(barato.name)) : '-', barato ? `${fmt(barato.r)} por caja` : 'sin cajas registradas'], ['Ahorro potencial estimado', kaF0(A.tot), `${pctA.toFixed(0)} % del costo del período`]]);
  const cards = tr.map(x => { const cl = kaClase(x.r, refT);
    return `<div class="ka-card ka-tc"><div class="top"><h3>${esc(kaNom(x.name))}</h3><span class="ka-share">${(x.costo / T * 100).toFixed(0)} % del costo</span></div>
      <div class="big">${kaF0(x.costo)}</div>${hayP ? `<div class="row"><span>frente al período anterior</span>${kaDl(x.costo, x.prev)}</div>` : ''}
      <div class="row"><span>Tarifa por caja</span><span>${x.u ? `<span class="ka-rp ${cl}">${fmt(x.r)}</span>` : '-'}</span></div>
      <div class="row"><span>${fmtN(x.cajas)} cajas · ${x.n} ${x.n === 1 ? 'registro' : 'registros'}</span><span>${kaSpark(x.b)}</span></div>
      <div class="ka-mb"><i class="c" style="width:${x.contado / x.costo * 100}%"></i><i class="r" style="width:${x.credito / x.costo * 100}%"></i></div>
      <div class="row" style="margin-top:6px"><span>Contado <b>${(x.contado / x.costo * 100).toFixed(0)} %</b></span><span>Crédito <b>${(x.credito / x.costo * 100).toFixed(0)} %</b></span></div></div>`; }).join('');
  const ah = A.det.length ? `<div class="ka-card ka-ah"><span class="ka-eye" style="color:var(--kok)">Oportunidad de ahorro</span><h3>Ahorro potencial por zona</h3><div class="big">${kaF0(A.tot)}</div><p class="ka-hint">Costo adicional respecto de transportar las mismas cajas con el transporte de menor tarifa en cada zona.</p>
    ${[...A.det].sort((a, b) => b.a - a.a).map(d => `<div class="li"><span><b>Zona ${esc(d.z)}</b></span><b style="color:var(--kok)">${kaF0(d.a)}</b><small>${esc(kaNom(d.best.t))}: ${fmt(d.best.tar)} por caja; ${esc(kaNom(d.worst.t))}: ${fmt(d.worst.tar)} por caja</small></div>`).join('')}
    <p class="ka-aviso">Estimación referencial: no considera capacidad, plazos de entrega ni condiciones contractuales. Solo se comparan rutas con al menos ${KA_MIN} cajas.</p></div>` : '';
  const vals = []; A.trs.forEach(t => A.zs.forEach(z => { const x = A.cel[t + '|' + z]; if (x && x.c >= KA_MIN) vals.push(x.k / x.c); })); const lo = Math.min(...vals), hi = Math.max(...vals), best = {}; A.det.forEach(d => best[d.z] = d.best.t);
  const mx = A.trs.length && A.zs.length ? `<div class="ka-card ka-full"><div class="ka-cab"><div><h3>Tarifa por caja según transporte y zona</h3><p class="ka-hint">A mayor intensidad del color, mayor tarifa por caja. El borde verde identifica el transporte de menor tarifa en cada zona. Las rutas con pocas cajas no se comparan.</p></div></div>
    <div class="ka-tw"><table class="ka-hm"><thead><tr><th>Transporte</th>${A.zs.map(z => `<th>${esc(kaNom(z))}</th>`).join('')}</tr></thead><tbody>${A.trs.map(t => `<tr><td class="r">${esc(kaNom(t))}</td>${A.zs.map(z => { const x = A.cel[t + '|' + z];
      if (!x) return '<td class="c z"><b>-</b><small>sin envíos</small></td>'; const v = x.c ? x.k / x.c : 0;
      if (x.c < KA_MIN) return `<td class="c z"><b>${x.c ? fmt(v) : '-'}</b><small>${fmtN(x.c)} cajas · muestra insuficiente</small></td>`;
      const q = (v - lo) / (hi - lo || 1), a = .07 + q * .5; return `<td class="c ${best[z] === t ? 'best' : ''}" style="background:rgba(26,93,184,${a.toFixed(2)});color:${a > .34 ? '#fff' : '#1d1d1f'}" ${kaTipA(`<b>${esc(kaNom(t))} · ${esc(kaNom(z))}</b>${kaTf('Tarifa por caja', fmt(v))}${kaTf('Cajas', fmtN(x.c))}${kaTf('Costo', fmt(x.k))}`)}><b>${fmt(v)}</b><small>${fmtN(x.c)} cajas</small></td>`; }).join('')}</tr>`).join('')}</tbody></table></div></div>` : '';
  kaPaint = [];
  return `${hall}${kp}<div class="ka-tg">${cards}</div><div class="ka-grid"${ah ? '' : ' style="grid-template-columns:minmax(0,1fr)"'}><div>${kaRanking('zona', zi, zref, false)}</div>${ah ? `<div class="ka-stack">${ah}</div>` : ''}</div>${mx}`;
}

// ---------- mapa de nodos por provincia ----------
// nodos cuyo nombre no coincide con el de su provincia (clave sin tildes y en mayúsculas)
const KA_ALIAS = { AYACUCHO: 'HUAMANGA', CHIMBOTE: 'SANTA', HUACHO: 'HUAURA', MOQUEGUA: 'MARISCAL NIETO', PUCALLPA: 'CORONEL PORTILLO', TARAPOTO: 'SAN MARTIN', 'TINGO MARIA': 'LEONCIO PRADO' };
let kaGeo = null, kaGeoEstado = '';
const kaProv = new Map(), kaM = { met: 'costo', vis: 'relleno', vb: null, ar: null, mov: false };
function kaCargarGeo() {
  if (kaGeoEstado === 'cargando' || kaGeoEstado === 'ok') return; kaGeoEstado = 'cargando';
  fetch('mapa_peru.json?v=20261008g').then(r => { if (!r.ok) throw new Error(r.status); return r.json(); })
    .then(g => { kaGeo = g; g.p.forEach(p => kaProv.set(norm(p[0]), p)); kaGeoEstado = 'ok'; acDibujar(); })
    .catch(() => { kaGeoEstado = 'error'; })
    .then(() => { if (kst.modo === 'r') { if (kRecR) pintarRecojos(); } else if (document.querySelector('#kpTabs .on')?.dataset.v === 'nodo') kaPintar('nodo'); });
}
function kaUbicar(items) { items.forEach(x => { x.p = kaProv.get(norm(KA_ALIAS[norm(x.name)] || x.name)) || null; }); }
const kaMVal = x => kaM.met === 'costo' ? x.costo : kaM.met === 'pedidos' ? x.pedidos : kaM.met === 'tarifa' ? (x.u ? x.r : null) : x.u;
const kaMTxt = x => kaM.met === 'costo' ? kaF0(x.costo) : kaM.met === 'pedidos' ? fmtN(x.pedidos) + ' pedidos' : kaM.met === 'tarifa' ? fmt(x.r) : fmtN(x.u) + ' cajas';
function kaMapaHTML(items) {
  if (kaGeoEstado === 'error') return `<div class="ka-card" style="margin-bottom:18px"><div class="ka-cab" style="margin:0"><div><h3>Mapa de costos por provincia</h3><p class="ka-hint">No se pudo cargar el mapa. Actualizar la página para intentarlo de nuevo.</p></div></div></div>`;
  const cab = `<div class="ka-cab"><div><h3>Mapa de costos por provincia</h3><p class="ka-hint">Cada nodo se ubica en su provincia. Los números identifican a los nodos de mayor valor y coinciden con la lista de la derecha.</p></div>`;
  if (!kaGeo) { kaCargarGeo(); return `<div class="ka-card" style="margin-bottom:18px">${cab}</div><div class="ka-lienzo" style="cursor:default"><div class="cargando">Cargando el mapa…</div></div></div>`; }
  kaUbicar(items); const ub = items.filter(x => x.p), sin = items.filter(x => !x.p);
  if (!ub.length) return '';
  const T = kA.R.total || 1, sinC = sin.reduce((s, x) => s + x.costo, 0);
  return `<div class="ka-card" style="margin-bottom:18px">${cab}<div style="display:flex;gap:10px;flex-wrap:wrap">
      <div class="ka-seg" id="kaVis"><button data-v="relleno" class="${kaM.vis === 'relleno' ? 'on' : ''}">Provincias</button><button data-v="burbujas" class="${kaM.vis === 'burbujas' ? 'on' : ''}">Burbujas</button></div>
      <div class="ka-seg" id="kaMet"><button data-m="costo" class="${kaM.met === 'costo' ? 'on' : ''}">Costo</button><button data-m="pedidos" class="${kaM.met === 'pedidos' ? 'on' : ''}">Pedidos</button><button data-m="tarifa" class="${kaM.met === 'tarifa' ? 'on' : ''}">Tarifa por caja</button><button data-m="cajas" class="${kaM.met === 'cajas' ? 'on' : ''}">Cajas</button></div></div></div>
    <div class="ka-mapa"><div><div class="ka-lienzo" id="kaLz"><svg id="kaSvg" role="img" aria-label="Mapa del Perú por provincias"></svg>
        <div class="ka-zoom"><button id="kaZi" aria-label="Acercar">+</button><button id="kaZo" aria-label="Alejar">&minus;</button><button id="kaZr" aria-label="Restablecer vista" style="font-size:13px">&#8634;</button></div><div class="ka-leyenda" id="kaLey"></div></div>
      <p class="ka-fuente">Límites provinciales: INEI (2007), publicados en el repositorio peru-geojson.${sin.length ? ` ${sin.length} ${sin.length === 1 ? 'nodo no está ubicado' : 'nodos no están ubicados'} en el mapa (${esc(sin.map(x => kaNom(x.name)).join(', '))}; ${(sinC / T * 100).toFixed(1)} % del costo).` : ''}</p></div>
      <div id="kaMapaPanel"></div></div><div class="ka-lst" id="kaMapaListas"></div></div>`;
}
function kaMapaDibujar() {
  const lz = $('#kaLz'); if (!lz || !kaGeo || !lz.clientWidth) return;
  const svg = $('#kaSvg'); if (!kaM.vb) kaM.vb = { x: -10, y: -10, w: kaGeo.w, h: kaGeo.h };
  if (!svg.querySelector('.ka-prov')) {
    const { items } = kaDatosVista('nodo'), por = new Map(items.filter(x => x.p).map(x => [x.p[0], x]));
    svg.innerHTML = `<g id="kaGp">${kaGeo.p.map(p => { const x = por.get(p[0]);
      const tip = x ? `<b>${esc(kaNom(x.name))}</b><small>${esc(p[0])} · ${esc(p[1])}</small>${kaTf('Costo', fmt(x.costo))}${kaTf('Pedidos', fmtN(x.pedidos))}${kaTf('Cajas', fmtN(x.u))}${kaTf('Tarifa por caja', x.u ? fmt(x.r) : '-')}` : `<b>${esc(p[0])}</b><small>Sin envíos en el período</small>`;
      return `<path class="ka-prov ${x ? 'd' : ''}" data-p="${esc(p[0])}" ${x ? `data-n="${esc(x.name)}"` : ''} ${kaTipA(tip)} d="${p[4]}"/>`; }).join('')}</g><g id="kaGm"></g>`;
  }
  kaMapaActualizar('todo');
}
function kaMapaActualizar(modo = 'sel') {   // modo: 'todo' (vista completa), 'sel' (selección) o 'zoom' (solo marcadores)
  const svg = $('#kaSvg'); if (!svg || !kaGeo) return;
  const { items, ref } = kaDatosVista('nodo'), ub = items.filter(x => x.p), sel = KS.nodo.sel, por = new Map(ub.map(x => [x.p[0], x]));
  const vals = ub.map(kaMVal).filter(v => v != null), lo = Math.min(...vals), hi = Math.max(...vals);
  const color = x => { const v = x && kaMVal(x); if (v == null) return '#e6eaf0';
    if (kaM.met === 'tarifa') return { g: '#7cc593', a: '#f2c35c', r: '#e2616e' }[kaClase(x.r, ref)];
    const t = Math.sqrt((v - lo) / (hi - lo || 1)), a = [214, 228, 250], b = [11, 61, 130]; return 'rgb(' + a.map((c, i) => Math.round(c + (b[i] - c) * t)).join(',') + ')'; };
  if (modo !== 'zoom') svg.querySelectorAll('.ka-prov').forEach(el => { const x = por.get(el.dataset.p); el.setAttribute('fill', kaM.vis === 'relleno' ? color(x) : x ? '#d9e0ea' : '#e6eaf0'); el.classList.toggle('sel', !!x && x.name === sel); if (x && x.name === sel) el.parentNode.appendChild(el); });
  // marcadores: numerados (provincias) o burbujas proporcionales
  const K = kaM.vb.w / kaGeo.w, ord = ub.filter(x => kaMVal(x) != null).sort((a, b) => kaMVal(b) - kaMVal(a)), top = new Map(ord.slice(0, 8).map((x, i) => [x.name, i + 1]));
  const tr = x => `translate(${x.p[2]},${x.p[3]}) scale(${K.toFixed(3)})`, tip = x => kaTipA(`<b>${esc(kaNom(x.name))}</b><small>${esc(x.p[0])} · ${esc(x.p[1])}</small>${kaTf('Costo', fmt(x.costo))}${kaTf('Pedidos', fmtN(x.pedidos))}${kaTf('Cajas', fmtN(x.u))}${kaTf('Tarifa por caja', x.u ? fmt(x.r) : '-')}`);
  $('#kaGm').innerHTML = kaM.vis === 'burbujas'
    ? ord.map(x => { const r = 5 + Math.sqrt(kaMVal(x) / hi) * 19, n = top.get(x.name); return `<g class="ka-bb ${x.name === sel ? 'sel' : ''}" data-n="${esc(x.name)}" ${tip(x)} transform="${tr(x)}"><circle r="${r.toFixed(1)}" fill="${color(x)}" fill-opacity=".88"/>${n && r >= 11 ? `<text style="fill:${kaM.met === 'tarifa' && kaClase(x.r, ref) === 'a' ? '#1d1d1f' : '#fff'}">${n}</text>` : ''}</g>`; }).reverse().join('')
    : ord.slice(0, 8).map((x, i) => `<g class="ka-mk ${x.name === sel ? 'sel' : ''}" data-n="${esc(x.name)}" ${tip(x)} transform="${tr(x)}"><circle r="9.5"/><text>${i + 1}</text></g>`).join('');
  if (modo === 'zoom') return;
  if (modo === 'todo') svg.setAttribute('viewBox', `${kaM.vb.x} ${kaM.vb.y} ${kaM.vb.w} ${kaM.vb.h}`);
  if (modo === 'todo') $('#kaLey').innerHTML = kaM.met === 'tarifa'
    ? `<b>Tarifa por caja frente a la mediana (${fmt(ref)})</b><div class="ka-cls"><span style="--c:#7cc593">en rango</span><span style="--c:#f2c35c">alta (hasta el doble)</span><span style="--c:#e2616e">muy alta</span><span style="--c:#e6eaf0">sin envíos</span></div>`
    : `<b>${{ costo: 'Costo del período', pedidos: 'Pedidos', cajas: 'Cajas transportadas' }[kaM.met]}${kaM.vis === 'burbujas' ? ' (tamaño y color)' : ''}</b><div class="ka-grad"></div><div class="ka-gl"><span>${kaM.met === 'costo' ? kaF0(lo) : fmtN(lo)}</span><span>${kaM.met === 'costo' ? kaF0(hi) : fmtN(hi)}</span></div>`;
  kaMapaPanel(items, ub, ref, ord);
}
function kaMapaPanel(items, ub, ref, ord) {
  const S = KS.nodo, T = kA.R.total || 1, top = ord.slice(0, 8), mx = top.length ? kaMVal(top[0]) : 1, dep = new Map();
  ub.forEach(n => dep.set(n.p[1], (dep.get(n.p[1]) || 0) + n.costo)); const dl = [...dep].sort((a, b) => b[1] - a[1]).slice(0, 5);
  $('#kaMapaPanel').innerHTML = kaFicha('nodo', items, ref, true); kaDiasDibujar('nodo');
  $('#kaMapaListas').innerHTML = `<div><div class="ka-sec" style="margin-top:0">${{ costo: 'Nodos de mayor costo', pedidos: 'Nodos con más pedidos', tarifa: 'Nodos de mayor tarifa por caja', cajas: 'Nodos con más cajas' }[kaM.met]}</div>
    <div class="ka-top">${top.map((n, i) => `<button data-n="${esc(n.name)}" class="${n.name === S.sel ? 'sel' : ''}"><span class="n">${i + 1}</span><span class="nm"><b>${esc(kaNom(n.name))}</b><i style="width:${kaMVal(n) / mx * 100}%"></i></span><span class="v">${kaMTxt(n)}</span></button>`).join('')}</div></div>
    <div><div class="ka-sec" style="margin-top:0">Costo por departamento</div>
    <div class="ka-dep">${dl.map(([d, v]) => `<div><span>${esc(d)}</span><em>${kaF0(v)} · ${(v / T * 100).toFixed(0)} %</em><i style="--w:${v / dl[0][1] * 100}%"></i></div>`).join('')}</div></div>`;
}
function kaElegirNodo(n) { KS.nodo.sel = n; kaTabla('nodo'); kaMapaActualizar('sel'); }
function kaAjustar() {
  const v = kaM.vb, w0 = kaGeo.w, h0 = kaGeo.h;
  v.x = Math.min(-10 + w0 - v.w * .25, Math.max(-10 - v.w * .75, v.x)); v.y = Math.min(-10 + h0 - v.h * .25, Math.max(-10 - v.h * .75, v.y));
  $('#kaSvg').setAttribute('viewBox', `${v.x} ${v.y} ${v.w} ${v.h}`); kaMapaActualizar('zoom');
}
function kaZoom(f, cx, cy) {
  const r = $('#kaSvg').getBoundingClientRect(), v = kaM.vb, s = Math.max(v.w / r.width, v.h / r.height), px = cx ?? r.left + r.width / 2, py = cy ?? r.top + r.height / 2;
  const mx = v.x + v.w / 2 + (px - (r.left + r.width / 2)) * s, my = v.y + v.h / 2 + (py - (r.top + r.height / 2)) * s, w = Math.min(kaGeo.w, Math.max(kaGeo.w / 14, v.w * f)), k = w / v.w;
  v.x = mx - (mx - v.x) * k; v.y = my - (my - v.y) * k; v.h *= k; v.w = w; kaAjustar();
}
function kaMapaEnlazar() {
  const lz = $('#kaLz'); if (!lz) return;
  lz.addEventListener('pointerdown', e => { if (e.target.closest('.ka-zoom')) return; kaM.ar = { x: e.clientX, y: e.clientY, vx: kaM.vb.x, vy: kaM.vb.y, t: e.target.closest('.ka-prov.d, .ka-mk, .ka-bb') }; kaM.mov = false; });
  lz.addEventListener('wheel', e => { e.preventDefault(); kaZoom(e.deltaY < 0 ? .8 : 1.25, e.clientX, e.clientY); }, { passive: false });
  $('#kaZi').onclick = () => kaZoom(.7); $('#kaZo').onclick = () => kaZoom(1.4); $('#kaZr').onclick = () => { kaM.vb = { x: -10, y: -10, w: kaGeo.w, h: kaGeo.h }; kaAjustar(); };
  $('#kaVis').onclick = e => { const b = e.target.closest('button'); if (!b) return; kaM.vis = b.dataset.v; document.querySelectorAll('#kaVis button').forEach(x => x.classList.toggle('on', x === b)); kaMapaActualizar('todo'); };
  $('#kaMet').onclick = e => { const b = e.target.closest('button'); if (!b) return; kaM.met = b.dataset.m; document.querySelectorAll('#kaMet button').forEach(x => x.classList.toggle('on', x === b)); kaMapaActualizar('todo'); };
  $('#kaMapaListas').onclick = e => { const b = e.target.closest('button[data-n]'); if (b) kaElegirNodo(b.dataset.n); };
}
document.addEventListener('pointermove', e => {
  const a = kaM.ar; if (!a || !kaM.vb) return; const dx = e.clientX - a.x, dy = e.clientY - a.y;
  if (!kaM.mov && Math.abs(dx) + Math.abs(dy) > 4) { kaM.mov = true; $('#kaLz')?.classList.add('mov'); $('#kaTip').classList.remove('on'); }
  if (kaM.mov) { const r = $('#kaSvg').getBoundingClientRect(), s = Math.max(kaM.vb.w / r.width, kaM.vb.h / r.height); kaM.vb.x = a.vx - dx * s; kaM.vb.y = a.vy - dy * s; kaAjustar(); }
});
document.addEventListener('pointerup', () => {
  const a = kaM.ar; if (!a) return; const fue = kaM.mov; kaM.ar = null; kaM.mov = false; $('#kaLz')?.classList.remove('mov'); if (!fue && a.t) kaElegirNodo(a.t.dataset.n);
});

// ---------- control ----------
function kaPintar(id) {
  const el = $('#kp-' + id); if (!el) return;
  if (!kA) { el.innerHTML = ''; kaPaint = []; return; }
  if (!kA.R.conImporte) { kaPaint = []; el.innerHTML = kaVacio('Sin importes registrados en el rango seleccionado', 'Completar el importe de los registros para visualizar el análisis.'); return; }
  el.innerHTML = id === 'zona' ? kaVistaZona() : kaVistaRanking(id);
  kaEnlazar(id); kaPaint.forEach(f => f());
}
function kaEnlazar(id) {
  const el = $('#kp-' + id), rt = el.querySelector('.ka-rt'); if (!rt) return;
  rt.onclick = e => {
    const th = e.target.closest('th[data-s]'); if (th) { KS[id].ord = th.dataset.s; return kaTabla(id); }
    if (id === 'zona') return; const tr = e.target.closest('tbody tr[data-n]'); if (tr) { KS[id].sel = tr.dataset.n; kaTabla(id); if (id === 'nodo' && $('#kaLz')) kaMapaActualizar('sel'); else kaFichaRefrescar(id); }
  };
  if (id === 'nodo') kaMapaEnlazar();
  const q = el.querySelector('.ka-q'); if (q) q.oninput = () => { KS[id].q = q.value; kaTabla(id); };
}
function kaTabla(id) {
  const el = $('#kp-' + id), { items, ref } = kaDatosVista(id), S = KS[id];
  el.querySelector('.ka-rt tbody').innerHTML = kaFilas(id, items, ref, S);
  el.querySelectorAll('.ka-rt th[data-s]').forEach(t => t.classList.toggle('on', t.dataset.s === S.ord));
}
function kaFichaRefrescar(id) {
  const { items, ref } = kaDatosVista(id), f = $('#kaFicha'); if (!f) return;
  f.outerHTML = kaFicha(id, items, ref, id === 'nodo' && !!$('#kaLz')); kaDiasDibujar(id);
}
document.addEventListener('mousemove', e => {
  const tip = $('#kaTip'), t = e.target.closest && e.target.closest('#tab-kpis [data-tip]');
  if (!t || kaM.mov) { tip.classList.remove('on'); return; }
  tip.innerHTML = t.dataset.tip; tip.classList.add('on'); const w = tip.offsetWidth, h = tip.offsetHeight;
  tip.style.left = Math.min(e.clientX + 14, innerWidth - w - 8) + 'px'; tip.style.top = (e.clientY + h + 24 > innerHeight ? e.clientY - h - 12 : e.clientY + 16) + 'px';
});
{ let w0 = 0, rz; new ResizeObserver(() => { clearTimeout(rz); rz = setTimeout(() => { const w = $('#tab-kpis').clientWidth; if (w && Math.abs(w - w0) > 2) { w0 = w; kaPaint.forEach(f => f()); } }, 80); }).observe($('#tab-kpis')); }
// se llama al terminar cada cálculo de los indicadores
function kaDatos(R, envios, prev) {
  const info = new Map(); envios.forEach(e => info.set(e.nodo, { t: e.transporte, z: zonaEnv(e) }));
  kA = { R, envios, prev: { ...prev, R: prev.R || null }, tr: kaTramos($('#kDesde').value, $('#kHasta').value), info, it: {}, act: 0 };
  { const s = new Set(); envios.forEach(e => { if (Number(e.importe) > 0) s.add(kA.tr.pos(e.fecha)); }); kA.act = s.size; }   // tramos con costo: sin varios no hay tendencia que mostrar
  const v = document.querySelector('#kpTabs .on')?.dataset.v; if (KA_TABS.includes(v)) kaPintar(v);
}
// ---------- cabecera de KPIs: filtros que se aplican solos ----------
const kst = { p: 'mes', m: 'DESPACHO', modo: 'd' };   // modo d = despacho, r = recojo; el motivo sigue al modo
const isoLocal = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const ym = d => isoLocal(d).slice(0, 7);
function rangoDeMes() {
  const mes = $('#kMes').value; if (!mes) return;
  const [y, m] = mes.split('-').map(Number);
  $('#kDesde').value = mes + '-01'; $('#kHasta').value = mes + '-' + String(new Date(y, m, 0).getDate()).padStart(2, '0');
}
function kSync() {
  document.querySelectorAll('#kPer button').forEach(b => b.classList.toggle('on', b.dataset.p === kst.p || (b.dataset.p === 'mes' && ['ant', 'sel'].includes(kst.p))));
  $('#kgMes').classList.toggle('hide', !['mes', 'ant', 'sel'].includes(kst.p)); $('#kgRango').classList.toggle('hide', kst.p !== 'per');
  document.querySelectorAll('#kTransSeg button').forEach(b => b.classList.toggle('on', b.dataset.t === $('#kTrans').value));
  const r = kst.modo === 'r'; $('#tab-kpis').classList.toggle('modo-r', r); $('#kModo').classList.toggle('r', r); document.querySelectorAll('#kModo button').forEach(b => b.classList.toggle('on', (b.dataset.md === 'r') === r));
  $('#kSub').textContent = r ? 'Costo, volumen y tiempo de retorno de las devoluciones, por nodo, proveedor y cuenta' : 'Análisis del costo de transporte por cuenta, nodo, agencia, transporte y zona';
}
function kPeriodo(p) {
  kst.p = p; const hoy = new Date();
  if (p === 'mes') { $('#kMes').value = ym(hoy); rangoDeMes(); }
  else if (p === 'ant') { $('#kMes').value = ym(new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1)); rangoDeMes(); }
  else if (p === '30') { const d = new Date(hoy); d.setDate(d.getDate() - 29); $('#kDesde').value = isoLocal(d); $('#kHasta').value = isoLocal(hoy); $('#kMes').value = ''; }
  kSync(); kAuto();
}
let kTimer, kSeq = 0, kUpdTs = 0;
const kAuto = () => { clearTimeout(kTimer); kTimer = setTimeout(calcularKpis, 220); };
$('#kPer').onclick = e => { const b = e.target.closest('button'); if (b) kPeriodo(b.dataset.p); };
$('#kMes').onchange = () => { kst.p = 'sel'; rangoDeMes(); kSync(); kAuto(); };
$('#kDesde').onchange = $('#kHasta').onchange = () => { kst.p = 'per'; $('#kMes').value = ''; kSync(); kAuto(); };
$('#kModo').onclick = e => { const b = e.target.closest('button'); if (!b || b.dataset.md === kst.modo) return; kst.modo = b.dataset.md; kst.m = kst.modo === 'r' ? 'RECOJO' : 'DESPACHO'; kSync(); kHash(); kAuto(); };
$('#kCuenta').onchange = $('#kNodoR').onchange = () => { kHash(); kAuto(); };
$('#kTransSeg').onclick = e => { const b = e.target.closest('button'); if (!b) return; $('#kTrans').value = b.dataset.t; kSync(); kAuto(); };
$('#kFiltros').onclick = () => { const o = $('#tab-kpis').classList.toggle('fopen'); $('#kFiltros').setAttribute('aria-expanded', o); $('#kFiltros').textContent = o ? 'Ocultar filtros' : 'Filtros'; };
$('#kRefrescar').onclick = () => { memoLimpiar(); calcularKpis(); };
function kRestablecer() { $('#kTrans').value = ''; $('#kAgencia').value = ''; $('#kZona').value = ''; $('#kCuenta').value = ''; $('#kNodoR').value = ''; kPeriodo('mes'); }
function pintarUpd() {
  if (!kUpdTs) { $('#kUpd').textContent = ''; return; }
  const s = Math.round((Date.now() - kUpdTs) / 1000), m = Math.round(s / 60);
  $('#kUpd').textContent = s < 45 ? 'Actualizado ahora' : `Actualizado hace ${m} min`;
}
setInterval(pintarUpd, 30000);
function pintarActivos() {
  const mes = $('#kMes').value; let per;
  if (mes && ['mes', 'ant', 'sel'].includes(kst.p)) { const [y, m] = mes.split('-').map(Number), t = new Date(y, m - 1, 1).toLocaleDateString('es-PE', { month: 'long', year: 'numeric' }); per = t.charAt(0).toUpperCase() + t.slice(1); }
  else per = `${kst.p === '30' ? 'Últimos 30 días, ' : ''}del ${fdmy($('#kDesde').value)} al ${fdmy($('#kHasta').value)}`;
  const hay = $('#kTrans').value || (kst.modo === 'r' ? $('#kCuenta').value || $('#kNodoR').value : $('#kAgencia').value || $('#kZona').value);
  $('#kAct').classList.toggle('hide', !hay);
  $('#kAct').innerHTML = `<span class="hide"><span class="pill">${esc(per)}</span><span class="pill">${esc($('#kTrans').value || 'Todos los transportes')}</span>${$('#kCuenta').value && kst.modo === 'r' ? `<span class="pill">${esc($('#kCuenta').value)}</span>` : ''}${$('#kNodoR').value && kst.modo === 'r' ? `<span class="pill">Nodo ${esc($('#kNodoR').value)}</span>` : ''}${$('#kAgencia').value ? `<span class="pill">${esc($('#kAgencia').value)}</span>` : ''}${$('#kZona').value ? `<span class="pill">Zona ${esc($('#kZona').value)}</span>` : ''}</span><button class="link" id="kLimpiar">Restablecer filtros</button>`;
  $('#kLimpiar').onclick = kRestablecer;
}
function pintarTransSeg(lista) {
  const sel = $('#kTrans').value;
  $('#kTrans').innerHTML = '<option value="">Todos</option>' + lista.map(t => `<option>${esc(t)}</option>`).join(''); $('#kTrans').value = sel;
  $('#kTransSeg').innerHTML = ['', ...lista].map(t => `<button data-t="${esc(t)}">${t ? esc(t) : 'Todos'}</button>`).join(''); kSync();
}
function kPasa(e) {
  const ag = $('#kAgencia').value, zo = $('#kZona').value;
  if (kst.m && (e.motivo || 'DESPACHO') !== kst.m) return false;
  if (ag && (ag === '(sin agencia)' ? e.agencia : e.agencia !== ag)) return false;
  if (zo && zonaEnv(e) !== zo) return false;
  return true;
}
function poblarFiltros(lista) {
  const ag = $('#kAgencia').value, zo = $('#kZona').value, L = lista || [];
  const ags = [...new Set([...L.map(e => e.agencia), ag && ag !== '(sin agencia)' ? ag : ''].filter(Boolean))].sort(), zs = [...new Set([...L.map(e => e.zona), zo && zo !== '(sin zona)' ? zo : ''].filter(Boolean))].sort();
  if (!lista && $('#kAgencia').options.length > 2) return;   // sin datos nuevos se conservan las opciones
  $('#kAgencia').innerHTML = '<option value="">Todas</option>' + ags.map(a => `<option>${esc(a)}</option>`).join('') + '<option>(sin agencia)</option>'; $('#kAgencia').value = ag;
  $('#kZona').innerHTML = '<option value="">Todas</option>' + zs.map(a => `<option>${esc(a)}</option>`).join('') + '<option>(sin zona)</option>'; $('#kZona').value = zo;
}
// consultas del Dashboard que casi no cambian: se reutilizan unos minutos; se descartan al actualizar, al llegar un cambio en tiempo real o al guardar datos
const memoK = new Map();
function memo(clave, ms, fn) {
  const m = memoK.get(clave); if (m && Date.now() - m.t < ms) return m.p;
  const p = fn(); memoK.set(clave, { t: Date.now(), p }); p.catch(() => memoK.delete(clave)); return p;
}
const memoLimpiar = pref => { for (const k of [...memoK.keys()]) if (!pref || k.startsWith(pref)) memoK.delete(k); };
async function periodoPrevio(desde, hasta, trans) {
  return memo(`prev|${desde}|${hasta}|${trans}|${kst.m}|${$('#kAgencia').value}|${$('#kZona').value}`, 120000, () => periodoPrevioCalc(desde, hasta, trans));
}
async function periodoPrevioCalc(desde, hasta, trans) {
  const vacio = () => ({ total: 0, pedidos: 0, bultos: 0, cajas: 0 });
  if (!desde || !hasta) return { ...vacio(), pp: { CONTADO: vacio(), CREDITO: vacio() } };
  const d0 = new Date(desde + 'T00:00:00'), d1 = new Date(hasta + 'T00:00:00'), n = Math.round((d1 - d0) / 86400000) + 1;
  const pf = isoLocal(new Date(d0.getFullYear(), d0.getMonth(), d0.getDate() - n)), pt = isoLocal(new Date(d0.getFullYear(), d0.getMonth(), d0.getDate() - 1));
  const [rows, hrn] = await Promise.all([
    traerTodo(() => { let q = sb.from('envios').select('fecha,nodo,motivo,transporte,agencia,zona,modo_pago,pedidos,bultos,cajas,importe,factura').gte('fecha', pf).lte('fecha', pt).order('id'); if (trans) q = q.eq('transporte', trans); return q; }),
    traerHrn(pf, pt).catch(() => [])]);
  const r = { total: 0, pedidos: 0, bultos: 0, cajas: 0, pp: { CONTADO: { total: 0, pedidos: 0, bultos: 0, cajas: 0 }, CREDITO: { total: 0, pedidos: 0, bultos: 0, cajas: 0 } } };
  r.R = calcularKpi(rows.filter(kPasa), hrn, 'bultos');   // desglose por cuenta, nodo, agencia, zona y transporte del período anterior
  rows.filter(kPasa).forEach(e => { const i = Number(e.importe) || 0; if (i > 0) {
    const p = Number(e.pedidos) || 0, b = Number(e.bultos) || 0, c = Number(e.cajas) || 0, q = r.pp[e.modo_pago === 'CONTADO' ? 'CONTADO' : 'CREDITO'];
    r.total += i; r.pedidos += p; r.bultos += b; r.cajas += c; q.total += i; q.pedidos += p; q.bultos += b; q.cajas += c; } });
  return r;
}
const KG_COL = ['#1a5db8', '#5aaa2a', '#8e6ad8', '#e8912d', '#d6457a', '#2aa7b8'];
let kRes = null, kHall = [], kgTipo = 'todo', kgRank = 'nodo';
const kgNum = n => fmt(n).replace(/^S\/\s?/, '');
// variación frente al período anterior: si el costo baja es buena noticia (verde)
function kgDelta(a, b) {
  if (!a || !b) return ''; const d = (a - b) / b * 100, c = Math.abs(d) < 0.5 ? 'eq' : d > 0 ? 'up' : 'dn';
  return `<span class="kg-tag ${c}">${c === 'eq' ? 'Sin cambio' : (d > 0 ? '+' : '-') + Math.abs(d).toFixed(1) + ' %'}</span>`;
}
const KG_IC = {
  linea: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 17l6-6 4 4 8-8"/><path d="M15 7h6v6"/></svg>',
  barras: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19V9M10 19V5M16 19v-7M22 19H2"/></svg>',
  alerta: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9L2.4 18a2 2 0 001.7 3h15.8a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/></svg>',
  caja: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="7" width="18" height="13" rx="2"/><path d="M8 7V4h8v3"/></svg>',
  bulto: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8l-9-5-9 5v8l9 5 9-5V8z"/><path d="M3 8l9 5 9-5M12 13v8"/></svg>',
  sol: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2v20M17 6H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>'
};
// evolución diaria: barras por día y curva del promedio móvil de 7 días
function kgEvoSvg(dias) {
  const con = dias.filter(x => x.costo > 0); if (!con.length) return '';
  // el eje recorre todos los días del rango elegido (hasta 92); los días sin costo quedan como una marca tenue
  const mapa = new Map(con.map(x => [x.fecha, x])), d0 = new Date($('#kDesde').value + 'T00:00:00'), d1 = new Date($('#kHasta').value + 'T00:00:00');
  let d = con;
  if (!isNaN(d0) && !isNaN(d1) && d1 >= d0 && (d1 - d0) / 86400000 <= 92) { d = []; for (const t = new Date(d0); t <= d1; t.setDate(t.getDate() + 1)) { const f = isoLocal(t); d.push(mapa.get(f) || { fecha: f, costo: 0 }); } if (!con.every(x => d.includes(x))) d = con; }
  const W = 360, H = 160, pad = 6, base = H - 20, n = d.length, mx = Math.max(...d.map(x => x.costo)) * 1.06, step = (W - pad * 2) / n, bw = Math.max(3, Math.min(22, step * .68));
  const X = i => pad + i * step + (step - bw) / 2, Y = v => base - v / mx * (base - 14), imax = d.reduce((m, x, i) => x.costo > d[m].costo ? i : m, 0);
  let g = '<defs><linearGradient id="kgb" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a5db8"/><stop offset="1" stop-color="#0b3d82"/></linearGradient><linearGradient id="kgl" x1="0" x2="1"><stop offset="0" stop-color="#5aaa2a"/><stop offset="1" stop-color="#1a5db8"/></linearGradient></defs>';
  [0, .5, 1].forEach(t => { const y = base - t * (base - 14); g += `<line x1="0" x2="${W}" y1="${y}" y2="${y}" stroke="rgba(0,0,0,.06)"/>`; });
  g += d.map((x, i) => x.costo > 0
    ? `<rect x="${X(i).toFixed(1)}" y="${Y(x.costo).toFixed(1)}" width="${bw.toFixed(1)}" height="${(base - Y(x.costo)).toFixed(1)}" rx="${Math.min(7, bw / 2).toFixed(1)}" fill="${i === imax ? 'url(#kgb)' : 'rgba(26,93,184,.16)'}"><title>${fdmy(x.fecha)}, ${fmt(x.costo)}</title></rect>`
    : `<rect x="${X(i).toFixed(1)}" y="${base - 3}" width="${bw.toFixed(1)}" height="3" rx="1.5" fill="rgba(0,0,0,.08)"><title>${fdmy(x.fecha)}, sin costo</title></rect>`).join('');
  const pts = d.map((x, i) => [i, x]).filter(([, x]) => x.costo > 0);
  if (pts.length > 1) {
    const ma = pts.map((_, k) => { const w = pts.slice(Math.max(0, k - 6), k + 1); return w.reduce((a, [, x]) => a + x.costo, 0) / w.length; });
    g += `<path d="${ma.map((v, k) => (k ? 'L' : 'M') + (X(pts[k][0]) + bw / 2).toFixed(1) + ' ' + Y(v).toFixed(1)).join(' ')}" fill="none" stroke="url(#kgl)" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>`;
    g += `<circle cx="${(X(pts[pts.length - 1][0]) + bw / 2).toFixed(1)}" cy="${Y(ma[ma.length - 1]).toFixed(1)}" r="5" fill="#fff" stroke="#0b3d82" stroke-width="3"/>`;
  }
  const corto = f => fdmy(f).slice(0, 5), et = n === 1 ? [[W / 2, 'middle', d[0].fecha]] : [[pad, 'start', d[0].fecha], ...(n > 2 ? [[W / 2, 'middle', d[Math.floor((n - 1) / 2)].fecha]] : []), [W - pad, 'end', d[n - 1].fecha]];
  g += et.map(([x, a, f]) => `<text x="${x}" y="${H - 4}" font-size="10" fill="#6e6e73" text-anchor="${a}">${corto(f)}</text>`).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Costo por día">${g}</svg>`;
}
// costo por unidad según el tipo de pago elegido
function kgUnidad(envios, prev, tipo) {
  const L = envios.filter(e => Number(e.importe) > 0 && (tipo === 'todo' || (tipo === 'CONTADO' ? e.modo_pago === 'CONTADO' : e.modo_pago !== 'CONTADO')));
  const t = L.reduce((x, e) => x + Number(e.importe), 0), p = L.reduce((x, e) => x + (Number(e.pedidos) || 0), 0), b = L.reduce((x, e) => x + (Number(e.bultos) || 0), 0), c = L.reduce((x, e) => x + (Number(e.cajas) || 0), 0);
  const pv = tipo === 'todo' ? prev : (prev.pp && prev.pp[tipo]) || { total: 0, pedidos: 0, bultos: 0, cajas: 0 };
  const r = (x, y) => y ? x / y : 0;
  const liq = L.filter(regTerminado), sumLiq = liq.reduce((x, e) => x + Number(e.importe), 0);
  const base = L.filter(e => e.modo_pago === 'CONTADO').reduce((x, e) => x + Number(e.importe), 0);
  return { cpp: r(t, p), cpc: r(t, c), cpb: r(t, b), ant: r(pv.total, pv.pedidos), dpp: kgDelta(r(t, p), r(pv.total, pv.pedidos)), dpc: kgDelta(r(t, c), r(pv.total, pv.cajas)), dpb: kgDelta(r(t, b), r(pv.total, pv.bultos)), n: L.length, liq: liq.length, sumLiq, base, tot: t };
}
function kgTransportes(R) {
  const trs = [...R.transportes].sort((a, b) => b[1].costo - a[1].costo), tot = trs.reduce((x, [, v]) => x + v.costo, 0) || 1;
  const conTar = trs.filter(([k, v]) => v.cajas && !k.startsWith('(')), menor = conTar.length > 1 ? conTar.reduce((a, b) => b[1].costo / b[1].cajas < a[1].costo / a[1].cajas ? b : a)[0] : null;
  return `<div class="kg-sep"></div><h3 class="kg-t kg-t2">Distribución por transporte</h3><p class="kg-h">Participación en el costo y tarifa por caja de cada uno</p>
    <div class="kg-apil">${trs.map(([k, v], i) => `<i style="width:${v.costo / tot * 100}%;background:${KG_COL[i % KG_COL.length]}" title="${esc(k.startsWith('(') ? 'Sin transporte' : k)}"></i>`).join('')}</div>
    <div class="kg-trl">${trs.map(([k, v], i) => { const nm = k.startsWith('(') ? 'Sin transporte' : k, pct = Math.round(v.costo / tot * 100), col = KG_COL[i % KG_COL.length];
      return `<div class="kg-tr"><span class="kg-av" style="background:${col}">${esc(nm.slice(0, 2).toUpperCase())}</span><div class="t"><b><span class="kg-trn" title="${esc(nm)}">${esc(nm)}</span>${k === menor ? '<span class="kg-best">Menor tarifa</span>' : ''}</b><small>${v.n} ${v.n === 1 ? 'registro' : 'registros'}${v.cajas ? ` · ${fmt(v.costo / v.cajas)} por caja` : ''}</small></div><div class="m"><b>${fmt(v.costo)}</b><small>${pct} %</small></div></div>`; }).join('')}</div>`;
}
// una sola tarjeta: camión con profundidad (sobresale), costo por unidad y reparto por transporte
function kgFlota(R, envios, prev) {
  const U = kgUnidad(envios, prev, kgTipo), v = x => x ? fmt(x) : '-', cred = kgTipo === 'CREDITO', rg = n => `${n} ${n === 1 ? 'registro' : 'registros'}`;
  const tile = (ic, t, val, tag) => `<div class="kg-tile"><span class="t">${KG_IC[ic]}${t}</span><b>${val}</b>${tag}</div>`, pct = U.base ? Math.min(100, U.sumLiq / U.base * 100) : 0;
  return `<section class="kg-c kg-flota"><div class="kg-truck kg-stage" aria-hidden="true"><div class="kg-refl"></div><img class="kg-foto" src="camion.webp" alt=""></div>
    <div class="kg-tit"><div><h3 class="kg-t">Costo por unidad</h3><p class="kg-h">Costo unitario del transporte</p></div>
      <div class="kg-seg" role="tablist" aria-label="Tipo de pago">${[['todo', 'Todo'], ['CONTADO', 'Contado'], ['CREDITO', 'Crédito']].map(([k, t]) => `<button type="button" data-ut="${k}" class="${kgTipo === k ? 'on' : ''}">${t}</button>`).join('')}</div></div>
    <div class="kg-uh"><div class="kg-un"><small>S/</small><span>${U.cpp ? kgNum(U.cpp) : '-'}</span><em>por pedido</em></div>${U.dpp}</div>
    <p class="kg-ant">${U.ant ? `Período anterior: ${fmt(U.ant)} por pedido` : 'Sin período anterior comparable'}</p>
    <div class="kg-tiles">${tile('caja', 'Tarifa por caja', v(U.cpc), U.dpc)}${tile('bulto', 'Costo por bulto', v(U.cpb), U.dpb)}</div>
    ${cred ? `<div class="kg-liq"><div class="r1"><span class="t">Liquidación</span><b>No aplica</b></div><div class="r2"><span>${rg(U.n)} a crédito, no sujetas a liquidación</span><span>${fmt(U.tot)}</span></div></div>`
      : `<div class="kg-liq"><div class="r1"><div><span class="t">Por liquidar</span><b>${fmt(U.sumLiq)}</b></div><span class="kg-tag eq">${rg(U.liq)}</span></div><div class="kg-bar"><i style="width:${pct}%"></i></div><div class="r2"><span>${pct.toFixed(0)} % del costo al contado</span><span>de ${fmt(U.base)}</span></div></div>`}
    ${kgTransportes(R)}</section>`;
}
function kgRanking(R, envios) {
  const T = R.total, vistas = { nodo: ['Nodos', R.nodos, 'cajas'], cuenta: ['Cuentas', R.cuentas, 'pedidos'], age: ['Agencias', R.agencias, 'cajas'] }, [tit, mapa, base] = vistas[kgRank];
  const info = new Map(); envios.forEach(e => info.set(e.nodo, { t: e.transporte, z: zonaEnv(e) }));
  const lista = [...mapa].filter(([k]) => kgRank !== 'cuenta' || (k !== SIN_HRN && !k.startsWith('('))), tarifa = v => base === 'cajas' ? (v.cajas ? v.costo / v.cajas : 0) : (v.pedidos ? v.costo / v.pedidos : 0);
  const ref = kgRank === 'cuenta' ? (R.pedidos ? T / R.pedidos : 0) : mediana(lista.map(([, v]) => tarifa(v)).filter(Boolean));
  const top = lista.sort((a, b) => b[1].costo - a[1].costo).slice(0, 5), mx = Math.max(...top.map(x => x[1].costo), 1);
  const sub = (k, v) => kgRank === 'nodo' ? [info.get(k)?.t, info.get(k)?.z !== '(sin zona)' ? info.get(k)?.z : 'Sin zona'].filter(Boolean).join(' · ') : kgRank === 'cuenta' ? `${(R.cuentaNodos.get(k) || new Map()).size} nodos` : `${v.nodos.size} ${v.nodos.size === 1 ? 'nodo' : 'nodos'}`;
  const est = x => !x || !ref ? '' : x <= ref * 1.3 ? '<span class="kg-pill g">En rango</span>' : x <= ref * 2 ? '<span class="kg-pill a">Alta</span>' : '<span class="kg-pill r">Muy alta</span>';
  const hint = kgRank === 'cuenta' ? `Costo por pedido comparado con el promedio general (${fmt(ref)})` : `Tarifa por caja comparada con la mediana (${fmt(ref)})`;
  const nom = k => k.startsWith('(') ? 'Sin ' + k.slice(1, -1).toLowerCase().replace(/^sin /, '') : k;
  return `<section class="kg-c kg-rank"><div class="kg-rh"><div><h3 class="kg-t">Concentración del costo</h3><p class="kg-h">${hint}</p></div>
      <div class="kg-seg" role="tablist" aria-label="Agrupar por">${[['nodo', 'Nodos'], ['cuenta', 'Cuentas'], ['age', 'Agencias']].map(([k, t]) => `<button type="button" data-rk="${k}" class="${kgRank === k ? 'on' : ''}">${t}</button>`).join('')}</div></div>
    <div class="kg-rows"><div class="kg-row hd"><span>${tit.slice(0, -1)}</span><span>Costo</span><span class="n ped">Pedidos</span><span class="n tar">${base === 'cajas' ? 'Tarifa por caja' : 'Costo por pedido'}</span><span>Estado</span></div>
    ${top.map(([k, v]) => `<div class="kg-row"><span><b title="${esc(k)}">${esc(nom(k))}</b><small>${esc(sub(k, v))}</small></span><span class="kg-bar"><span><i style="width:${Math.max(3, v.costo / mx * 100)}%"></i></span><b>${fmt(v.costo)}</b></span><span class="n ped">${fmtN(v.pedidos)}</span><span class="n tar">${tarifa(v) ? fmt(tarifa(v)) : '-'}</span>${est(tarifa(v)) || '<span></span>'}</div>`).join('') || '<p class="kg-h">Sin datos en el rango seleccionado.</p>'}</div>
    ${lista.length > 5 ? `<button type="button" class="kg-ver" data-kt="${kgRank}">Ver detalle de ${lista.length} ${tit.toLowerCase()}</button>` : ''}</section>`;
}
function pintarResumen(R, dias, prev, envios) {
  kRes = { R, dias, prev, envios };
  if (!R.conImporte) { $('#kgrid').innerHTML = `<div class="kg-c kg-vacio"><b>Sin importes registrados en el rango seleccionado</b><span>Completar el importe de los registros para visualizar los costos.</span></div>`; $('#kgAl').innerHTML = ''; return; }
  const conta = dias.reduce((x, d) => x + d.contado, 0), cred = dias.reduce((x, d) => x + d.credito, 0), sumaCD = conta + cred, pc = sumaCD ? Math.round(conta / sumaCD * 100) : 0;
  const costos = dias.filter(d => d.costo > 0), mxd = costos.length ? costos.reduce((a, b) => b.costo > a.costo ? b : a) : null;
  const hero = `<section class="kg-c kg-hero"><h3 class="kg-t">Costo del período</h3><p class="kg-h">${R.conImporte} ${R.conImporte === 1 ? 'registro con importe' : 'registros con importe'}</p><img class="kg-arte" src="costos_arte2.webp" alt="" aria-hidden="true">
    <div class="kg-cifra"><div class="kg-n"><small>S/</small>${kgNum(R.total)}</div><div class="kg-d">${kgDelta(R.total, prev.total)}${prev.total ? `<span>frente al período anterior (${fmt(prev.total)})</span>` : ''}<span>${fmtN(R.pedidos)} pedidos</span></div></div>
    <div class="kg-mez">${sumaCD ? `${pc ? `<div class="a${pc < 25 ? ' s' : ''}" style="width:${pc}%">${pc} %<small>contado</small></div>` : ''}${pc < 100 ? `<div class="b${100 - pc < 25 ? ' s' : ''}">${100 - pc} %<small>crédito</small></div>` : ''}` : '<div class="b">Sin modalidad de pago registrada</div>'}</div>
    <div class="kg-ley"><span>Contado ${fmt(conta)}</span><span>Crédito ${fmt(cred)}</span></div>${metaHtml(R)}</section>`;
  const evo = `<section class="kg-c kg-evo"><h3 class="kg-t">Evolución diaria</h3><p class="kg-h">Costo de cada día y promedio móvil de 7 días</p><span class="kg-ico">${KG_IC.barras}</span>
    ${kgEvoSvg(dias)}${mxd ? `<div class="kg-top"><span class="p">${KG_IC.alerta}</span><div><b>${costos.length > 1 ? 'Día de mayor costo' : 'Único día con costo'}</b><small>${fdmy(mxd.fecha)} · ${fmt(mxd.costo)}</small></div>${costos.length > 1 ? `<span class="tg">${Math.round(mxd.costo / R.total * 100)} %</span>` : ''}</div>` : ''}
    </section>`;
  const alertas = kHall.length ? `<section class="kg-c kg-al"><h3 class="kg-t">Alertas y hallazgos</h3><p class="kg-h">Puntos que requieren revisión en el período</p>
    <div class="kg-alg">${kHall.slice(0, 6).map(h => `<div class="kg-hl"${h.kt ? ` data-kt="${h.kt}"` : h.go ? ` data-go="${h.go}"` : ''}><i class="${h.c}"></i><div><b>${esc(h.t)}</b><small>${esc(h.d)}</small></div></div>`).join('')}</div></section>` : '';
  $('#kgrid').innerHTML = hero + evo + kgFlota(R, envios, prev) + kgRanking(R, envios); $('#kgAl').innerHTML = alertas; kgCamion();
}
$('#kgAl').addEventListener('click', e => { const g = e.target.closest('[data-go]'); if (g) irA(g.dataset.go); });
const KG_CAM = { raf: 0 };
// camión animado con una hoja de 30 cuadros que se mezclan entre sí: el movimiento sigue la frecuencia de la pantalla y funciona igual en todos los navegadores
const KG_HOJA = { n: 30, col: 6, w: 490, h: 320, img: null, estado: '' };
function kgHoja(st) {
  const H = KG_HOJA;
  const arrancar = img => {
    if (!st.isConnected) return;
    const cv = document.createElement('canvas'); cv.width = H.w; cv.height = H.h; cv.className = 'kg-cv'; const x = cv.getContext('2d');
    if (!H.cuadros && window.createImageBitmap) { H.cuadros = []; Promise.all(Array.from({ length: H.n }, (_, k) => createImageBitmap(img, (k % H.col) * H.w, Math.floor(k / H.col) * H.h, H.w, H.h))).then(l => { H.cuadros = l; }).catch(() => { H.cuadros = null; }); }
    const cuadro = k => H.cuadros && H.cuadros.length === H.n ? x.drawImage(H.cuadros[k], 0, 0) : x.drawImage(img, (k % H.col) * H.w, Math.floor(k / H.col) * H.h, H.w, H.h, 0, 0, H.w, H.h);
    const dibujar = p => { const a = Math.floor(p), b = Math.min(a + 1, H.n - 1), f = p - a; x.clearRect(0, 0, H.w, H.h); x.globalAlpha = 1; cuadro(a); if (f > .02 && b !== a) { x.globalAlpha = f; cuadro(b); } x.globalAlpha = 1; };
    const reducir = matchMedia('(prefers-reduced-motion:reduce)').matches; let visible = true, ult = -1;
    dibujar(reducir ? 7 : 0); st.innerHTML = '<div class="kg-refl"></div>'; st.appendChild(cv); requestAnimationFrame(() => cv.classList.add('lista'));
    if (reducir) return;
    const io = new IntersectionObserver(es => { visible = es[0].isIntersecting; }); io.observe(st); const t0 = performance.now(), ciclo = 2 * 58 / 12;
    const paso = ahora => { if (!st.isConnected) return io.disconnect(); KG_CAM.raf = requestAnimationFrame(paso); if (!visible || document.hidden) return;
      const f = (((ahora - t0) / 1000) % ciclo) / ciclo, tri = f < .5 ? f * 2 : 2 - f * 2, p = (1 - Math.cos(Math.PI * tri)) / 2 * (H.n - 1);
      if (Math.abs(p - ult) > .04) { ult = p; dibujar(p); } };   // frena suavemente en los extremos
    cancelAnimationFrame(KG_CAM.raf); KG_CAM.raf = requestAnimationFrame(paso);
  };
  if (H.img) return arrancar(H.img);
  const im = new Image(); im.onload = () => { H.img = im; arrancar(im); }; im.src = 'camion_hoja2.webp';   // si no carga, queda la foto fija
}
function kgCamion() {
  cancelAnimationFrame(KG_CAM.raf);
  const st = document.querySelector('#kgrid .kg-stage'); if (st) kgHoja(st);
}
$('#kgrid').addEventListener('click', e => {
  const u = e.target.closest('[data-ut]'); if (u) { kgTipo = u.dataset.ut; if (kRes) pintarResumen(kRes.R, kRes.dias, kRes.prev, kRes.envios); return; }
  const r = e.target.closest('[data-rk]'); if (r) { kgRank = r.dataset.rk; if (kRes) pintarResumen(kRes.R, kRes.dias, kRes.prev, kRes.envios); return; }
  const g = e.target.closest('[data-go]'); if (g) irA(g.dataset.go);
});
// avisos y hallazgos: lo más grave primero
function hallazgosKpi(R, envios) {
  const out = [], T = R.total; if (!R.conImporte) return out;
  if (R.sinHrn) { const n = R.sinHrnNodos.size, f = [...R.sinHrnFechas].sort().map(x => fdmy(x).slice(0, 5)); out.push({ c: 'w', t: `${fmt(R.sinHrn)} sin cuenta asignada`, d: `Falta cargar el detalle de ${n} ${n === 1 ? 'nodo' : 'nodos'} (${f.slice(0, 4).join(', ')}${f.length > 4 ? '…' : ''}).`, go: 'detalle' }); }
  if (R.pendientes) out.push({ c: 'w', t: `${R.pendientes} ${R.pendientes === 1 ? 'registro sin importe' : 'registros sin importe'}`, d: `Se ${R.pendientes === 1 ? 'excluye' : 'excluyen'} del cálculo. Completar el importe en Registros.`, go: 'registros' });
  const rep = {}; envios.forEach(e => { const k = norm(e.factura); if (k) (rep[k] = rep[k] || []).push(e); });
  const dup = Object.values(rep).filter(l => l.length > 1);
  if (dup.length) out.push({ c: 'r', t: `Factura ${dup[0][0].factura} repetida`, d: `${[...new Set(dup[0].map(e => e.nodo))].slice(0, 3).join(', ')}${dup.length > 1 ? ` y ${dup.length - 1} más` : ''}: verificar que no se pague dos veces.`, go: 'registros' });
  const tr = [...R.transportes].filter(([, v]) => v.pedidos).map(([k, v]) => ({ k, cpp: v.costo / v.pedidos })).sort((a, b) => b.cpp - a.cpp);
  if (tr.length > 1 && tr[0].cpp / tr[tr.length - 1].cpp >= 1.3) { const f = tr[0].cpp / tr[tr.length - 1].cpp; out.push({ c: f >= 2 ? 'r' : 'w', t: `${tr[0].k} presenta un costo por pedido ${f.toFixed(1)}× superior al de ${tr[tr.length - 1].k}`, d: `${fmt(tr[0].cpp)} frente a ${fmt(tr[tr.length - 1].cpp)} por pedido.`, kt: 'trans' }); }
  const nod = [...R.nodos].filter(([, v]) => v.cajas).map(([k, v]) => ({ k, t: v.costo / v.cajas })), med = mediana(nod.map(x => x.t)), caros = nod.filter(x => nod.length >= 4 && x.t > med * 2).sort((a, b) => b.t - a.t);
  if (caros.length) out.push({ c: 'w', t: `${caros.length} ${caros.length === 1 ? 'nodo supera' : 'nodos superan'} el doble de la tarifa mediana por caja`, d: `${caros.slice(0, 3).map(x => `${x.k} ${fmt(x.t)}`).join(', ')}. La mediana es ${fmt(med)}.`, kt: 'nodo' });
  const sinAg = envios.filter(e => !e.agencia);
  if (sinAg.length) out.push({ c: 'w', t: `${sinAg.length} ${sinAg.length === 1 ? 'registro sin agencia asignada' : 'registros sin agencia asignada'}`, d: `${[...new Set(sinAg.map(e => e.nodo))].slice(0, 4).join(', ')}${new Set(sinAg.map(e => e.nodo)).size > 4 ? '…' : ''}: falta asignar la agencia en Registros.`, kt: 'cal' });
  const sinFac = envios.filter(e => !e.factura && Number(e.importe) > 0), mFac = sinFac.reduce((s, e) => s + Number(e.importe), 0);
  if (sinFac.length) out.push({ c: 'w', t: `${fmt(mFac)} (${Math.round(mFac / T * 100)} %) sin N° de factura`, d: `${sinFac.length} de ${R.conImporte} registros no cuentan con factura registrada.`, kt: 'cal' });
  const cu = [...R.cuentas].filter(([k]) => !k.startsWith('(')).sort((a, b) => b[1].costo - a[1].costo)[0];
  if (cu && cu[1].costo / T >= 0.3) out.push({ c: 'b', t: `${cu[0]} concentra el ${Math.round(cu[1].costo / T * 100)} % del costo`, d: `${fmtN(cu[1].pedidos)} de ${fmtN(R.pedidos)} pedidos del período.`, kt: 'cuenta' });
  const orden = { r: 0, w: 1, b: 2, g: 3 }; return out.sort((a, b) => orden[a.c] - orden[b.c]);
}
function calidadKpi(R, envios) {
  const n = envios.length; if (!n) { $('#calList').innerHTML = '<div class="empty"><b>Sin registros en este rango</b></div>'; $('#dotCal').className = ''; return; }
  const sinImp = envios.filter(e => !(Number(e.importe) > 0)), sinFac = envios.filter(e => !e.factura), sinAdj = envios.filter(e => !e.factura_archivo), sinAg = envios.filter(e => !e.agencia), sinZona = envios.filter(e => zonaEnv(e) === '(sin zona)');
  const item = (ic, bg, t, d, malos, go, btn) => { const pc = (1 - malos / n) * 100, col = pc >= 95 ? 'var(--grn)' : pc >= 50 ? 'var(--amb)' : 'var(--red)';
    return `<div class="cq-i"><div class="tx"><b>${t}</b><span>${d}</span></div><div class="mt"><div class="tr"><i style="width:${pc}%;background:${col}"></i></div><small>${n - malos} de ${n} completos</small></div><button class="b sec" data-go="${go}">${btn}</button></div>`; };
  const hrnPc = R.total - R.recojo > 0 ? (1 - R.sinHrn / (R.total - R.recojo)) * 100 : 100;
  $('#calList').innerHTML =
    item('', '', 'Registros con importe', 'Los registros sin importe no se incluyen en los indicadores.', sinImp.length, 'registros', 'Ir a Registros') +
    item('', '#fffbeb', 'Con N° de factura', `${fmt(sinFac.reduce((s, e) => s + (Number(e.importe) || 0), 0))} sin número de factura registrado.`, sinFac.length, 'registros', 'Completar en Registros') +
    item('', '', 'Con foto de factura adjunta', 'Permite la lectura automática del N° de factura y del importe a partir de la imagen.', sinAdj.length, 'registros', 'Adjuntar fotos') +
    item('', '#fef2f2', 'Con agencia asignada', `${esc([...new Set(sinAg.map(e => e.nodo))].join(', ') || 'Todos los nodos tienen agencia')}${sinAg.length ? ': falta asignar la agencia en Registros.' : '.'}`, sinAg.length, 'registros', 'Ir a Registros') +
    item('', '', 'Con zona', 'Requerida para el análisis por zona y la matriz transporte por zona. Se completa con la opción Editar en Registros.', sinZona.length, 'registros', 'Ir a Registros') +
    `<div class="cq-i"><div class="tx"><b>Importe repartido entre cuentas</b><span>${R.sinHrn ? `${fmt(R.sinHrn)} sin cuenta asignada: falta el detalle de ${R.sinHrnNodos.size} ${R.sinHrnNodos.size === 1 ? 'nodo' : 'nodos'}.` : 'Todo el costo está asignado a una cuenta.'}</span></div><div class="mt"><div class="tr"><i style="width:${hrnPc}%;background:${hrnPc >= 95 ? 'var(--grn)' : 'var(--amb)'}"></i></div><small>${hrnPc.toFixed(0)} % asignado</small></div><button class="b sec" data-go="detalle">Cargar detalle</button></div>`;
  $('#dotCal').className = (sinImp.length || sinFac.length || sinAdj.length || sinAg.length || sinZona.length || R.sinHrn) ? 'dotw' : '';
}
document.addEventListener('click', e => {
  const kt = e.target.closest('#tab-kpis [data-kt]'); if (kt) return kTab(kt.dataset.kt);
  const g = e.target.closest('#kp-cal [data-go]'); if (g) return irA(g.dataset.go);
});
$('#kAgencia').onchange = $('#kZona').onchange = () => { kSync(); kAuto(); };
let kpEnvios = [];
let kMeta = { aplica: false };

// ---------- HRN agregado en la base de datos (con respaldo al detalle) ----------
function traerHrn(desde, hasta) { return memo(`hrn|${desde}|${hasta}`, 300000, () => traerHrnCalc(desde, hasta)); }
async function traerHrnCalc(desde, hasta) {
  try {
    return await traerTodo(() => {
      let q = sb.from('hrn_grupos').select('fecha_reporte,nodo,descripcion_cuenta,pedidos,bultos,peso,volumen').order('fecha_reporte').order('nodo').order('descripcion_cuenta');
      if (desde) q = q.gte('fecha_reporte', desde); if (hasta) q = q.lte('fecha_reporte', hasta);
      return q;
    });
  } catch (e) {
    return await traerTodo(() => {   // la vista aún no existe: se usa el detalle completo
      let q = sb.from('hrn_detalle').select('fecha_reporte,nodo,descripcion_cuenta,pedido_cliente,bultos,peso,volumen').order('id');
      if (desde) q = q.gte('fecha_reporte', desde); if (hasta) q = q.lte('fecha_reporte', hasta);
      return q;
    });
  }
}

// ---------- meta mensual ----------
const mesNombre = m => { const [y, n] = m.split('-').map(Number), t = new Date(y, n - 1, 1).toLocaleDateString('es-PE', { month: 'long', year: 'numeric' }); return t.charAt(0).toUpperCase() + t.slice(1); };
function cargarMeta() { const mes = $('#kMes').value; return memo(`meta|${mes}`, 300000, cargarMetaCalc); }
async function cargarMetaCalc() {
  const mes = $('#kMes').value, sinFiltros = !$('#kTrans').value && !$('#kAgencia').value && !$('#kZona').value;
  if (!mes || !['mes', 'ant', 'sel'].includes(kst.p)) return { aplica: false, motivo: 'rango' };
  if (!sinFiltros) return { aplica: false, motivo: 'filtros', mes };
  const { data, error } = await sb.from('metas').select('monto').eq('mes', mes).maybeSingle();
  if (error) return { aplica: false, motivo: 'sql', mes };
  return { aplica: true, mes, monto: data ? Number(data.monto) : null };
}
function metaHtml(R) {
  const m = kMeta;
  if (!m.aplica) return m.motivo === 'filtros' ? '<div class="goal"><small>La meta mensual se evalúa sobre el total, sin filtros.</small></div>' : '';
  if (!m.monto) return `<div class="goal"><button class="link" id="metaBtn">+ Definir meta de ${esc(mesNombre(m.mes))}</button></div>`;
  const uso = R.total / m.monto * 100; let px = '';
  const hoy = new Date(), [y, n] = m.mes.split('-').map(Number);
  if (ym(hoy) === m.mes) { const dm = new Date(y, n, 0).getDate(), tr = hoy.getDate(); if (tr < dm && R.total) { const pr = R.total / tr * dm, dif = (pr / m.monto - 1) * 100;
    px = `<div class="px">Proyección al cierre del mes: <b>${fmt(pr)}</b> (${dif >= 0 ? '+' : ''}${dif.toFixed(0)} % ${dif >= 0 ? 'sobre' : 'bajo'} la meta)</div>`; } }
  return `<div class="goal"><div class="tr"><i style="width:${Math.min(100, uso)}%;background:${uso > 100 ? '#d70015' : '#1d1d1f'}"></i></div><div class="tx"><span>Meta ${fmt(m.monto)}, <button class="link" id="metaBtn">editar</button></span><span>${uso.toFixed(0)} % ejecutado</span></div>${px}</div>`;
}
function abrirMeta() {
  const m = kMeta; if (!m.mes) return; $('#metaT').textContent = 'Meta de ' + mesNombre(m.mes); $('#metaS').textContent = 'Se compara con el costo total del mes, sin filtros.';
  $('#metaV').value = m.monto || ''; $('#metaQ').classList.toggle('hide', !m.monto); $('#metaDlg').showModal(); $('#metaV').focus();
}
$('#metaX').onclick = () => $('#metaDlg').close();
$('#metaOk').onclick = async () => {
  const v = Number($('#metaV').value); if (!(v > 0)) return toast('Ingresar un monto mayor que cero', 'err');
  const { error } = await sb.from('metas').upsert({ mes: kMeta.mes, monto: v, updated_at: new Date().toISOString() });
  if (error) return toast('No se pudo guardar la meta. Verificar que se haya ejecutado supabase_fase2.sql. ' + error.message, 'err');
  $('#metaDlg').close(); toast('Meta guardada', 'ok'); memoLimpiar('meta'); calcularKpis();
};
$('#metaQ').onclick = async () => {
  const { error } = await sb.from('metas').delete().eq('mes', kMeta.mes);
  if (error) return toast(error.message, 'err'); $('#metaDlg').close(); toast('Meta quitada', 'ok'); memoLimpiar('meta'); calcularKpis();
};
document.addEventListener('click', e => { if (e.target.closest('#metaBtn')) abrirMeta(); });

// ---------- enlaces con filtros ----------
function kParams() {
  const p = new URLSearchParams(), porMes = ['mes', 'ant', 'sel'].includes(kst.p) && $('#kMes').value;
  if (porMes) p.set('mes', $('#kMes').value); else { p.set('d', $('#kDesde').value); p.set('h', $('#kHasta').value); }
  [['t', '#kTrans'], ...(kst.modo === 'r' ? [['cu', '#kCuenta'], ['no', '#kNodoR']] : [['ag', '#kAgencia'], ['zo', '#kZona']])].forEach(([k, id]) => { if ($(id).value) p.set(k, $(id).value); });
  if (kst.modo === 'r') p.set('md', 'r');
  const v = kst.modo === 'r' ? '' : document.querySelector('#kpTabs .on')?.dataset.v; if (v && v !== 'res') p.set('v', v);
  return p.toString();
}
function kHash() { if ((location.hash || '').startsWith('#kpis')) history.replaceState(null, '', '#kpis?' + kParams()); }
function kAplicar(qs) {
  const p = new URLSearchParams(qs); if (![...p.keys()].length) return;
  if (p.get('mes')) { kst.p = 'sel'; $('#kMes').value = p.get('mes'); rangoDeMes(); }
  else if (p.get('d')) { kst.p = 'per'; $('#kMes').value = ''; $('#kDesde').value = p.get('d'); $('#kHasta').value = p.get('h') || p.get('d'); }
  const t = p.get('t'); $('#kTrans').innerHTML = '<option value="">Todos</option>' + (t ? `<option>${esc(t)}</option>` : ''); $('#kTrans').value = t || '';
  $('#kAgencia').value = p.get('ag') || ''; $('#kZona').value = p.get('zo') || ''; kst.modo = p.get('md') === 'r' || p.get('mo') === 'RECOJO' ? 'r' : 'd'; kst.m = kst.modo === 'r' ? 'RECOJO' : 'DESPACHO';
  [['#kCuenta', p.get('cu'), 'Todas'], ['#kNodoR', p.get('no'), 'Todos']].forEach(([id, v, t]) => { $(id).innerHTML = `<option value="">${t}</option>` + (v ? `<option>${esc(v)}</option>` : ''); $(id).value = v || ''; });
  kSync(); if (KT.includes(p.get('v'))) kTab(p.get('v'));
}
$('#kLink').onclick = () => { history.replaceState(null, '', '#kpis?' + kParams()); (navigator.clipboard ? navigator.clipboard.writeText(location.href) : Promise.reject()).then(() => toast('Enlace copiado, con los filtros aplicados', 'ok'), () => toast('No se pudo copiar el enlace', 'err')); };

// ---------- costo por período: semanas, meses, trimestres y años ----------
const KP_MES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'set', 'oct', 'nov', 'dic'];
const KP_N = { sem: ['semana', 'semanas', 'Semana de mayor costo'], mes: ['mes', 'meses', 'Mes de mayor costo'], tri: ['trimestre', 'trimestres', 'Trimestre de mayor costo'], anio: ['año', 'años', 'Año de mayor costo'] };
const KP_T = { sem: 'Últimas 12 semanas', mes: 'Últimos 12 meses', tri: 'Últimos 6 trimestres', anio: 'Últimos 5 años' };
let kpVista = 'mes', kpFilas = [], kpSeq = 0; const kpCache = new Map();
const kpYMD = f => String(f).slice(0, 10).split('-').map(Number), kp2 = n => String(n).padStart(2, '0');
function kpBuckets(v) {   // del más antiguo al período en curso
  const [y, m, d] = kpYMD(isoLocal(new Date())), out = [];
  if (v === 'sem') { const l0 = new Date(y, m - 1, d); l0.setDate(l0.getDate() - ((l0.getDay() + 6) % 7)); for (let i = 11; i >= 0; i--) { const t = new Date(l0); t.setDate(t.getDate() - 7 * i); out.push({ k: isoLocal(t), l: `${t.getDate()} ${KP_MES[t.getMonth()]}`, desde: isoLocal(t) }); } }
  else if (v === 'mes') for (let i = 11; i >= 0; i--) { const t = new Date(y, m - 1 - i, 1), k = `${t.getFullYear()}-${kp2(t.getMonth() + 1)}`; out.push({ k, l: `${KP_MES[t.getMonth()]} ${String(t.getFullYear()).slice(2)}`, desde: k + '-01' }); }
  else if (v === 'tri') { const q = Math.floor((m - 1) / 3); for (let i = 5; i >= 0; i--) { const t = new Date(y, (q - i) * 3, 1), qq = Math.floor(t.getMonth() / 3) + 1; out.push({ k: `${t.getFullYear()}-T${qq}`, l: `T${qq} ${String(t.getFullYear()).slice(2)}`, desde: `${t.getFullYear()}-${kp2(t.getMonth() + 1)}-01` }); } }
  else for (let i = 4; i >= 0; i--) out.push({ k: String(y - i), l: String(y - i), desde: `${y - i}-01-01` });
  out[out.length - 1].enCurso = true; return out;
}
function kpClave(v, f) {
  const [y, m, d] = kpYMD(f);
  if (v === 'sem') { const t = new Date(y, m - 1, d); t.setDate(t.getDate() - ((t.getDay() + 6) % 7)); return isoLocal(t); }
  return v === 'mes' ? `${y}-${kp2(m)}` : v === 'tri' ? `${y}-T${Math.floor((m - 1) / 3) + 1}` : String(y);
}
async function kpCargar() {
  const v = kpVista, mi = ++kpSeq, B = kpBuckets(v), trans = $('#kTrans').value, clave = [v, trans, kst.m, $('#kAgencia').value, $('#kZona').value].join('|');
  $('#perCard').classList.add('cargando');
  try {
    let rows = kpCache.get(clave);
    if (!rows) {
      rows = (await traerTodo(() => { let q = sb.from('envios').select('fecha,motivo,agencia,zona,modo_pago,importe').gt('importe', 0).gte('fecha', B[0].desde).order('id'); if (trans) q = q.eq('transporte', trans); return q; })).filter(kPasa);
      kpCache.set(clave, rows);
    }
    if (mi !== kpSeq) return;
    const por = new Map(B.map(b => [b.k, { l: b.l, c: 0, r: 0, enCurso: !!b.enCurso }]));
    rows.forEach(e => { const x = por.get(kpClave(v, e.fecha)); if (!x) return; const i = Number(e.importe) || 0; if (e.modo_pago === 'CONTADO') x.c += i; else x.r += i; });
    let f = [...por.values()].map(x => ({ ...x, t: x.c + x.r }));
    if (v === 'anio') { const i = f.findIndex(x => x.t > 0); f = i < 0 ? f.slice(-1) : f.slice(i); }
    kpFilas = f; kpPintar();
  } catch (err) { if (mi === kpSeq) kpMsg(`No se pudo cargar el gráfico (${err.message}).`); }
  if (mi === kpSeq) $('#perCard').classList.remove('cargando');
}
function kpMsg(t) { const b = $('#kpBox'); b.querySelector('svg')?.remove(); b.querySelector('.kc-vacio')?.remove(); b.insertAdjacentHTML('afterbegin', `<div class="kc-vacio">${esc(t)}</div>`); }
function kpPintar() {
  const v = kpVista, rows = kpFilas, n = rows.length, N = KP_N[v], tot = rows.reduce((a, x) => a + x.t, 0);
  $('#kpHint').textContent = KP_T[v] + '. Se aplican los filtros seleccionados de transporte, motivo, agencia y zona.';
  document.querySelectorAll('#kpVista button').forEach(b => b.classList.toggle('on', b.dataset.v === v));
  if (!tot) { $('#kpKpis').innerHTML = ''; $('#kpProm').textContent = ''; kpMsg('Sin costos registrados en este período.'); return; }
  const cer = rows.filter(x => !x.enCurso), base = cer.length ? cer : rows, prom = base.reduce((a, x) => a + x.t, 0) / base.length, mx = rows.reduce((m, x) => x.t > m.t ? x : m, rows[0]);
  const conta = Math.round(rows.reduce((a, x) => a + x.c, 0) / tot * 100);
  const [a1, b1] = [cer[cer.length - 1], cer[cer.length - 2]], dl = a1 && b1 && b1.t > 0 ? (a1.t - b1.t) / b1.t * 100 : null;
  $('#kpKpis').innerHTML =
    `<div class="kc-k"><small>Total del período graficado</small><b>${fmt(tot)}</b><em>${n} ${n === 1 ? N[0] : N[1]}</em></div>` +
    `<div class="kc-k"><small>Promedio por ${N[0]}</small><b>${fmt(prom)}</b><em>${conta} % al contado</em></div>` +
    `<div class="kc-k"><small>${N[2]}</small><b>${fmt(mx.t)}</b><em>${esc(mx.l)}</em></div>` +
    (dl == null ? `<div class="kc-k"><small>Variación del último período cerrado</small><b>-</b><em>sin períodos comparables</em></div>`
      : `<div class="kc-k"><small>Variación del último período cerrado</small><b>${dl > 0 ? '+' : ''}${dl.toFixed(1)} %<span class="kg-tag ${Math.abs(dl) < .5 ? 'eq' : dl > 0 ? 'up' : 'dn'}">${Math.abs(dl) < .5 ? 'sin variación' : dl > 0 ? 'aumento' : 'reducción'}</span></b><em>${esc(a1.l)} frente a ${esc(b1.l)}</em></div>`);
  $('#kpProm').textContent = 'Promedio ' + fmt(prom);
  kpDibujar(rows, prom);
}
function kpDibujar(rows, prom) {
  const box = $('#kpBox'), tip = $('#kpTip'); if (!box || !rows.length) return;
  const W = Math.max(280, box.clientWidth), H = W < 520 ? 250 : 300, m = { l: 52, r: 8, t: 14, b: 30 }, n = rows.length, f0 = x => Math.round(x).toLocaleString('es-PE');
  const top = Math.max(...rows.map(x => x.t)) * 1.06, raw = top / 4, pw = Math.pow(10, Math.floor(Math.log10(raw))), fr = raw / pw, paso = (fr <= 1 ? 1 : fr <= 2 ? 2 : fr <= 2.5 ? 2.5 : fr <= 5 ? 5 : 10) * pw, nt = Math.ceil(top / paso), nice = paso * nt;
  const step = (W - m.l - m.r) / n, bw = Math.min(54, step * .62), X = i => m.l + i * step + (step - bw) / 2, Y = v => H - m.b - v / nice * (H - m.t - m.b), alto = H - m.t - m.b;
  let g = '<defs><linearGradient id="kpa" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a5db8"/><stop offset="1" stop-color="#0b3d82"/></linearGradient></defs>';
  for (let i = 0; i <= nt; i++) { const v = paso * i, y = Y(v); g += `<line x1="${m.l}" x2="${W - m.r}" y1="${y}" y2="${y}" stroke="rgba(0,0,0,${i ? .06 : .12})"/><text x="${m.l - 8}" y="${y + 4}" font-size="11" fill="#8e8e93" text-anchor="end">${v >= 1000 ? (v / 1000).toLocaleString('es-PE') + ' mil' : f0(v)}</text>`; }
  const salto = Math.max(1, Math.ceil(46 / step));   // en pantallas angostas se muestran solo algunas etiquetas
  rows.forEach((x, i) => {
    const hc = x.c / nice * alto, hr = x.r / nice * alto, yc = H - m.b - hc, yr = yc - hr, ult = i === n - 1, rad = Math.min(10, bw / 2), xi = X(i);
    g += `<g class="col" data-i="${i}"><rect x="${m.l + i * step}" y="${m.t}" width="${step}" height="${alto}" fill="transparent"/>
      <rect class="hv" x="${m.l + i * step + 3}" y="${m.t}" width="${step - 6}" height="${alto}" rx="12" fill="rgba(26,93,184,.07)" opacity="0"/>
      ${hc > 0 ? `<rect x="${xi}" y="${yc}" width="${bw}" height="${hc}" fill="url(#kpa)" opacity="${ult ? 1 : .86}"/>` : ''}
      ${hr > 0 ? `<path d="M${xi} ${yc} V${yr + (hc > 0 ? rad : 0)} ${hc > 0 ? `Q${xi} ${yr} ${xi + rad} ${yr} H${xi + bw - rad} Q${xi + bw} ${yr} ${xi + bw} ${yr + rad}` : `H${xi + bw}`} V${yc} H${xi} Z" fill="#9dbdf0" opacity="${ult ? 1 : .9}"/>` : ''}
      ${(n - 1 - i) % salto === 0 ? `<text x="${xi + bw / 2}" y="${H - 9}" font-size="11.5" fill="${ult ? '#1d1d1f' : '#6e6e73'}" font-weight="${ult ? 600 : 400}" text-anchor="middle">${esc(x.l)}</text>` : ''}</g>`;
  });
  g += `<line x1="${m.l}" x2="${W - m.r}" y1="${Y(prom)}" y2="${Y(prom)}" stroke="#8e8e93" stroke-width="1.5" stroke-dasharray="5 5"/>`;
  box.querySelector('svg')?.remove(); box.querySelector('.kc-vacio')?.remove(); box.insertAdjacentHTML('afterbegin', `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="Costo por período">${g}</svg>`);
  box.querySelectorAll('.col').forEach(c => {
    const i = +c.dataset.i, x = rows[i], p = rows[i - 1], d = p && p.t > 0 && !x.enCurso ? (x.t - p.t) / p.t * 100 : null;
    const mostrar = () => { tip.innerHTML = `<b>${esc(x.l)}</b><div class="f"><span>Contado</span><span>${fmt(x.c)}</span></div><div class="f"><span>Crédito</span><span>${fmt(x.r)}</span></div><hr><div class="f"><span>Total</span><span><b style="display:inline;margin:0">${fmt(x.t)}</b></span></div>${d == null ? '' : `<div class="f"><span>Variación frente al anterior</span><span class="d ${d > 0 ? 'up' : 'dn'}">${d > 0 ? '+' : ''}${d.toFixed(1)} %</span></div>`}${x.enCurso ? '<i>Período en curso: información parcial</i>' : ''}`;
      const cx = m.l + i * step + step / 2, w = tip.offsetWidth || 180, left = Math.min(Math.max(0, cx - w / 2), W - w);
      tip.style.left = left + 'px'; tip.style.top = Math.max(0, Y(x.t) - tip.offsetHeight - 10) + 'px'; tip.classList.add('on'); };
    c.addEventListener('mouseenter', mostrar); c.addEventListener('click', mostrar); c.addEventListener('mouseleave', () => tip.classList.remove('on'));
  });
}
$('#kpVista').onclick = e => { const b = e.target.closest('button'); if (!b || b.dataset.v === kpVista) return; kpVista = b.dataset.v; kpCargar(); };
{ let rz; new ResizeObserver(() => { clearTimeout(rz); rz = setTimeout(() => { if (kpFilas.length && $('#perCard').offsetParent) kpPintar(); }, 100); }).observe($('#kpBox')); }

// ---------- informe PDF (se imprime / guarda como PDF desde el navegador) ----------
function informePDF() {
  if (kst.modo === 'r') return informeRecojosPDF();
  const R = kpiResultado; if (!R) return toast('Esperar a que finalice el cálculo de los indicadores', 'err');
  const w = window.open('', '_blank'); if (!w) return toast('El navegador bloqueó la ventana del informe. Habilitar las ventanas emergentes para este sitio.', 'err');
  const T = R.total, dias = diasResultado, env = kpEnvios, nF = env.filter(e => e.factura).length, conta = dias.reduce((x, d) => x + d.contado, 0), cred = dias.reduce((x, d) => x + d.credito, 0);
  const filtros = [...document.querySelectorAll('#kAct .pill')].map(x => x.textContent).join(', '), hall = kHall.map(h => `<li><b>${esc(h.t)}.</b> ${esc(h.d)}</li>`).join('');
  const tabla = (t, cab, filas) => `<h2>${t}</h2><table><tr>${cab.map((c, i) => `<th${i ? ' class="r"' : ''}>${c}</th>`).join('')}</tr>${filas.map(f => `<tr>${f.map((c, i) => `<td${i ? ' class="r"' : ''}>${c}</td>`).join('')}</tr>`).join('')}</table>`;
  const top = (mapa, n) => [...mapa].sort((a, b) => b[1].costo - a[1].costo).slice(0, n), pc = x => T ? (x / T * 100).toFixed(1) + ' %' : '-';
  const caja = v => v.cajas ? fmt(v.costo / v.cajas) : '-', ped = v => v.pedidos ? fmt(v.costo / v.pedidos) : '-';
  const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Informe de costos de Red Troncal</title><style>
    @page { size:A4; margin:14mm } * { box-sizing:border-box } body { font:11px/1.45 Inter,system-ui,Arial,sans-serif; color:#0f172a; margin:0 }
    header { display:flex; justify-content:space-between; align-items:center; border-bottom:3px solid #059669; padding-bottom:10px; margin-bottom:12px } header img { height:42px } h1 { font-size:20px; margin:0 } .sub { color:#64748b; font-size:11px }
    h2 { font-size:13px; margin:18px 0 6px; border-left:4px solid #2563eb; padding-left:8px; break-after:avoid } table { width:100%; border-collapse:collapse; margin-bottom:6px } th { background:#f1f5f9; text-align:left; padding:5px 7px; font-size:10px; text-transform:uppercase; letter-spacing:.04em; color:#475569 }
    td { padding:4px 7px; border-top:1px solid #e2e8f0 } th.r, td.r { text-align:right; font-variant-numeric:tabular-nums } tr { break-inside:avoid }
    .kp { display:grid; grid-template-columns:repeat(3,1fr); gap:8px } .kp div { border:1px solid #e2e8f0; border-radius:8px; padding:8px 10px } .kp span { display:block; font-size:9px; color:#94a3b8; font-weight:700; text-transform:uppercase; letter-spacing:.05em } .kp b { font-size:17px } .kp small { display:block; color:#64748b }
    ul { margin:4px 0 0 16px; padding:0 } li { margin:3px 0 } footer { margin-top:18px; color:#94a3b8; font-size:9px; border-top:1px solid #e2e8f0; padding-top:6px }
  </style></head><body>
    <header><div><h1>Informe de costos de transporte</h1><div class="sub">${esc(filtros)}</div></div><img src="${esc(new URL('logo.png', location.href).href)}" alt="Dinet"></header>
    <div class="kp">
      <div><span>Costo total</span><b>${fmt(T)}</b><small>${R.conImporte} registros con importe</small></div>
      <div><span>Costo por pedido</span><b>${R.pedidos ? fmt(T / R.pedidos) : '-'}</b><small>${fmtN(R.pedidos)} pedidos</small></div>
      <div><span>Costo por caja (tarifa efectiva)</span><b>${R.cajas ? fmt(T / R.cajas) : '-'}</b><small>${fmtN(R.cajas)} cajas</small></div>
      <div><span>Promedio por día</span><b>${dias.length ? fmt(T / dias.length) : '-'}</b><small>${dias.length} ${dias.length === 1 ? 'día' : 'días'} con envíos</small></div>
      <div><span>Contado y crédito</span><b>${fmt(conta)} / ${fmt(cred)}</b><small>${conta + cred ? Math.round(conta / (conta + cred) * 100) : 0} % contado</small></div>
      <div><span>Respaldo documental</span><b>${env.length ? Math.round(nF / env.length * 100) : 0} %</b><small>${nF} de ${env.length} con N° de factura</small></div>
    </div>
    ${hall ? `<h2>Hallazgos</h2><ul>${hall}</ul>` : ''}
    ${tabla('Costo por cuenta (top 15)', ['Cuenta', 'Costo', '% total', 'Pedidos', 'Costo por pedido'], top(R.cuentas, 15).map(([k, v]) => [esc(k), fmt(v.costo), pc(v.costo), fmtN(v.pedidos), ped(v)]))}
    ${tabla('Costo por nodo (top 15)', ['Nodo', 'Costo', '% total', 'Cajas', 'Tarifa por caja'], top(R.nodos, 15).map(([k, v]) => [esc(k), fmt(v.costo), pc(v.costo), fmtN(v.cajas), caja(v)]))}
    ${tabla('Costo por agencia', ['Agencia', 'Costo', '% total', 'Cajas', 'Tarifa por caja', 'Contado', 'Crédito'], top(R.agencias, 20).map(([k, v]) => [esc(k), fmt(v.costo), pc(v.costo), fmtN(v.cajas), caja(v), fmt(v.contado), fmt(v.credito)]))}
    ${tabla('Costo por zona', ['Zona', 'Costo', '% total', 'Cajas', 'Tarifa por caja'], top(R.zonas, 20).map(([k, v]) => [esc(k), fmt(v.costo), pc(v.costo), fmtN(v.cajas), caja(v)]))}
    ${tabla('Costo por transporte', ['Transporte', 'Costo', '% total', 'Pedidos', 'Costo por pedido', 'Tarifa por caja'], top(R.transportes, 20).map(([k, v]) => [esc(k), fmt(v.costo), pc(v.costo), fmtN(v.pedidos), ped(v), caja(v)]))}
    ${dias.length > 1 ? tabla('Costo por día', ['Día', 'Registros', 'Pedidos', 'Contado', 'Crédito', 'Costo total'], dias.map(d => [fdmy(d.fecha), d.registros, fmtN(d.pedidos), fmt(d.contado), fmt(d.credito), fmt(d.costo)])) : ''}
    <footer>Informe generado el ${new Date().toLocaleString('es-PE')} | Red Troncal, Dinet Logística</footer>
  </body></html>`;
  w.document.open(); w.document.write(html); w.document.close(); setTimeout(() => { w.focus(); w.print(); }, 600);
}
$('#kPdf').onclick = informePDF;

async function calcularKpis(silencioso) {
  if (kst.modo === 'r') return calcularRecojos(silencioso);
  const desde = $('#kDesde').value, hasta = $('#kHasta').value, trans = $('#kTrans').value, mi = ++kSeq;
  if (desde && hasta && desde > hasta) { $('#kgrid').innerHTML = '<div class="kg-c kg-vacio"><b>La fecha inicial es posterior a la fecha final</b><span>Corregir el rango para visualizar los costos.</span></div>'; $('#kgAl').innerHTML = ''; $('#tab-kpis').classList.remove('kload'); return; }
  $('#kRefrescar').disabled = true;
  poblarFiltros(); pintarActivos();
  // si ya había resultados se dejan a la vista (atenuados) mientras se recalcula; solo la primera vez se muestra el esqueleto
  $('#tab-kpis').classList.add('kload');
  if (silencioso !== true && !$('#kgrid .kg-hero')) $('#kgrid').innerHTML = ['s5', 's4', 's3', 's9'].map(c => `<div class="kg-sk ${c}"></div>`).join('');
  try {
    // las cuatro consultas salen a la vez para que el recálculo sea más rápido
    const pEnvios = traerTodo(() => {
      let q = sb.from('envios').select('id,fecha,nodo,motivo,transporte,agencia,zona,pedidos,bultos,cajas,importe,modo_pago,factura,factura_archivo,liquidacion_id').order('id');
      if (desde) q = q.gte('fecha', desde); if (hasta) q = q.lte('fecha', hasta); if (trans) q = q.eq('transporte', trans);
      return q;
    });
    const [envios0, hrn, prev, meta] = await Promise.all([pEnvios, traerHrn(desde, hasta), periodoPrevio(desde, hasta, trans), cargarMeta()]);
    let envios = envios0;
    kMeta = meta;
    if (mi !== kSeq) return;
    if (!trans) pintarTransSeg([...new Set(envios.map(e => e.transporte).filter(Boolean))].sort());
    poblarFiltros(envios);
    envios = envios.filter(kPasa); kpEnvios = envios;
    const R = calcularKpi(envios, hrn, 'bultos');
    kpiResultado = R;
    const dias = agruparDias(envios); diasResultado = dias;
    const nombradas = [...R.cuentas].filter(([k]) => !k.startsWith('('));
    kHall = hallazgosKpi(R, envios); pintarResumen(R, dias, prev, envios);
    $('#nCta').textContent = nombradas.length; $('#nNodo').textContent = R.nodos.size; $('#nAge').textContent = R.agencias.size;
    pintarDias(dias); kpCache.clear(); kpCargar();
    kaDatos(R, envios, prev); calidadKpi(R, envios);
    $('#kpMsg').className = 'msg'; kUpdTs = Date.now(); pintarUpd(); kHash();
  } catch (err) {
    flash($('#kpMsg'), 'No se pudo calcular: ' + err.message, 'err');
    if (mi === kSeq && !$('#kgrid .kg-hero')) $('#kgrid').innerHTML = '<div class="kg-c kg-vacio"><b>No fue posible calcular los costos</b><span>Verificar la conexión y presionar Actualizar.</span></div>';
  }
  if (mi === kSeq) { $('#kRefrescar').disabled = false; $('#tab-kpis').classList.remove('kload'); }
}
// ---------- dashboard de recojos: costo del recojo repartido entre los pedidos de su detalle (fecha + nodo) ----------
let kRecTab = 'nodo', kRecR = null, kRecPrev = null, kRecCpp = 0, kRecDet = true, krPlazo = 5;
try { const v = Number(localStorage.getItem('recPlazo')); if (v >= 1 && v <= 60) krPlazo = v; } catch (e) {}
const REC_SEL = 'fecha_reporte,nombre_cuenta,cuenta,nro_pedido,nodo,bultos,motivo_devolucion,fecha_pedido,fecha_solicitud_cx,fecha_solicitud_nodo,fecha_salida_nodo,fecha_recojo_agencia,fecha_llegada_ctd,guia,proveedor';
const recDias = (a, b) => a && b ? Math.round((new Date(b) - new Date(a)) / 864e5) : null;
const recClave = (f, n) => String(f).slice(0, 10) + '|' + norm(n);
function recCalcular(envios, det, cuentaSel, nodoSel) {
  const eMap = new Map(), dMap = new Map(), cta = r => r.nombre_cuenta || r.cuenta || '(sin cuenta)';
  envios.forEach(e => {
    const k = recClave(e.fecha, e.nodo), g = eMap.get(k) || { k, nodo: norm(e.nodo), fecha: String(e.fecha).slice(0, 10), importe: 0, pedidos: 0, bultos: 0, sinImporte: 0, trans: e.transporte || '' }, i = Number(e.importe) || 0;
    if (i > 0) g.importe += i; else g.sinImporte++;
    g.pedidos += Number(e.pedidos) || 0; g.bultos += Number(e.bultos) || 0; eMap.set(k, g);
  });
  det.forEach(r => { const k = recClave(r.fecha_reporte, r.nodo); if (!dMap.has(k)) dMap.set(k, []); dMap.get(k).push(r); });
  const todos = [];
  dMap.forEach((rows, k) => {
    const g = eMap.get(k), tot = rows.reduce((x, r) => x + (Number(r.bultos) || 1), 0);
    rows.forEach(r => todos.push({ k, nodo: norm(r.nodo), cuenta: cta(r), motivo: r.motivo_devolucion || '(sin motivo)', costo: g && g.importe > 0 ? g.importe * (Number(r.bultos) || 1) / tot : 0, n: 1, bultos: Number(r.bultos) || 0, dias: recDias(r.fecha_solicitud_cx, r.fecha_llegada_ctd), guia: r.guia, fecha: String(r.fecha_reporte).slice(0, 10), pedido: r.nro_pedido, prov: r.proveedor || '(sin proveedor)', ped: r.fecha_pedido, sol: r.fecha_solicitud_cx, sn: r.fecha_solicitud_nodo, sa: r.fecha_salida_nodo, ra: r.fecha_recojo_agencia, lleg: r.fecha_llegada_ctd }));
  });
  const sinDet = [...eMap.values()].filter(g => g.importe > 0 && !dMap.has(g.k));
  sinDet.forEach(g => todos.push({ k: g.k, nodo: g.nodo, cuenta: '(sin detalle)', motivo: '(sin detalle)', costo: g.importe, n: g.pedidos, bultos: g.bultos, dias: null, guia: null, fecha: g.fecha, prov: g.trans || '(sin proveedor)', synth: true }));
  const it = todos.filter(x => (!cuentaSel || x.cuenta === cuentaSel) && (!nodoSel || x.nodo === nodoSel));
  const grupo = f => { const m = new Map(); it.forEach(x => { const k = f(x), a = m.get(k) || { costo: 0, n: 0, dsum: 0, dn: 0 }; a.costo += x.costo; a.n += x.n; if (x.dias != null) { a.dsum += x.dias; a.dn++; } m.set(k, a); }); return m; };
  const conDias = it.filter(x => x.dias != null), g = { nodo: grupo(x => x.nodo), cuenta: grupo(x => x.cuenta), motivo: grupo(x => x.motivo), prov: grupo(x => x.prov || '(sin proveedor)') };
  const etapa = (p, q) => { const v = it.map(x => recDias(x[p], x[q])).filter(x => x != null); return v.length ? v.reduce((s, y) => s + y, 0) / v.length : null; };
  const total = [...eMap.values()].filter(x => x.importe > 0);
  return {
    etapas: [['Pedido a solicitud CX', etapa('ped', 'sol')], ['Solicitud CX a solicitud al nodo', etapa('sol', 'sn')], ['Solicitud al nodo a salida del nodo', etapa('sn', 'sa')], ['Salida del nodo a recojo en agencia', etapa('sa', 'ra')], ['Recojo en agencia a llegada a CTD', etapa('ra', 'lleg')]], totalCompleto: etapa('ped', 'lleg'),
    items: it, costo: it.reduce((x, y) => x + y.costo, 0), pedidos: it.reduce((x, y) => x + y.n, 0), bultos: it.reduce((x, y) => x + y.bultos, 0), guias: new Set(it.filter(x => x.guia).map(x => x.guia)).size,
    registros: new Set(it.filter(x => x.costo > 0).map(x => x.k)).size, nodos: new Set(it.filter(x => x.costo > 0).map(x => x.nodo)).size,
    dias: conDias.length ? conDias.reduce((x, y) => x + y.dias, 0) / conDias.length : null, grupos: g,
    totalEnv: total.length, conDetalle: total.filter(x => dMap.has(x.k)).length, sinDetalle: sinDet, sinRecojo: [...dMap].filter(([k]) => !eMap.has(k)).map(([k, rows]) => ({ k, nodo: norm(rows[0].nodo), fecha: String(rows[0].fecha_reporte).slice(0, 10), n: rows.length })),
    sinImporte: [...eMap.values()].reduce((x, y) => x + y.sinImporte, 0)
  };
}
function recOpciones(id, lista, vacio) {
  const sel = $(id), v = sel.value; sel.innerHTML = `<option value="">${vacio}</option>` + [...new Set([...lista, v].filter(Boolean))].sort().map(x => `<option>${esc(x)}</option>`).join(''); sel.value = v;
}
async function recPrevio(desde, hasta, trans) {
  if (!desde || !hasta) return null;
  const d0 = new Date(desde + 'T00:00:00'), d1 = new Date(hasta + 'T00:00:00'), n = Math.round((d1 - d0) / 86400000) + 1;
  const pf = isoLocal(new Date(d0.getFullYear(), d0.getMonth(), d0.getDate() - n)), pt = isoLocal(new Date(d0.getFullYear(), d0.getMonth(), d0.getDate() - 1)), rg = (q, col) => q.gte(col, pf).lte(col, pt);
  try {
    const [envios, det] = await Promise.all([
      traerTodo(() => { let q = rg(sb.from('envios').select('id,fecha,nodo,motivo,transporte,pedidos,bultos,importe').eq('motivo', 'RECOJO'), 'fecha').order('id'); if (trans) q = q.eq('transporte', trans); return q; }),
      traerTodo(() => rg(sb.from('recojo_detalle').select(REC_SEL).order('id'), 'fecha_reporte'))]);
    const ok = trans ? new Set(envios.map(e => recClave(e.fecha, e.nodo))) : null;
    return { envios, det: ok ? det.filter(r => ok.has(recClave(r.fecha_reporte, r.nodo))) : det };
  } catch (e) { return null; }
}
async function calcularRecojos(silencioso) {
  const desde = $('#kDesde').value, hasta = $('#kHasta').value, trans = $('#kTrans').value, mi = ++kSeq;
  if (desde && hasta && desde > hasta) { $('#kRec').innerHTML = '<div class="kr-c kr-vacio"><b>La fecha inicial es posterior a la fecha final</b>Corregir el rango para visualizar la información.</div>'; $('#tab-kpis').classList.remove('kload'); return; }
  $('#kRefrescar').disabled = true; pintarActivos(); $('#tab-kpis').classList.add('kload');
  try {
    const rango = (q, col) => { if (desde) q = q.gte(col, desde); if (hasta) q = q.lte(col, hasta); return q; };
    const [envios, det, desp, pv] = await Promise.all([
      traerTodo(() => { let q = rango(sb.from('envios').select('id,fecha,nodo,motivo,transporte,pedidos,bultos,importe').eq('motivo', 'RECOJO'), 'fecha').order('id'); if (trans) q = q.eq('transporte', trans); return q; }),
      traerTodo(() => rango(sb.from('recojo_detalle').select(REC_SEL).order('id'), 'fecha_reporte')).catch(e => { if (/recojo_detalle/.test(e.message)) return null; throw e; }),
      traerTodo(() => { let q = rango(sb.from('envios').select('importe,pedidos,motivo'), 'fecha').order('id'); if (trans) q = q.eq('transporte', trans); return q; }).catch(() => []),
      recPrevio(desde, hasta, trans)]);
    if (mi !== kSeq) return;
    if (!trans) pintarTransSeg([...new Set(envios.map(e => e.transporte).filter(Boolean))].sort());
    const d = det || [], enlazadas = trans ? new Set(envios.map(e => recClave(e.fecha, e.nodo))) : null;
    const detF = enlazadas ? d.filter(r => enlazadas.has(recClave(r.fecha_reporte, r.nodo))) : d;
    recOpciones('#kCuenta', detF.map(r => r.nombre_cuenta || r.cuenta || '(sin cuenta)'), 'Todas'); recOpciones('#kNodoR', [...envios.map(e => norm(e.nodo)), ...detF.map(r => norm(r.nodo))], 'Todos');
    const dp = desp.filter(e => e.motivo !== 'RECOJO' && Number(e.importe) > 0), dped = dp.reduce((x, e) => x + (Number(e.pedidos) || 0), 0);
    kRecCpp = dped ? dp.reduce((x, e) => x + Number(e.importe), 0) / dped : 0; kRecDet = !!det;
    kRecR = recCalcular(envios, detF, $('#kCuenta').value, $('#kNodoR').value); krM.vb = null; kRecPrev = pv ? recCalcular(pv.envios, pv.det, $('#kCuenta').value, $('#kNodoR').value) : null;
    pintarRecojos();
    $('#kpMsg').className = 'msg'; kUpdTs = Date.now(); pintarUpd(); kHash();
  } catch (err) { flash($('#kpMsg'), 'No se pudo calcular: ' + err.message, 'err'); }
  if (mi === kSeq) { $('#kRefrescar').disabled = false; $('#tab-kpis').classList.remove('kload'); }
}
const recCumple = (R, plazo) => { const v = R.items.filter(x => x.dias != null), dentro = v.filter(x => x.dias <= plazo).length; return { n: v.length, dentro, pct: v.length ? dentro / v.length * 100 : 0, fuera: v.filter(x => x.dias > plazo).sort((a, b) => b.dias - a.dias) }; };
function recHallazgos(R, P) {
  const h = [], lentos = krLentos(R), C = recCumple(R, krPlazo);
  if (lentos.length > 1 && R.dias) { const [n, v] = lentos[0], dif = (v / R.dias - 1) * 100; if (dif >= 25) h.push(['w', `${n} presenta el mayor tiempo de retorno: ${v.toFixed(1)} días, ${dif.toFixed(0)} % sobre el promedio (${R.dias.toFixed(1)} días).`]); }
  const pr = [...R.grupos.prov].filter(([k, v]) => v.dn && !k.startsWith('(')).map(([k, v]) => [k, v.dsum / v.dn]).sort((a, b) => b[1] - a[1]);
  if (pr.length > 1 && pr[0][1] >= pr[pr.length - 1][1] * 1.5) h.push(['w', `${pr[0][0]} registra un retorno promedio de ${pr[0][1].toFixed(1)} días, frente a ${pr[pr.length - 1][1].toFixed(1)} días de ${pr[pr.length - 1][0]}.`]);
  const et = R.etapas.filter(e => e[1] != null && e[1] > 0), tot = et.reduce((x, e) => x + e[1], 0);
  if (et.length && tot > 0) { const m = et.reduce((a, b) => b[1] > a[1] ? b : a); h.push(['i', `La etapa de mayor duración es "${m[0]}": ${m[1].toFixed(1)} días, equivalente al ${(m[1] / tot * 100).toFixed(0)} % del recorrido.`]); }
  if (C.n) h.push([C.pct >= 90 ? 'ok' : C.pct >= 70 ? 'i' : 'w', `${C.dentro} de ${C.n} pedidos (${C.pct.toFixed(0)} %) retornaron dentro del plazo objetivo de ${krPlazo} días.`]);
  if (kRecCpp && R.costo && R.pedidos) { const cpp = R.costo / R.pedidos; h.push(['i', `El costo por pedido de recojo equivale a ${(cpp / kRecCpp).toFixed(1)} veces el de despacho (${fmt(cpp)} frente a ${fmt(kRecCpp)}).`]); }
  if (P && P.costo && R.costo) { const d = (R.costo / P.costo - 1) * 100; if (Math.abs(d) >= 10) h.push([d > 0 ? 'w' : 'ok', `El costo de recojos ${d > 0 ? 'aumentó' : 'disminuyó'} ${Math.abs(d).toFixed(0)} % respecto del período anterior (${fmt(P.costo)}).`]); }
  if (P && P.dias && R.dias) { const d = (R.dias / P.dias - 1) * 100; if (Math.abs(d) >= 15) h.push([d > 0 ? 'w' : 'ok', `El tiempo promedio de retorno ${d > 0 ? 'aumentó' : 'disminuyó'} ${Math.abs(d).toFixed(0)} % respecto del período anterior (${P.dias.toFixed(1)} días).`]); }
  if (R.totalEnv && R.conDetalle < R.totalEnv) h.push(['w', `${R.totalEnv - R.conDetalle} ${R.totalEnv - R.conDetalle === 1 ? 'recojo con importe no tiene' : 'recojos con importe no tienen'} detalle asociado: su costo no se distribuye por cuenta, motivo ni proveedor.`]);
  return h;
}
function krHallazgos(R) {
  const h = recHallazgos(R, kRecPrev).slice(0, 6); if (!h.length) return '<span id="krHall"></span>';
  return `<section class="kr-c kr-hall" id="krHall"><div class="kr-cab"><div><h3>Hallazgos del período</h3><p>Principales observaciones, generadas a partir de los datos del período.</p></div></div><ul>${h.map(([c, t]) => `<li><i class="kr-hd ${c}"></i><span>${esc(t)}</span></li>`).join('')}</ul></section>`;
}
const KR_ET = ['#1a5db8', '#4f93e6', '#34a853', '#e8a317', '#8e8e93'];
function krEtapas(R) {
  const et = R.etapas, hay = et.some(e => e[1] != null), suma = et.reduce((x, e) => x + (e[1] > 0 ? e[1] : 0), 0), mx = et.reduce((m, e) => e[1] != null && e[1] > m ? e[1] : m, 0);
  if (!hay) return `<section class="kr-c"><div class="kr-cab"><div><h3>Tiempo por etapa del recorrido</h3><p>Promedio de días de cada etapa, desde el pedido hasta la llegada a CTD.</p></div></div><p class="kr-h">El detalle del período no incluye las fechas de cada etapa.</p></section>`;
  return `<section class="kr-c"><div class="kr-cab"><div><h3>Tiempo por etapa del recorrido</h3><p>Promedio de días de cada etapa, desde el pedido hasta la llegada a CTD.</p></div>${R.totalCompleto != null ? `<div class="kr-tot">${R.totalCompleto.toFixed(1)} días<small>recorrido total</small></div>` : ''}</div>
    <div class="kr-et">${et.map((e, i) => e[1] > 0 ? `<i style="flex:${(e[1] / suma).toFixed(4)};background:${KR_ET[i]}" title="${esc(e[0])}: ${e[1].toFixed(1)} días"></i>` : '').join('')}</div>
    <ul class="kr-etl">${et.map((e, i) => `<li><i style="background:${KR_ET[i]}"></i><span>${esc(e[0])}${e[1] != null && e[1] === mx && mx > 0 ? '<em>Mayor duración</em>' : ''}</span><b>${e[1] != null ? e[1].toFixed(1) + ' días' : '-'}</b></li>`).join('')}</ul></section>`;
}
function krCump(R) {
  const C = recCumple(R, krPlazo), cls = C.pct >= 90 ? 'v' : C.pct >= 70 ? 'a' : 'r';
  const cab = `<div class="kr-cab"><div><h3>Cumplimiento de plazo</h3><p>Porcentaje de pedidos retornados dentro del plazo objetivo, medido desde la solicitud CX.</p></div><label class="kr-pl">Plazo objetivo <input type="number" id="krPlazo" min="1" max="60" value="${krPlazo}" aria-label="Plazo objetivo en días"> días</label></div>`;
  if (!C.n) return `<section class="kr-c" id="krCump">${cab}<p class="kr-h">El detalle no incluye las fechas necesarias para medir el plazo.</p></section>`;
  return `<section class="kr-c" id="krCump">${cab}<div class="kr-pct"><b>${C.pct.toFixed(0)} %</b><span>${C.dentro} de ${C.n} pedidos dentro de ${krPlazo} días</span></div><div class="kr-pb"><i class="${cls}" style="width:${C.pct.toFixed(1)}%"></i></div>
    ${C.fuera.length ? `<div class="kr-fu">Pedidos fuera de plazo${C.fuera.length > 5 ? ` (los 5 de mayor duración, de ${C.fuera.length})` : ` (${C.fuera.length})`}</div><div class="kr-tw"><table class="kr-t"><thead><tr><th>Pedido</th><th>Nodo</th><th>Proveedor</th><th class="n">Días</th></tr></thead><tbody>${C.fuera.slice(0, 5).map(x => `<tr><td class="cut" title="${esc(x.pedido || '')}">${esc(x.pedido || '-')}</td><td>${esc(x.nodo)}</td><td class="cut">${esc(x.prov)}</td><td class="n dlento">${x.dias}</td></tr>`).join('')}</tbody></table></div>` : '<p class="kr-h">Todos los pedidos retornaron dentro del plazo objetivo.</p>'}</section>`;
}
const krLentos = R => [...R.grupos.nodo].filter(([, v]) => v.dn).map(([k, v]) => [k, v.dsum / v.dn]).sort((a, b) => b[1] - a[1]);
const krBarras = () => requestAnimationFrame(() => requestAnimationFrame(() => document.querySelectorAll('#kRec .kr-bar i').forEach(i => i.style.width = i.dataset.w + '%')));
function krTabla(R) {
  const lentos = krLentos(R), tipo = kRecTab, conDias = tipo === 'nodo' || tipo === 'prov', G = [...R.grupos[tipo]].sort((a, b) => b[1].costo - a[1].costo), mx = Math.max(...G.map(([, v]) => conDias ? (v.dn ? v.dsum / v.dn : 0) : v.costo), 1);
  const nombre = { nodo: 'Nodo', cuenta: 'Cuenta', motivo: 'Motivo', prov: 'Proveedor' }[tipo];
  const filas = G.map(([k, v], i) => {
    const dp = v.dn ? v.dsum / v.dn : null, w = conDias ? (dp || 0) / mx * 100 : v.costo / mx * 100, lento = conDias && dp != null && lentos.length > 1 && dp >= lentos[0][1] && dp > R.dias * 1.25;
    return `<tr style="animation-delay:${Math.min(i, 10) * .05}s"><td class="nom"><b>${esc(k)}</b></td><td class="n">${fmtN(v.n)}</td><td class="n">${v.costo ? fmt(v.costo) : '-'}</td><td class="n">${v.costo && v.n ? fmt(v.costo / v.n) : '-'}</td>
      <td><div class="kr-bar${lento ? ' l' : ''}"><span><i data-w="${w.toFixed(1)}"></i></span>${conDias ? `<b>${dp != null ? dp.toFixed(1) : '-'}</b>` : ''}</div></td></tr>`; }).join('');
  const ok = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>', ale = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4m0 4h.01M10.3 3.9L1.8 18a2 2 0 001.7 3h17a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/></svg>';
  const lista = (a, f) => a.slice(0, 3).map(f).join(', ') + (a.length > 3 ? '…' : '');
  const avisos = [];
  if (!kRecDet) avisos.push(['w', 'Detalle de recojos no disponible: falta crear la tabla en la base de datos (supabase_recojo_detalle.sql)']);
  if (R.totalEnv) avisos.push(R.conDetalle === R.totalEnv ? ['', `${R.conDetalle} de ${R.totalEnv} ${R.totalEnv === 1 ? 'recojo cuenta' : 'recojos cuentan'} con detalle asociado`] : ['w', `${R.totalEnv - R.conDetalle} ${R.totalEnv - R.conDetalle === 1 ? 'recojo con importe sin detalle asociado' : 'recojos con importe sin detalle asociado'}: ${lista(R.sinDetalle, g => `${g.nodo} ${fdmy(g.fecha).slice(0, 5)}`)}`]);
  if (R.sinRecojo.length) avisos.push(['w', `Detalle sin recojo registrado en el período: ${lista(R.sinRecojo, g => `${g.nodo} ${fdmy(g.fecha).slice(0, 5)} (${g.n})`)}`]);
  if (R.sinImporte) avisos.push(['w', `${R.sinImporte} ${R.sinImporte === 1 ? 'recojo sin importe registrado' : 'recojos sin importe registrado'}`]);
  return `<section class="kr-c" id="krTabla"><div class="kr-tabs"><b>${{ nodo: 'Costo y tiempo de retorno por nodo', cuenta: 'Costo distribuido por cuenta', motivo: 'Costo distribuido por motivo de devolución', prov: 'Costo y tiempo de retorno por proveedor' }[tipo]}</b>
      <div class="pseg" id="krSeg">${[['nodo', 'Por nodo'], ['prov', 'Por proveedor'], ['cuenta', 'Por cuenta'], ['motivo', 'Por motivo']].map(([k, t]) => `<button type="button" data-rt="${k}" class="${tipo === k ? 'on' : ''}">${t}</button>`).join('')}</div></div>
      <div class="kr-tw"><table class="kr-t"><thead><tr><th>${nombre}</th><th class="n">Pedidos</th><th class="n">Costo</th><th class="n">Por pedido</th><th style="width:28%">${conDias ? 'Días de retorno' : 'Participación'}</th></tr></thead><tbody>${filas}</tbody></table></div>
      ${avisos.length ? `<div class="kr-av">${avisos.map(([c, t]) => `<span class="kr-a ${c}">${c ? ale : ok}${esc(t)}</span>`).join('')}</div>` : ''}</section>`;
}
function pintarRecojos() {
  const R = kRecR, el = $('#kRec'); if (!R) return;
  const hay = R.pedidos || R.costo || R.sinRecojo.length || R.sinImporte;
  if (!hay) { el.innerHTML = '<div class="kr-c kr-vacio"><b>No hay recojos registrados en el período seleccionado</b>Verificar el rango de fechas o cargar la información en la pestaña Cargar despacho.</div>'; return; }
  const cpp = R.pedidos && R.costo ? R.costo / R.pedidos : 0, lentos = krLentos(R);
  const tag = lentos.length > 1 ? `<span class="kr-tag">Mayor: ${esc(lentos[0][0])} ${lentos[0][1].toFixed(1)}</span>` : '';
  const P = kRecPrev, pc = P && P.pedidos && P.costo ? P.costo / P.pedidos : 0;
  const dl = (x, y, tipo) => { if (!x || !y) return ''; const d = (x - y) / y * 100, c = Math.abs(d) < .5 ? '' : tipo === 'neutro' ? '' : d > 0 ? 'up' : 'dn'; return `<em class="kr-d ${c}" title="Período anterior">${Math.abs(d) < .5 ? 'Sin cambio' : (d > 0 ? '+' : '-') + Math.abs(d).toFixed(1) + ' %'}</em>`; };
  const cifra = (t, v, s, d = '') => `<div class="kr-cifra"><span>${t}</span><b>${v}${d}</b><small>${s}</small></div>`;
  const cifras = cifra('Costo de recojos', R.costo ? fmt(R.costo) : '-', R.costo ? `${R.registros} ${R.registros === 1 ? 'recojo' : 'recojos'} en ${R.nodos} ${R.nodos === 1 ? 'nodo' : 'nodos'}${P && P.costo ? `. Período anterior: ${fmt(P.costo)}` : ''}` : 'Importe pendiente de registro', P ? dl(R.costo, P.costo) : '') +
    cifra('Pedidos recogidos', fmtN(R.pedidos), `${fmtN(R.bultos)} ${R.bultos === 1 ? 'bulto' : 'bultos'}${R.guias ? `, ${R.guias} ${R.guias === 1 ? 'guía' : 'guías'}` : ''}${P && P.pedidos ? `. Período anterior: ${fmtN(P.pedidos)}` : ''}`, P ? dl(R.pedidos, P.pedidos, 'neutro') : '') +
    cifra('Costo por pedido', cpp ? fmt(cpp) : '-', kRecCpp ? (cpp ? `frente a ${fmt(kRecCpp)} en despacho` : `en despacho: ${fmt(kRecCpp)}`) : 'Sin despachos en el período para comparar', dl(cpp, pc)) +
    `<div class="kr-cifra"><span>Días de retorno${tag}</span><b>${R.dias != null ? R.dias.toFixed(1) : '-'}${P ? dl(R.dias, P.dias) : ''}</b><small>${R.dias != null ? 'de la solicitud CX a la llegada a CTD' : 'Sin fechas suficientes en el detalle'}${P && P.dias ? `. Período anterior: ${P.dias.toFixed(1)}` : ''}</small></div>`;
  el.innerHTML = `<section class="kr-c"><div class="kr-cab"><div><h3>Resumen de recojos</h3><p>Devoluciones recogidas en el período. El costo de cada recojo se distribuye entre los pedidos de su detalle, según fecha y nodo.</p></div><button type="button" class="kr-ver" id="krVer">Ver detalle de recojos</button></div><div class="kr-cifras">${cifras}</div></section>${krHallazgos(R)}<div class="kr-dos">${krEtapas(R)}${krCump(R)}</div>${krMapaHTML()}${krTabla(R)}`;
  krBarras(); krMapaDibujar();
}

// ---------- mapa de días de retorno por provincia ----------
const krM = { met: 'dias', vb: null, sel: null, ar: null, mov: false };
function krItems() {
  const out = []; kRecR.grupos.nodo.forEach((v, k) => out.push({ name: k, costo: v.costo, n: v.n, dias: v.dn ? v.dsum / v.dn : null, p: kaProv.get(norm(KA_ALIAS[k] || k)) || null })); return out;
}
const krVal = x => krM.met === 'dias' ? x.dias : krM.met === 'pedidos' ? (x.n || null) : (x.costo || null);
const krTxt = x => krM.met === 'dias' ? (x.dias != null ? x.dias.toFixed(1) : '-') : krM.met === 'pedidos' ? fmtN(x.n) : x.costo ? Math.round(x.costo).toLocaleString('es-PE') : '-';
function krColor(v, lo, hi) {
  if (v == null) return '#e6eaf0';
  const t = Math.min(1, Math.max(0, (v - lo) / (hi - lo || 1)));
  if (krM.met === 'dias') { const c = t < .5 ? [[124, 197, 147], [242, 195, 92], t * 2] : [[242, 195, 92], [226, 97, 110], (t - .5) * 2]; return 'rgb(' + c[0].map((x, i) => Math.round(x + (c[1][i] - x) * c[2])).join(',') + ')'; }
  const a = [214, 228, 250], b = [11, 61, 130], s = Math.sqrt(t); return 'rgb(' + a.map((x, i) => Math.round(x + (b[i] - x) * s)).join(',') + ')';
}
function krMapaHTML() {
  const cab = '<div class="kr-cab"><div><h3>Días de retorno por provincia</h3><p>Tiempo promedio de retorno por provincia. Las provincias con mayor tiempo se muestran en rojo.</p></div>';
  if (kaGeoEstado === 'error') return `<section class="kr-c">${cab}</div><p class="kr-h">No se pudo cargar el mapa. Actualizar la página para intentarlo de nuevo.</p></section>`;
  if (!kaGeo) { kaCargarGeo(); return `<section class="kr-c">${cab}</div><div class="ka-lienzo kr-lz" style="cursor:default"><div class="cargando">Cargando el mapa…</div></div></section>`; }
  krItems().forEach(x => { if (!x.p) x.p = null; });
  if (!krItems().some(x => x.p)) return '';
  return `<section class="kr-c" id="krMapa">${cab}<div class="pseg" id="krMet">${[['dias', 'Días de retorno'], ['pedidos', 'Pedidos'], ['costo', 'Costo']].map(([k, t]) => `<button type="button" data-km="${k}" class="${krM.met === k ? 'on' : ''}">${t}</button>`).join('')}</div></div>
    <div class="ka-mapa" style="margin-top:14px"><div><div class="ka-lienzo kr-lz" id="krLz"><svg id="krSvg" role="img" aria-label="Mapa del Perú por provincias"></svg>
      <div class="ka-zoom"><button type="button" id="krZi" aria-label="Acercar">+</button><button type="button" id="krZo" aria-label="Alejar">&minus;</button><button type="button" id="krZr" aria-label="Ver los nodos" style="font-size:13px">&#8634;</button></div><div class="ka-leyenda" id="krLey"></div></div>
      <p class="ka-fuente" id="krFuente"></p></div><div id="krPanel"></div></div></section>`;
}
function krFit(ub) {
  const W = kaGeo.w, H = kaGeo.h; if (!ub.length) return { x: -10, y: -10, w: W, h: H };
  const xs = ub.map(x => x.p[2]), ys = ub.map(x => x.p[3]), x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
  const w = Math.min(W, Math.max(W / 9, (x1 - x0) * 1.6, (y1 - y0) * 1.6 * W / H)), h = w * H / W;
  return { x: (x0 + x1) / 2 - w / 2, y: (y0 + y1) / 2 - h / 2, w, h };
}
function krMapaDibujar() {
  const lz = $('#krLz'), svg = $('#krSvg'); if (!lz || !svg || !kaGeo) return;
  const items = krItems(), ub = items.filter(x => x.p), por = new Map(ub.map(x => [x.p[0], x]));
  if (!krM.vb) krM.vb = krFit(ub);
  if (!krM.sel || !items.some(x => x.name === krM.sel)) krM.sel = (krLentos(kRecR)[0] || [ub[0]?.name])[0];
  svg.innerHTML = `<g id="krGp">${kaGeo.p.map(p => { const x = por.get(p[0]);
    const tip = x ? `<b>${esc(x.name)}</b><small>${esc(p[0])} · ${esc(p[1])}</small>${kaTf('Días de retorno', x.dias != null ? x.dias.toFixed(1) : '-')}${kaTf('Pedidos', fmtN(x.n))}${kaTf('Costo', x.costo ? fmt(x.costo) : '-')}` : `<b>${esc(p[0])}</b><small>Sin recojos en el período</small>`;
    return `<path class="ka-prov ${x ? 'd' : ''}" data-p="${esc(p[0])}" ${x ? `data-n="${esc(x.name)}"` : ''} ${kaTipA(tip)} d="${p[4]}"/>`; }).join('')}</g><g id="krGm"></g>`;
  const sin = items.filter(x => !x.p); $('#krFuente').textContent = `Límites provinciales: INEI (2007).${sin.length ? ` ${sin.length} ${sin.length === 1 ? 'nodo sin ubicación' : 'nodos sin ubicación'} en el mapa (${sin.map(x => x.name).join(', ')}).` : ''}`;
  krMapaActualizar(true);
  const zl = $('#krLz');
  zl.addEventListener('pointerdown', e => { if (e.target.closest('.ka-zoom')) return; krM.ar = { x: e.clientX, y: e.clientY, vx: krM.vb.x, vy: krM.vb.y, t: e.target.closest('.ka-prov.d, .kr-mk') }; krM.mov = false; });
  zl.addEventListener('wheel', e => { e.preventDefault(); krZoom(e.deltaY < 0 ? .8 : 1.25, e.clientX, e.clientY); }, { passive: false });
  $('#krZi').onclick = () => krZoom(.7); $('#krZo').onclick = () => krZoom(1.4); $('#krZr').onclick = () => { krM.vb = krFit(krItems().filter(x => x.p)); krAjustar(); };
}
function krMapaActualizar(todo) {
  const svg = $('#krSvg'); if (!svg || !kaGeo) return;
  const items = krItems(), ub = items.filter(x => x.p), por = new Map(ub.map(x => [x.p[0], x])), vals = ub.map(krVal).filter(v => v != null), lo = vals.length ? Math.min(...vals) : 0, hi = vals.length ? Math.max(...vals) : 1;
  svg.querySelectorAll('.ka-prov').forEach(el => { const x = por.get(el.dataset.p); el.setAttribute('fill', x ? krColor(krVal(x), lo, hi) : '#e6eaf0'); el.classList.toggle('sel', !!x && x.name === krM.sel); if (x && x.name === krM.sel) el.parentNode.appendChild(el); });
  const lz = $('#krLz'), K = Math.max(krM.vb.w / (lz.clientWidth || 600), krM.vb.h / (lz.clientHeight || 400));   // unidades del mapa por píxel: los marcadores mantienen su tamaño en pantalla
  $('#krGm').innerHTML = ub.filter(x => krVal(x) != null || x.name === krM.sel).map(x => `<g class="kr-mk${x.name === krM.sel ? ' sel' : ''}" data-n="${esc(x.name)}" ${kaTipA(`<b>${esc(x.name)}</b><small>${esc(x.p[0])} · ${esc(x.p[1])}</small>${kaTf('Días de retorno', x.dias != null ? x.dias.toFixed(1) : '-')}${kaTf('Pedidos', fmtN(x.n))}${kaTf('Costo', x.costo ? fmt(x.costo) : '-')}`)} transform="translate(${x.p[2]},${x.p[3]}) scale(${K.toFixed(3)})"><circle r="${krTxt(x).length > 4 ? 17 : 13}" style="stroke:${krColor(krVal(x), lo, hi)}"/><text>${krTxt(x)}</text></g>`).join('');
  if (todo) {
    svg.setAttribute('viewBox', `${krM.vb.x} ${krM.vb.y} ${krM.vb.w} ${krM.vb.h}`);
    $('#krLey').innerHTML = krM.met === 'dias' ? `<b>Días de retorno promedio</b><div class="ka-grad" style="background:linear-gradient(90deg,#7cc593,#f2c35c,#e2616e)"></div><div class="ka-gl"><span>${lo.toFixed(1)} días</span><span>${hi.toFixed(1)} días</span></div>`
      : `<b>${krM.met === 'pedidos' ? 'Pedidos recogidos' : 'Costo del recojo'}</b><div class="ka-grad"></div><div class="ka-gl"><span>${krM.met === 'pedidos' ? fmtN(lo) : kaF0(lo)}</span><span>${krM.met === 'pedidos' ? fmtN(hi) : kaF0(hi)}</span></div>`;
  }
  krPanel(items, ub);
}
function krPanel(items, ub) {
  const R = kRecR, ord = ub.filter(x => krVal(x) != null).sort((a, b) => krVal(b) - krVal(a)), mx = ord.length ? krVal(ord[0]) : 1, x = items.find(i => i.name === krM.sel);
  let ficha = '';
  if (x) {
    const dif = x.dias != null && R.dias ? (x.dias / R.dias - 1) * 100 : null;
    ficha = `<div class="kr-fi"><small>${x.p ? esc(x.p[0] + ' · ' + x.p[1]) : 'Sin ubicación en el mapa'}</small><h4>${esc(x.name)}</h4>
      <div class="kr-fg"><div><span>Días de retorno</span><b>${x.dias != null ? x.dias.toFixed(1) : '-'}</b></div><div><span>Pedidos</span><b>${fmtN(x.n)}</b></div><div><span>Costo</span><b>${x.costo ? fmt(x.costo) : '-'}</b></div></div>
      ${dif != null ? `<p class="${dif > 5 ? 'lento' : dif < -5 ? 'rapido' : ''}">${Math.abs(dif) <= 5 ? 'En línea con el promedio' : `${Math.abs(dif).toFixed(0)} % ${dif > 0 ? 'sobre' : 'bajo'} el promedio`} (${R.dias.toFixed(1)} días)</p>` : ''}</div>`;
  }
  const tit = { dias: 'Provincias con mayor tiempo de retorno', pedidos: 'Provincias con más pedidos', costo: 'Provincias de mayor costo' }[krM.met];
  $('#krPanel').innerHTML = ficha + `<div class="ka-sec" style="margin-top:14px">${tit}</div><div class="ka-top">${ord.slice(0, 8).map((n, i) => `<button type="button" data-n="${esc(n.name)}" class="${n.name === krM.sel ? 'sel' : ''}"><span class="n">${i + 1}</span><span class="nm"><b>${esc(n.name)}</b><i style="width:${krVal(n) / mx * 100}%"></i></span><span class="v">${krM.met === 'dias' ? n.dias.toFixed(1) + ' días' : krM.met === 'pedidos' ? fmtN(n.n) + ' pedidos' : fmt(n.costo)}</span></button>`).join('') || '<p class="kr-h">Sin datos para este indicador en el período.</p>'}</div>`;
}
function krAjustar() {
  const v = krM.vb, w0 = kaGeo.w, h0 = kaGeo.h;
  v.x = Math.min(-10 + w0 - v.w * .25, Math.max(-10 - v.w * .75, v.x)); v.y = Math.min(-10 + h0 - v.h * .25, Math.max(-10 - v.h * .75, v.y));
  $('#krSvg').setAttribute('viewBox', `${v.x} ${v.y} ${v.w} ${v.h}`); krMapaActualizar(false);
}
function krZoom(f, cx, cy) {
  const r = $('#krSvg').getBoundingClientRect(), v = krM.vb, s = Math.max(v.w / r.width, v.h / r.height), px = cx ?? r.left + r.width / 2, py = cy ?? r.top + r.height / 2;
  const mx = v.x + v.w / 2 + (px - (r.left + r.width / 2)) * s, my = v.y + v.h / 2 + (py - (r.top + r.height / 2)) * s, w = Math.min(kaGeo.w, Math.max(kaGeo.w / 24, v.w * f)), k = w / v.w;
  v.x = mx - (mx - v.x) * k; v.y = my - (my - v.y) * k; v.h *= k; v.w = w; krAjustar();
}
document.addEventListener('pointermove', e => {
  const a = krM.ar; if (!a || !krM.vb) return; const dx = e.clientX - a.x, dy = e.clientY - a.y;
  if (!krM.mov && Math.abs(dx) + Math.abs(dy) > 4) { krM.mov = true; $('#krLz')?.classList.add('mov'); $('#kaTip').classList.remove('on'); }
  if (krM.mov) { const r = $('#krSvg').getBoundingClientRect(), s = Math.max(krM.vb.w / r.width, krM.vb.h / r.height); krM.vb.x = a.vx - dx * s; krM.vb.y = a.vy - dy * s; krAjustar(); }
});
document.addEventListener('pointerup', () => {
  const a = krM.ar; if (!a) return; const fue = krM.mov; krM.ar = null; krM.mov = false; $('#krLz')?.classList.remove('mov');
  if (!fue && a.t) { krM.sel = a.t.dataset.n; krMapaActualizar(false); }
});
$('#kRec').addEventListener('change', e => {
  if (e.target.id !== 'krPlazo') return; const v = Math.round(Number(e.target.value)); if (!(v >= 1 && v <= 60)) { e.target.value = krPlazo; return; }
  krPlazo = v; try { localStorage.setItem('recPlazo', v); } catch (x) {}
  $('#krCump').outerHTML = krCump(kRecR); $('#krHall').outerHTML = krHallazgos(kRecR); $('#krPlazo').focus();
});
$('#kRec').addEventListener('click', e => {
  const t = e.target.closest('[data-rt]'); if (t) { kRecTab = t.dataset.rt; $('#krTabla').outerHTML = krTabla(kRecR); krBarras(); return; }
  const m = e.target.closest('[data-km]'); if (m) { krM.met = m.dataset.km; document.querySelectorAll('#krMet button').forEach(x => x.classList.toggle('on', x === m)); krMapaActualizar(true); return; }
  const n = e.target.closest('#krPanel button[data-n]'); if (n) { krM.sel = n.dataset.n; krMapaActualizar(false); return; }
  if (e.target.closest('#krVer')) { $('#qDesde').value = $('#kDesde').value; $('#qHasta').value = $('#kHasta').value; $('#qNodo').value = ''; $('#qPed').value = ''; $('#qTipo [data-t=rec]').click(); irA('detalle'); }
});

// ---------- informe PDF y Excel de recojos ----------
const recFilasGrupo = (m, r2) => [...m].sort((a, b) => b[1].costo - a[1].costo).map(([k, v]) => [k, v.n, r2(v.costo), v.n && v.costo ? r2(v.costo / v.n) : null, v.dn ? r2(v.dsum / v.dn) : null]);
async function exportarRecojos() {
  const R = kRecR; if (!R || !R.items.length) return toast('No hay recojos para exportar en el período seleccionado', 'err');
  await xlsxLib();
  const wb = XLSX.utils.book_new(), r2 = n => Math.round(n * 100) / 100, cpp = R.pedidos && R.costo ? r2(R.costo / R.pedidos) : null;
  const filtros = [...document.querySelectorAll('#kAct .pill')].map(x => x.textContent).join(', ');
  const res = [['Indicador', 'Valor'], ['Filtros', filtros], ['Costo de recojos', r2(R.costo)], ['Pedidos recogidos', R.pedidos], ['Bultos', R.bultos], ['Guías', R.guias], ['Recojos con importe', R.registros], ['Nodos', R.nodos], ['Costo por pedido', cpp],
    ['Costo por pedido en despacho (referencia)', kRecCpp ? r2(kRecCpp) : null], ['Días de retorno promedio', R.dias != null ? r2(R.dias) : null], ['Recojos con importe y sin detalle', R.totalEnv - R.conDetalle], ['Pedidos de detalle sin recojo registrado', R.sinRecojo.reduce((x, g) => x + g.n, 0)], ['Recojos sin importe', R.sinImporte],
    ['Plazo objetivo (días)', krPlazo], ['Pedidos dentro del plazo', recCumple(R, krPlazo).dentro], ['Pedidos con fechas para medir el plazo', recCumple(R, krPlazo).n], ['Cumplimiento del plazo (%)', recCumple(R, krPlazo).n ? r2(recCumple(R, krPlazo).pct) : null], ['Recorrido completo, pedido a llegada a CTD (días)', R.totalCompleto != null ? r2(R.totalCompleto) : null],
    ['Costo del período anterior', kRecPrev && kRecPrev.costo ? r2(kRecPrev.costo) : null], ['Pedidos del período anterior', kRecPrev ? kRecPrev.pedidos : null], ['Días de retorno del período anterior', kRecPrev && kRecPrev.dias != null ? r2(kRecPrev.dias) : null]];
  const ws0 = XLSX.utils.aoa_to_sheet(res); ws0['!cols'] = [{ wch: 42 }, { wch: 28 }]; XLSX.utils.book_append_sheet(wb, ws0, 'RESUMEN');
  const nodoFila = ([k, v]) => { const p = kaProv.get(norm(KA_ALIAS[k] || k)); return [k, v.n, r2(v.costo), v.n && v.costo ? r2(v.costo / v.n) : null, v.dn ? r2(v.dsum / v.dn) : null, p ? p[0] : null, p ? p[1] : null]; };
  [['POR NODO', 'Nodo', R.grupos.nodo], ['POR PROVEEDOR', 'Proveedor', R.grupos.prov], ['POR CUENTA', 'Cuenta', R.grupos.cuenta], ['POR MOTIVO', 'Motivo', R.grupos.motivo]].forEach(([hoja, t, m]) => {
    const cab = [t, 'Pedidos', 'Costo', 'Costo por pedido', 'Días de retorno'], filas = hoja === 'POR NODO' ? [...m].sort((a, b) => b[1].costo - a[1].costo).map(nodoFila) : recFilasGrupo(m, r2);
    const ws = XLSX.utils.aoa_to_sheet([hoja === 'POR NODO' ? [...cab, 'Provincia', 'Departamento'] : cab, ...filas]); ws['!cols'] = [{ wch: 30 }, ...Array(6).fill({ wch: 16 })]; XLSX.utils.book_append_sheet(wb, ws, hoja);
  });
  const wse = XLSX.utils.aoa_to_sheet([['Etapa', 'Días promedio'], ...R.etapas.map(e => [e[0], e[1] != null ? r2(e[1]) : null]), ['Recorrido completo (pedido a llegada a CTD)', R.totalCompleto != null ? r2(R.totalCompleto) : null]]); wse['!cols'] = [{ wch: 46 }, { wch: 16 }]; XLSX.utils.book_append_sheet(wb, wse, 'ETAPAS');
  const det = [['Fecha reporte', 'Nodo', 'Cuenta', 'N° de pedido', 'Motivo de devolución', 'Bultos', 'Guía', 'Proveedor', 'Fecha solicitud CX', 'Fecha llegada CTD', 'Días de retorno', 'Costo asignado']]
    .concat(R.items.filter(x => !x.synth && x.cuenta !== '(sin detalle)').sort((a, b) => (a.fecha + a.nodo).localeCompare(b.fecha + b.nodo)).map(x => [x.fecha, x.nodo, x.cuenta, x.pedido, x.motivo, x.bultos, x.guia, x.prov, x.sol, x.lleg, x.dias, x.costo ? r2(x.costo) : null]));
  const wsd = XLSX.utils.aoa_to_sheet(det); wsd['!cols'] = det[0].map(() => ({ wch: 18 })); XLSX.utils.book_append_sheet(wb, wsd, 'DETALLE');
  XLSX.writeFile(wb, `recojos_${$('#kDesde').value || 'inicio'}_${$('#kHasta').value || 'hoy'}.xlsx`);
}
function informeRecojosPDF() {
  const R = kRecR; if (!R || !R.items.length) return toast('No hay recojos para el informe en el período seleccionado', 'err');
  const w = window.open('', '_blank'); if (!w) return toast('El navegador bloqueó la ventana del informe. Habilitar las ventanas emergentes para este sitio.', 'err');
  const filtros = [...document.querySelectorAll('#kAct .pill')].map(x => x.textContent).join(', '), cpp = R.pedidos && R.costo ? R.costo / R.pedidos : 0;
  const tabla = (t, cab, filas) => `<h2>${t}</h2><table><tr>${cab.map((c, i) => `<th${i ? ' class="r"' : ''}>${c}</th>`).join('')}</tr>${filas.map(f => `<tr>${f.map((c, i) => `<td${i ? ' class="r"' : ''}>${c}</td>`).join('')}</tr>`).join('')}</table>`;
  const filas = m => [...m].sort((a, b) => b[1].costo - a[1].costo).slice(0, 20).map(([k, v]) => [esc(k), fmtN(v.n), v.costo ? fmt(v.costo) : '-', v.n && v.costo ? fmt(v.costo / v.n) : '-', v.dn ? (v.dsum / v.dn).toFixed(1) : '-']);
  const obs = [];
  if (R.totalEnv && R.conDetalle < R.totalEnv) obs.push(`${R.totalEnv - R.conDetalle} ${R.totalEnv - R.conDetalle === 1 ? 'recojo con importe sin detalle asociado' : 'recojos con importe sin detalle asociado'}: ${R.sinDetalle.map(g => `${g.nodo} ${fdmy(g.fecha).slice(0, 5)}`).join(', ')}.`);
  if (R.sinRecojo.length) obs.push(`Detalle sin recojo registrado en el período: ${R.sinRecojo.map(g => `${g.nodo} ${fdmy(g.fecha).slice(0, 5)} (${g.n})`).join(', ')}.`);
  if (R.sinImporte) obs.push(`${R.sinImporte} ${R.sinImporte === 1 ? 'recojo sin importe registrado' : 'recojos sin importe registrado'}.`);
  const hallP = recHallazgos(R, kRecPrev), CP = recCumple(R, krPlazo), lentos = krLentos(R), mapa = document.querySelector('#krSvg') && document.querySelector('#krSvg .ka-prov.d') ? document.querySelector('#krSvg').outerHTML.replace(/class="([^"]*)"/g, (m, c) => 'class="' + c.replace(/\bsel\b/g, '').trim() + '"').replace(/ data-tip="[^"]*"/g, '') : '';
  const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Informe de recojos y devoluciones</title><style>
    @page { size:A4; margin:14mm } * { box-sizing:border-box } body { font:11px/1.45 Inter,system-ui,Arial,sans-serif; color:#0f172a; margin:0 }
    header { display:flex; justify-content:space-between; align-items:center; border-bottom:3px solid #059669; padding-bottom:10px; margin-bottom:12px } header img { height:42px } h1 { font-size:20px; margin:0 } .sub { color:#64748b; font-size:11px }
    h2 { font-size:13px; margin:18px 0 6px; border-left:4px solid #2563eb; padding-left:8px; break-after:avoid } table { width:100%; border-collapse:collapse; margin-bottom:6px } th { background:#f1f5f9; text-align:left; padding:5px 7px; font-size:10px; text-transform:uppercase; letter-spacing:.04em; color:#475569 }
    td { padding:4px 7px; border-top:1px solid #e2e8f0 } th.r, td.r { text-align:right; font-variant-numeric:tabular-nums } tr { break-inside:avoid }
    .kp { display:grid; grid-template-columns:repeat(3,1fr); gap:8px } .kp div { border:1px solid #e2e8f0; border-radius:8px; padding:8px 10px } .kp span { display:block; font-size:9px; color:#94a3b8; font-weight:700; text-transform:uppercase; letter-spacing:.05em } .kp b { font-size:17px } .kp small { display:block; color:#64748b }
    ul { margin:4px 0 0 16px; padding:0 } li { margin:3px 0 } footer { margin-top:18px; color:#94a3b8; font-size:9px; border-top:1px solid #e2e8f0; padding-top:6px }
    .mapa { break-inside:avoid } .mapa svg { width:100%; height:300px; background:#eef2f8; border-radius:10px; display:block } .ka-prov { stroke:#fff; stroke-width:.6; vector-effect:non-scaling-stroke } .kr-mk circle { fill:#fff; stroke-width:3.4 } .kr-mk text { font-size:11.5px; font-weight:700; fill:#0f172a; text-anchor:middle; dominant-baseline:central }
    .ley { color:#64748b; font-size:10px; margin-top:4px }
  </style></head><body>
    <header><div><h1>Informe de recojos y devoluciones</h1><div class="sub">${esc(filtros)}</div></div><img src="${esc(new URL('logo.png', location.href).href)}" alt="Dinet"></header>
    <div class="kp">
      <div><span>Costo de recojos</span><b>${R.costo ? fmt(R.costo) : '-'}</b><small>${R.registros} ${R.registros === 1 ? 'recojo' : 'recojos'} con importe, ${R.nodos} ${R.nodos === 1 ? 'nodo' : 'nodos'}</small></div>
      <div><span>Pedidos recogidos</span><b>${fmtN(R.pedidos)}</b><small>${fmtN(R.bultos)} bultos${R.guias ? `, ${R.guias} guías` : ''}</small></div>
      <div><span>Costo por pedido</span><b>${cpp ? fmt(cpp) : '-'}</b><small>${kRecCpp ? `Despacho: ${fmt(kRecCpp)}` : 'Sin despachos en el período para comparar'}</small></div>
      <div><span>Días de retorno (promedio)</span><b>${R.dias != null ? R.dias.toFixed(1) : '-'}</b><small>Solicitud CX a llegada a CTD</small></div>
      <div><span>Nodo con mayor tiempo</span><b>${lentos.length ? esc(lentos[0][0]) : '-'}</b><small>${lentos.length ? lentos[0][1].toFixed(1) + ' días' : ''}</small></div>
      <div><span>Nodo con menor tiempo</span><b>${lentos.length ? esc(lentos[lentos.length - 1][0]) : '-'}</b><small>${lentos.length ? lentos[lentos.length - 1][1].toFixed(1) + ' días' : ''}</small></div>
    </div>
    ${obs.length ? `<h2>Observaciones</h2><ul>${obs.map(o => `<li>${esc(o)}</li>`).join('')}</ul>` : ''}
    ${hallP.length ? `<h2>Hallazgos</h2><ul>${hallP.map(([, t]) => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}
    ${R.etapas.some(e => e[1] != null) ? tabla('Tiempo por etapa del recorrido', ['Etapa', 'Días promedio'], [...R.etapas.map(e => [esc(e[0]), e[1] != null ? e[1].toFixed(1) : '-']), ['<b>Recorrido completo (pedido a llegada a CTD)</b>', R.totalCompleto != null ? '<b>' + R.totalCompleto.toFixed(1) + '</b>' : '-']]) : ''}
    ${CP.n ? `<h2>Cumplimiento de plazo</h2><p><b>${CP.pct.toFixed(0)} %</b> de los pedidos (${CP.dentro} de ${CP.n}) retornaron dentro del plazo objetivo de ${krPlazo} días, medido desde la solicitud CX.${CP.fuera.length ? ` Pedidos de mayor duración: ${CP.fuera.slice(0, 5).map(x => `${esc(x.pedido || '-')} (${esc(x.nodo)}, ${x.dias} días)`).join('; ')}.` : ''}</p>` : ''}
    ${mapa ? `<h2>Días de retorno por provincia</h2><div class="mapa">${mapa}<div class="ley">Verde: menor tiempo de retorno. Rojo: mayor tiempo de retorno. Cada cifra corresponde al promedio de días de la provincia.</div></div>` : ''}
    ${tabla('Por nodo', ['Nodo', 'Pedidos', 'Costo', 'Costo por pedido', 'Días de retorno'], filas(R.grupos.nodo))}
    ${tabla('Por proveedor', ['Proveedor', 'Pedidos', 'Costo', 'Costo por pedido', 'Días de retorno'], filas(R.grupos.prov))}
    ${tabla('Por cuenta', ['Cuenta', 'Pedidos', 'Costo', 'Costo por pedido', 'Días de retorno'], filas(R.grupos.cuenta))}
    ${tabla('Por motivo de devolución', ['Motivo', 'Pedidos', 'Costo', 'Costo por pedido', 'Días de retorno'], filas(R.grupos.motivo))}
    <footer>Informe generado el ${new Date().toLocaleString('es-PE')} | Red Troncal, Dinet Logística</footer>
  </body></html>`;
  w.document.open(); w.document.write(html); w.document.close(); setTimeout(() => { w.focus(); w.print(); }, 600);
}

$('#kExport').onclick = async () => {
  if (kst.modo === 'r') return exportarRecojos();
  if (!kpiResultado) return; await xlsxLib();
  const wb = XLSX.utils.book_new(), r2 = n => Math.round(n * 100) / 100;
  const aoaD = [['Fecha','Registros','Pedidos','Bultos','Cajas','Costo contado','Costo crédito','Costo total','Costo por pedido','Registros sin importe']]
    .concat(diasResultado.map(d => [d.fecha, d.registros, d.pedidos, d.bultos, d.cajas, r2(d.contado), r2(d.credito), r2(d.costo), d.pedConImporte ? r2(d.costo / d.pedConImporte) : null, d.sinImporte]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaD), 'POR DIA');
  const cu = [['Cuenta', 'Costo', 'Pedidos', 'Bultos', 'Costo por pedido', 'Costo por bulto']].concat([...kpiResultado.cuentas].sort((a, b) => b[1].costo - a[1].costo)
    .map(([k, v]) => [k, r2(v.costo), v.pedidos, v.bultos, v.pedidos ? r2(v.costo / v.pedidos) : null, v.bultos ? r2(v.costo / v.bultos) : null]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cu), 'POR CUENTA');
  [['POR NODO', kpiResultado.nodos, 'Nodo'], ['POR AGENCIA', kpiResultado.agencias, 'Agencia'], ['POR ZONA', kpiResultado.zonas, 'Zona'], ['POR TRANSPORTE', kpiResultado.transportes, 'Transporte']].forEach(([hoja, mapa, t]) => {
    const aoa = [[t, 'Costo', 'Pedidos', 'Cajas', 'Costo contado', 'Costo crédito', 'Costo por pedido', 'Tarifa por caja']].concat([...mapa].sort((a, b) => b[1].costo - a[1].costo)
      .map(([k, v]) => [k, r2(v.costo), v.pedidos, v.cajas, r2(v.contado), r2(v.credito), v.pedidos ? r2(v.costo / v.pedidos) : null, v.cajas ? r2(v.costo / v.cajas) : null]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), hoja);
  });
  XLSX.writeFile(wb, `kpis_${$('#kDesde').value || 'inicio'}_${$('#kHasta').value || 'hoy'}.xlsx`);
};

// ---------- liquidaciones ----------
const LQ_DEF = { liquidador: 'BENITES VEGA LUIS MIGUEL', area: 'ECOMMERCE', moneda: 'SOLES', concepto: 'Flete interprovincial', concepto_recojo: 'Recojo interprovincial', tipo_doc: 'FACTURA', cuenta: '63111002', centro: '8003825', denominacion: 'Distribución Red Troncal', limite: '1500' };
const LQ_FILAS = 47;   // filas que tiene el formato (14 a 60)
let lqPar = { ...LQ_DEF }, lqLista = [], lqPend = [], lqSel = new Set(), lqHistData = [], lqVista = 'listos', lqPer = '', lqTexto = '', lqMes = today().slice(0, 7), lqGran = 'd', lqLiq = { clave: '', rows: [] }, lqDashClave = '';
const lqScript = src => new Promise((ok, ko) => { const t = document.createElement('script'); t.src = src; t.onload = ok; t.onerror = () => ko(new Error('No se pudo cargar una librería. Verificar la conexión a internet.')); document.head.appendChild(t); });
async function lqLibs() {
  if (!window.ExcelJS) await lqScript('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js');
  if (!window.JSZip) await lqScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
}
async function lqAbrir() {
  const { data, error } = await sb.from('parametros').select('clave,valor');
  const falta = !!error || !!(await sb.from('envios').select('liquidacion_id').limit(1)).error;
  $('#lqSetup').classList.toggle('hide', !falta); $('#lqMain').classList.toggle('hide', falta); $('#lqPar').classList.toggle('hide', falta);
  if (falta) return;
  (data || []).forEach(r => { const k = r.clave.replace(/^lq_/, ''); if (k in LQ_DEF) lqPar[k] = r.valor; });
  lqBuscar(); lqHistorial();
}
const lqLimite = () => Number(lqPar.limite) || 1500;
let lqSeq = 0, lqOcupado = false;
async function lqBuscar(opc) {
  lqCargarLiq();
  const mi = ++lqSeq, crear = () => {
    let q = sb.from('envios').select('*').is('liquidacion_id', null).eq('modo_pago', 'CONTADO').order('fecha').order('transporte').order('nodo').order('id');
    if ($('#lqTrans').value) q = q.eq('transporte', $('#lqTrans').value);
    if ($('#lqMot').value) q = q.eq('motivo', $('#lqMot').value);
    return q;
  };
  let data; try { data = await traerTodo(crear); } catch (error) { return flash($('#lqMsg'), error.message, 'err'); }
  if (mi !== lqSeq) return;
  const sel = $('#lqTrans').value;
  if (!sel) $('#lqTrans').innerHTML = '<option value="">Todos</option>' + [...new Set(data.map(r => r.transporte).filter(Boolean))].sort().map(t => `<option>${esc(t)}</option>`).join('');
  const lqLista0 = lqLista; lqLista = data.filter(regListo); lqPend = data.filter(r => !regListo(r));
  if (opc && opc.conservar) { const antes = new Set(lqLista0.map(r => r.id)); lqSel = new Set(lqLista.filter(r => lqSel.has(r.id) || !antes.has(r.id)).map(r => r.id)); }
  else lqSel = new Set(lqLista.map(r => r.id));
  lqPintar();
}
// Reparte los registros en archivos: cada archivo es de UNA sola fecha de despacho y no supera el monto máximo ni las filas del formato.
// Dentro de una fecha se llena cada archivo lo más posible (primero los montos grandes, en el primer archivo donde quepan).
function lqPartir(list) {
  const lim = Math.round(lqLimite() * 100), cts = e => Math.round(Number(e.importe) * 100), porFecha = new Map(), grupos = [];
  list.forEach(e => { const f = String(e.fecha).slice(0, 10); if (!porFecha.has(f)) porFecha.set(f, []); porFecha.get(f).push(e); });
  [...porFecha.keys()].sort().forEach(f => {
    const cajas = [];   // { filas, suma }
    [...porFecha.get(f)].sort((a, b) => cts(b) - cts(a)).forEach(e => {
      const imp = cts(e), c = cajas.find(x => x.suma + imp <= lim && x.filas.length < LQ_FILAS);
      if (c) { c.filas.push(e); c.suma += imp; } else cajas.push({ filas: [e], suma: imp });
    });
    cajas.forEach(c => grupos.push(c.filas.sort((a, b) => list.indexOf(a) - list.indexOf(b))));
  });
  return grupos;
}
const lqF = r => String(r.fecha).slice(0, 10);
const lqEnVista = r => lqF(r).startsWith(lqMes) && (!lqPer || lqF(r) === lqPer);
const lqBase = () => { const q = norm(lqTexto); return lqLista.filter(r => lqEnVista(r) && (!q || norm(r.nodo + ' ' + (r.factura || '')).includes(q))); };
const lqElegidos = () => lqBase().filter(r => lqSel.has(r.id));
const lqFechaTxt = f => `${diaC(f)} ${f.slice(0, 4)}`;
const LFIL = {
  trans: { sel: '#lqTrans', def: 'todos los transportes', txt: v => v },
  mot: { sel: '#lqMot', def: 'despachos y recojos', txt: v => v === 'DESPACHO' ? 'solo despachos' : 'solo recojos', ops: [['', 'Despachos y recojos'], ['DESPACHO', 'Solo despachos'], ['RECOJO', 'Solo recojos']] }
};
const lqOpciones = k => k === 'trans' ? [['', 'Todos los transportes'], ...[...$('#lqTrans').options].filter(o => o.value).map(o => [o.value, o.value])] : LFIL[k].ops;
function lqFrase() {
  let n = 0;
  Object.keys(LFIL).forEach(k => { const v = $(LFIL[k].sel).value, tk = document.querySelector(`.rx-tk[data-lk="${k}"]`); tk.textContent = v ? LFIL[k].txt(v) : LFIL[k].def; tk.classList.toggle('set', !!v); if (v) n++; });
  $('#lqLimp').hidden = !n;
}
// ----- vista por mes y resumen de liquidaciones (por fecha de despacho) -----
const LQ_MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'], LQ_DIAS = ['D', 'L', 'M', 'M', 'J', 'V', 'S'];
const lqDiasMes = ym => new Date(Number(ym.slice(0, 4)), Number(ym.slice(5, 7)), 0).getDate();
const lqDt = f => { const [y, m, d] = f.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const lqLunes = f => { const t = lqDt(f); t.setUTCDate(t.getUTCDate() - ((t.getUTCDay() + 6) % 7)); return t.toISOString().slice(0, 10); };
const lqKilo = n => n >= 1000 ? (n / 1000).toFixed(1).replace('.0', '') + 'k' : String(Math.round(n));
const lqSuma = (rows, f) => rows.reduce((s, r) => s + (f(r) ? Number(r.importe) || 0 : 0), 0);
const lqLiqClave = () => `liq|${lqMes.slice(0, 4)}|${$('#lqTrans').value}|${$('#lqMot').value}`;
function lqCargarLiq() {
  const clave = lqLiqClave(), anio = lqMes.slice(0, 4), t = $('#lqTrans').value, m = $('#lqMot').value;
  memo(clave, 300000, () => traerTodo(() => { let q = sb.from('envios').select('fecha,importe,liquidacion_id').not('liquidacion_id', 'is', null).gte('fecha', anio + '-01-01').lte('fecha', anio + '-12-31').order('id'); if (t) q = q.eq('transporte', t); if (m) q = q.eq('motivo', m); return q; }))
    .then(rows => { if (clave === lqLiqClave()) { lqLiq = { clave, rows }; lqPintar(); } }).catch(() => {});
}
function lqMesPintar() {
  const hoy = today(), liqM = lqLiq.clave === lqLiqClave() ? lqLiq.rows.filter(r => lqF(r).startsWith(lqMes)) : null, listosM = lqLista.filter(r => lqF(r).startsWith(lqMes)), pendM = lqPend.filter(r => lqF(r).startsWith(lqMes));
  $('#lqMesT').textContent = `${LQ_MESES[Number(lqMes.slice(5, 7)) - 1]} ${lqMes.slice(0, 4)}`; $('#lqNext').disabled = lqMes >= hoy.slice(0, 7); $('#lqPrev').disabled = lqMes <= '2025-01'; $('#lqTodo').classList.toggle('on', !lqPer);
  $('#lqMt').innerHTML = `<div><small>Liquidado en el mes</small><b>${liqM ? fmt(lqSuma(liqM, () => true)) : '—'}</b></div><div><small>Listo por liquidar</small><b>${fmt(lqSuma(listosM, () => true))}</b></div><div><small>Incompletos</small><b>${pendM.length}</b></div>`;
  const otros = lqLista.filter(r => !lqF(r).startsWith(lqMes)), ob = $('#lqOtros'); ob.classList.toggle('hide', !otros.length);
  if (otros.length) { ob.textContent = `${otros.length} ${otros.length === 1 ? 'listo en otro mes' : 'listos en otros meses'}`; ob.dataset.m = otros.map(lqF).sort()[0].slice(0, 7); }
  const dias = lqDiasMes(lqMes), porDia = (rows, d) => rows.filter(r => lqF(r) === `${lqMes}-${String(d).padStart(2, '0')}`);
  $('#lqTira').innerHTML = Array.from({ length: dias }, (_, i) => {
    const d = i + 1, f = `${lqMes}-${String(d).padStart(2, '0')}`, fut = f > hoy, liq = liqM ? lqSuma(porDia(liqM, d), () => true) : 0, pen = lqSuma(porDia(listosM, d), () => true), inc = porDia(pendM, d).length;
    const monto = liq || pen, cls = [fut ? 'fut' : '', !fut && !liq && !pen && !inc ? 'nada' : '', liq ? 'liq' : '', pen ? 'pen' : '', inc ? 'inc' : '', lqPer === f ? 'on' : '', f === hoy ? 'hoy' : ''].filter(Boolean).join(' ');
    return `<button type="button" class="lqm-dc ${cls}" data-f="${f}" ${fut ? 'disabled' : ''} title="${d} de ${LQ_MESES[Number(lqMes.slice(5, 7)) - 1].toLowerCase()}"><small>${LQ_DIAS[lqDt(f).getUTCDay()]}</small><b>${d}</b><i></i><em>${monto ? lqKilo(monto) : ''}</em></button>`;
  }).join('');
  const clave = [lqMes, lqGran, lqLiq.clave, lqLiq.rows.length, lqLista.length].join('|'); if (clave !== lqDashClave) { lqDashClave = clave; lqDash(); }
}
function lqBuckets() {
  const anio = lqMes.slice(0, 4), hoy = today(), liq = lqLiq.clave === lqLiqClave() ? lqLiq.rows : [], B = new Map();
  const clave = f => lqGran === 'd' ? f : lqGran === 's' ? lqLunes(f) : f.slice(0, 7), cab = (k, f) => { if (!B.has(k)) B.set(k, { k, ini: f, fin: f, liq: 0, pen: 0, ids: new Set(), reg: 0 }); const b = B.get(k); if (f < b.ini) b.ini = f; if (f > b.fin) b.fin = f; return b; };
  if (lqGran === 'm') { for (let m = 1; m <= 12; m++) { const f = `${anio}-${String(m).padStart(2, '0')}-01`; if (f.slice(0, 7) <= hoy.slice(0, 7)) cab(f.slice(0, 7), f); } }
  else for (let d = 1; d <= lqDiasMes(lqMes); d++) { const f = `${lqMes}-${String(d).padStart(2, '0')}`; if (f <= hoy) cab(clave(f), f); }
  const dentro = f => lqGran === 'm' ? f.startsWith(anio) : f.startsWith(lqMes);
  liq.forEach(r => { const f = lqF(r); if (!dentro(f)) return; const b = cab(clave(f), f); b.liq += Number(r.importe) || 0; b.ids.add(r.liquidacion_id); b.reg++; });
  lqLista.forEach(r => { const f = lqF(r); if (!dentro(f)) return; cab(clave(f), f).pen += Number(r.importe) || 0; });
  return [...B.values()].sort((x, y) => x.k.localeCompare(y.k)).map((b, i) => {
    const d = Number(b.ini.slice(8, 10)), mc = LQ_MESES[Number(b.ini.slice(5, 7)) - 1];
    return { ...b, n: b.ids.size, t: lqGran === 'd' ? `${fdmy(b.ini).slice(0, 5)}` : lqGran === 's' ? `${d} al ${Number(b.fin.slice(8, 10))} de ${mc.slice(0, 3).toLowerCase()}` : `${mc} ${b.ini.slice(0, 4)}`, l: lqGran === 'd' ? `${LQ_DIAS[lqDt(b.ini).getUTCDay()]} ${d}` : lqGran === 's' ? `Sem ${i + 1}` : mc.slice(0, 3) };
  });
}
function lqDash() {
  const B = lqBuckets(), tot = B.reduce((s, b) => s + b.liq, 0), nl = B.reduce((s, b) => s + b.n, 0), reg = B.reduce((s, b) => s + b.reg, 0), con = B.filter(b => b.liq > 0), prom = con.length ? tot / con.length : 0, mx = con.reduce((a, b) => b.liq > (a ? a.liq : 0) ? b : a, null);
  const un = { d: 'día', s: 'semana', m: 'mes' }[lqGran], unp = { d: 'días', s: 'semanas', m: 'meses' }[lqGran];
  $('#lqdSub').textContent = lqGran === 'm' ? `Año ${lqMes.slice(0, 4)}` : `${LQ_MESES[Number(lqMes.slice(5, 7)) - 1]} ${lqMes.slice(0, 4)}`;
  $('#lqdK').innerHTML = `<div><small>Total liquidado</small><b>${fmt(tot)}</b><em>${lqGran === 'm' ? 'en el año' : 'en el mes'}</em></div><div><small>Liquidaciones</small><b>${nl}</b><em>${fmtN(reg)} registros</em></div><div><small>Promedio por ${un}</small><b>${fmt(prom)}</b><em>${con.length} ${con.length === 1 ? un : unp} con liquidación</em></div><div><small>Mayor ${un}</small><b>${mx ? fmt(mx.liq) : '—'}</b><em>${mx ? mx.t : ''}</em></div>`;
  const W = 900, H = 250, pl = 46, pb = 26, pt = 10, n = B.length || 1, max = Math.max(1, ...B.map(b => b.liq + b.pen)), ymax = Math.ceil(max / 1000) * 1000 || 1000, paso = (W - pl - 10) / n, bw = Math.min(46, paso * .62);
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Liquidaciones por ${un}">`;
  for (let i = 0; i <= 4; i++) { const y = pt + (H - pt - pb) * (1 - i / 4); s += `<line class="gl" x1="${pl}" x2="${W - 10}" y1="${y}" y2="${y}"/><text class="ax" x="${pl - 8}" y="${y + 4}" text-anchor="end">${lqKilo(ymax * i / 4)}</text>`; }
  B.forEach((b, i) => {
    const x = pl + paso * i + (paso - bw) / 2, hl = (H - pt - pb) * b.liq / ymax, hp = (H - pt - pb) * b.pen / ymax, y0 = H - pb;
    s += `<g class="b" data-i="${i}">${hl ? `<rect x="${x}" y="${y0 - hl}" width="${bw}" height="${hl}" rx="5" fill="#0066cc"/>` : ''}${hp ? `<rect x="${x}" y="${y0 - hl - hp}" width="${bw}" height="${hp}" rx="5" fill="#b9d4f5"/>` : ''}<rect x="${x - 4}" y="${pt}" width="${bw + 8}" height="${H - pt - pb}" fill="transparent"/></g>`;
    if (n <= 16 || i % 2 === 0) s += `<text class="ax" x="${x + bw / 2}" y="${H - 8}" text-anchor="middle">${b.l}</text>`;
  });
  const ch = $('#lqdChart'); ch.querySelectorAll('svg').forEach(e => e.remove()); ch.insertAdjacentHTML('afterbegin', s + '</svg>'); ch._b = B;
  $('#lqdTabla').innerHTML = `<tr><th>${{ d: 'Día', s: 'Semana', m: 'Mes' }[lqGran]}</th><th class="num">Liquidaciones</th><th class="num">Registros</th><th class="num">Total liquidado</th><th>Participación</th></tr>` +
    ([...B].reverse().filter(b => b.liq).slice(0, 10).map(b => `<tr><td><b>${b.t}</b></td><td class="num">${b.n}</td><td class="num">${b.reg}</td><td class="num"><b>${fmt(b.liq)}</b></td><td><span class="lqd-pb"><i style="width:${tot ? b.liq / tot * 100 : 0}%"></i></span> ${tot ? Math.round(b.liq / tot * 100) : 0} %</td></tr>`).join('') || '<tr><td colspan="5"><div class="empty"><b>Sin liquidaciones en este período</b><span>Las liquidaciones descargadas se resumirán aquí.</span></div></td></tr>');
}
$('#lqdChart').addEventListener('mousemove', e => {
  const ch = e.currentTarget, g = e.target.closest('.b'), t = $('#lqdTip'); if (!g || !ch._b) { t.style.opacity = 0; return; }
  const b = ch._b[Number(g.dataset.i)], r = ch.getBoundingClientRect();
  t.innerHTML = `<b>${b.t}</b>Liquidado: ${fmt(b.liq)}<br>${b.pen ? 'Listo por liquidar: ' + fmt(b.pen) + '<br>' : ''}${b.n} ${b.n === 1 ? 'liquidación' : 'liquidaciones'} · ${b.reg} registros`;
  t.style.left = Math.min(e.clientX - r.left + 14, r.width - 230) + 'px'; t.style.top = Math.max(0, e.clientY - r.top - 70) + 'px'; t.style.opacity = 1;
});
$('#lqdChart').addEventListener('mouseleave', () => { $('#lqdTip').style.opacity = 0; });
$('#lqGran').addEventListener('click', e => { const b = e.target.closest('button[data-g]'); if (!b) return; lqGran = b.dataset.g; document.querySelectorAll('#lqGran button').forEach(x => x.classList.toggle('on', x === b)); lqPintar(); });
const lqIrMes = ym => { const cambiaAnio = ym.slice(0, 4) !== lqMes.slice(0, 4); lqMes = ym; lqPer = ''; if (cambiaAnio) lqCargarLiq(); lqPintar(); };
$('#lqPrev').onclick = () => { const [y, m] = lqMes.split('-').map(Number); lqIrMes(m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, '0')}`); };
$('#lqNext').onclick = () => { const [y, m] = lqMes.split('-').map(Number); lqIrMes(m === 12 ? `${y + 1}-01` : `${y}-${String(m + 1).padStart(2, '0')}`); };
$('#lqTodo').onclick = () => { lqPer = ''; lqPintar(); };
$('#lqOtros').onclick = e => lqIrMes(e.currentTarget.dataset.m);
$('#lqTira').addEventListener('click', e => { const b = e.target.closest('.lqm-dc'); if (!b || b.disabled) return; lqPer = lqPer === b.dataset.f ? '' : b.dataset.f; lqPintar(); });
let lqPendVista = [];
function lqPintar() {
  lqFrase(); lqMesPintar();
  const nListos = lqLista.filter(lqEnVista).length, pendM = lqPend.filter(lqEnVista), histM = lqHistData.filter(l => String(l.fecha).slice(0, 7) === lqMes);
  const base = lqBase(), elegidos = lqElegidos(), total = elegidos.reduce((s, r) => s + Number(r.importe), 0), grupos = lqPartir(elegidos), lim = lqLimite();
  const idx = new Map(); grupos.forEach((g, i) => g.forEach(r => idx.set(r.id, i + 1)));
  $('#lqSeg').innerHTML = [['listos', 'Listos para liquidar', nListos], ['inc', 'Incompletos', pendM.length], ['hist', 'Historial', histM.length]]
    .map(([v, t, n]) => `<button type="button" role="tab" data-v="${v}" aria-selected="${lqVista === v}" class="${lqVista === v ? 'on' : ''}">${t} <b>${n}</b>${v === 'inc' && n ? '<i></i>' : ''}</button>`).join('');
  $('#lqFrase').classList.toggle('sinfiltros', lqVista === 'hist');
  $('#lqVListos').classList.toggle('hide', lqVista !== 'listos'); $('#lqVInc').classList.toggle('hide', lqVista !== 'inc'); $('#lqVHist').classList.toggle('hide', lqVista !== 'hist');
  if (lqVista === 'listos') {
    $('#lqCuenta').innerHTML = '';
    const mostrar = [...base].sort((a, b) => lqF(a).localeCompare(lqF(b)) || ((idx.get(a.id) || 99) - (idx.get(b.id) || 99)) || Number(b.importe) - Number(a.importe));
    const todos = mostrar.length > 0 && mostrar.every(r => lqSel.has(r.id)), nArch = `${grupos.length} ${grupos.length === 1 ? 'archivo' : 'archivos'}`;
    const barra = $('#lqBarra'); barra.hidden = !nListos;
    barra.innerHTML = `<div class="l"><input type="checkbox" id="lqTodos" ${todos ? 'checked' : ''} aria-label="Seleccionar todos"><span><b>${elegidos.length} de ${base.length}</b> seleccionados</span><span class="m">${fmt(total)} · ${nArch}</span></div>
      <div class="lq-ach">${grupos.map((g, i) => { const s = g.reduce((x, r) => x + Number(r.importe), 0); return `<span class="lq-fl" style="--p:${Math.min(100, s / lim * 100)}%"><b>${i + 1}</b> ${diaC(lqF(g[0]))} <span>·</span> ${fmt(s)}</span>`; }).join('') || '<span class="lq-vacio">Seleccionar registros para generar la liquidación</span>'}</div>
      <button type="button" class="lq-go" id="lqDescargar" ${elegidos.length && !lqOcupado ? '' : 'disabled'}>${lqOcupado ? 'Preparando…' : 'Descargar liquidación'}</button>`;
    const tt = $('#lqTodos'); if (tt) tt.indeterminate = elegidos.length > 0 && !todos;
    const cuenta = {}; elegidos.forEach(r => { const k = norm(r.factura); cuenta[k] = (cuenta[k] || 0) + 1; });
    const repetidas = Object.keys(cuenta).filter(k => cuenta[k] > 1), grandes = elegidos.filter(r => Number(r.importe) > lim).length;
    $('#lqAviso').innerHTML = [repetidas.length && `<b>N° de factura repetido:</b> ${repetidas.map(esc).join(', ')}. Verificar que no se pague dos veces.`, grandes && `${grandes} ${grandes === 1 ? 'registro supera' : 'registros superan'} el máximo y ${grandes === 1 ? 'irá' : 'irán'} solo en su archivo.`].filter(Boolean).join('<br>');
    const sf = new Map(); elegidos.forEach(r => sf.set(lqF(r), (sf.get(lqF(r)) || 0) + Number(r.importe))); let ult = '';
    $('#lqTabla').innerHTML = `<tr><th data-c="chk"></th><th data-c="fecha">Fecha</th><th data-c="nodo">Nodo</th><th data-c="factura">Factura</th><th class="num" data-c="importe">Importe</th><th data-c="arch">Archivo</th></tr>` +
      (mostrar.map(r => { const a = idx.get(r.id), f = lqF(r), cab = f !== ult ? (ult = f, `<tr class="lqg"><td colspan="6">${fdmy(f)}<span>${fmt(sf.get(f) || 0)}</span></td></tr>`) : '';
        return cab + `<tr data-id="${r.id}" class="${lqSel.has(r.id) ? 'sel' : ''}"><td data-c="chk"><input type="checkbox" data-id="${r.id}" ${lqSel.has(r.id) ? 'checked' : ''} aria-label="Incluir ${esc(r.nodo)}"></td><td data-c="fecha">${fdmy(r.fecha)}</td><td data-c="nodo"><div class="nodo">${esc(r.nodo)}${r.motivo === 'RECOJO' ? '<span class="tb rec">Recojo</span>' : ''}</div><div class="meta">${esc(r.transporte || '')}</div></td><td data-c="factura">${esc(r.factura)}</td><td class="num" data-c="importe"><b>${fmt(r.importe)}</b></td><td data-c="arch">${a ? `<span class="lq-arch"><i>${a}</i>Archivo ${a}</span>` : '<span class="lq-arch nn"><i>—</i>No incluido</span>'}</td><td data-c="meta">${[diaC(lqF(r)), r.factura, a ? 'Archivo ' + a : ''].filter(Boolean).map(esc).join(' · ')}</td></tr>`; }).join('') ||
        `<tr><td colspan="6"><div class="empty"><b>No hay registros listos para liquidar</b><span>Un registro está listo cuando tiene importe, N° de factura y foto, y aún no fue enviado.</span></div></td></tr>`);
    const pie = $('#lqPieL'); pie.hidden = !nListos; pie.innerHTML = `<span>Un archivo por fecha de despacho · máximo ${fmt(lim)} por archivo</span><button type="button" id="lqCambiar">Cambiar máximo</button>`;
    const bm = $('#lqBarM'); bm.hidden = !nListos;
    bm.innerHTML = `<div><b>${fmt(total)}</b><small>${elegidos.length} ${elegidos.length === 1 ? 'registro' : 'registros'} · ${nArch}</small></div><button type="button" class="lq-go" id="lqDescargarM" ${elegidos.length && !lqOcupado ? '' : 'disabled'}>${lqOcupado ? 'Preparando…' : 'Descargar'}</button>`;
  } else if (lqVista === 'inc') {
    const q = norm(lqTexto), pend = pendM.filter(r => !q || norm(r.nodo).includes(q)); lqPendVista = pend;
    $('#lqCuenta').innerHTML = `<b>${pend.length}</b> ${pend.length === 1 ? 'registro sin completar' : 'registros sin completar'}`;
    $('#lqPend').innerHTML = '<tr><th>Fecha</th><th>Nodo</th><th>Le falta</th><th></th></tr>' + (pend.map((r, i) => `<tr data-i="${i}"><td>${fdmy(r.fecha)}</td><td><div class="nodo">${esc(r.nodo)}${r.motivo === 'RECOJO' ? '<span class="tb rec">Recojo</span>' : ''}</div><div class="meta">${esc(r.transporte || '')}</div></td><td><span class="lq-falta">${[Number(r.importe) > 0 ? '' : 'Importe', r.factura ? '' : 'N° de factura', r.factura_archivo ? '' : 'Foto'].filter(Boolean).map(x => `<span>${x}</span>`).join('')}</span></td><td style="text-align:right"><button type="button" class="lnk" data-a="completar">Completar en Registros</button></td></tr>`).join('') ||
      '<tr><td colspan="4"><div class="empty"><b>No hay registros incompletos</b><span>Todo lo pendiente tiene importe, factura y foto.</span></div></td></tr>');
  } else lqHistPintar();
}
function lqHistPintar() {
  const q = norm(lqTexto), data = lqHistData.filter(l => String(l.fecha).slice(0, 7) === lqMes && (!q || norm(lqCod(l.numero)).includes(q)));
  $('#lqCuenta').innerHTML = `<b>${data.length}</b> ${data.length === 1 ? 'liquidación' : 'liquidaciones'}`;
  $('#lqHist').innerHTML = '<tr><th>N°</th><th>Fecha</th><th>Pago</th><th class="num">Registros</th><th class="num">Total</th><th></th></tr>' +
    (data.map(l => `<tr data-id="${l.id}"><td><b>${lqCod(l.numero)}</b></td><td>${fdmy(l.fecha)}</td><td>${l.modo_pago === 'CREDITO' ? 'Crédito' : l.modo_pago === 'CONTADO' ? 'Contado' : 'Contado y crédito'}</td><td class="num">${l.items}</td><td class="num"><b>${fmt(l.total)}</b></td>
      <td style="text-align:right"><span class="lq-acc"><button type="button" class="lnk" data-a="bajar">Descargar</button><button type="button" class="rib" data-act="lqmenu" aria-label="Más acciones" title="Más acciones"><svg viewBox="0 0 24 24" fill="currentColor"><circle cx="5" cy="12" r="1.8"/><circle cx="12" cy="12" r="1.8"/><circle cx="19" cy="12" r="1.8"/></svg></button></span></td></tr>`).join('') ||
    '<tr><td colspan="6"><div class="empty"><b>Aún no hay liquidaciones</b><span>Las liquidaciones descargadas se mostrarán en esta sección.</span></div></td></tr>');
}
// ----- eventos de la pantalla -----
$('#lqSeg').addEventListener('click', e => { const b = e.target.closest('button[data-v]'); if (!b) return; lqVista = b.dataset.v; $('#lqQ').value = ''; lqTexto = ''; lqPintar(); });
$('#lqQ').addEventListener('input', e => { lqTexto = e.target.value.trim(); lqPintar(); });
$('#lqBarra').addEventListener('change', e => { if (e.target.id !== 'lqTodos') return; lqBase().forEach(r => e.target.checked ? lqSel.add(r.id) : lqSel.delete(r.id)); lqPintar(); });
$('#lqTabla').addEventListener('change', e => {
  if (e.target.dataset.id) e.target.checked ? lqSel.add(e.target.dataset.id) : lqSel.delete(e.target.dataset.id);
  lqPintar();
});
$('#lqTabla').addEventListener('click', e => {   // tocar la fila marca o desmarca
  const tr = e.target.closest('tr'), cb = tr && tr.querySelector('input[data-id]');
  if (cb && e.target !== cb) { cb.checked = !cb.checked; cb.dispatchEvent(new Event('change', { bubbles: true })); }
});
const lqClickAccion = e => { if (e.target.closest('#lqDescargar, #lqDescargarM')) lqDescargar(); else if (e.target.closest('#lqCambiar')) $('#lqPar').click(); };
['#lqBarra', '#lqBarM', '#lqPieL'].forEach(q => $(q).addEventListener('click', lqClickAccion));
$('#lqPend').addEventListener('click', e => {
  const b = e.target.closest('button[data-a=completar]'); if (!b) return;
  const r = lqPendVista[Number(b.closest('tr').dataset.i)]; if (!r) return;
  $('#rDesde').value = lqF(r); $('#rHasta').value = lqF(r); $('#rNodo').value = r.nodo; irA('registros'); rBuscarYa();
});
// filtros escritos como frase (mismas piezas que Registros)
document.addEventListener('click', e => {
  if ($('#tab-liquidaciones').classList.contains('hide')) return;
  const dentro = e.target.closest('#lqFrase .rx-fb'), tk = e.target.closest('#lqFrase .rx-tk[data-pop]');
  document.querySelectorAll('#lqFrase .rx-pop.on').forEach(p => { if (p.closest('.rx-fb') !== dentro) p.classList.remove('on'); });
  if (tk) {
    const pop = $('#' + tk.dataset.pop), k = tk.dataset.lk, actual = $(LFIL[k].sel).value;
    pop.innerHTML = lqOpciones(k).map(([v, t]) => `<button type="button" class="rx-op${actual === v ? ' on' : ''}" data-lk="${k}" data-v="${esc(v)}">${esc(t)}${RX_CHECK}</button>`).join('');
    pop.classList.toggle('on'); return;
  }
  const op = e.target.closest('#lqFrase .rx-op');
  if (op) {
    op.closest('.rx-pop').classList.remove('on');
    $(LFIL[op.dataset.lk].sel).value = op.dataset.v; lqBuscar();
  }
});
$('#lqLimp').onclick = () => { Object.values(LFIL).forEach(f => { $(f.sel).value = ''; }); lqPer = ''; lqBuscar(); };

const lqNombreFoto = (r, usados) => {
  const ext = (String(r.factura_archivo).split('.').pop() || 'jpg').toLowerCase(), dm = fdmy(r.fecha).slice(0, 5).replace('/', '-');
  const base = `${r.factura} ${r.transporte || ''} ${dm}`.replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
  let n = base, i = 2; while (usados.has(n.toLowerCase())) n = `${base} (${i++})`; usados.add(n.toLowerCase());
  return `${n}.${ext}`;
};
// Excel en el formato de la empresa, a partir de la plantilla
async function lqExcel(filas, L) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await (await fetch('plantilla_liquidacion.xlsx?v=1')).arrayBuffer());
  const vacia = wb.worksheets.find(w => w.name === 'Hoja1'); if (vacia) wb.removeWorksheet(vacia.id);
  const ws = wb.worksheets[0], dia = f => { const [y, m, d] = String(f).slice(0, 10).split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  ws.getCell('F5').value = lqPar.liquidador; ws.getCell('F6').value = lqPar.area; ws.getCell('F8').value = lqPar.moneda;
  ws.getCell('F7').value = dia(L.fecha); ws.getCell('F7').numFmt = 'dd/mm/yyyy';
  for (let i = 0; i < LQ_FILAS; i++) {
    const f = 14 + i, r = filas[i], c = col => ws.getCell(col + f);
    if (!r) { 'ABCDEFGHIJK'.split('').forEach(k => { c(k).value = null; }); continue; }
    c('A').value = i + 1; c('B').value = dia(r.fecha); c('B').numFmt = 'dd/mm/yyyy'; c('C').value = r.transporte || ''; c('D').value = r.motivo === 'RECOJO' ? (lqPar.concepto_recojo || lqPar.concepto) : lqPar.concepto; c('E').value = r.nodo;
    c('F').value = lqPar.tipo_doc; c('G').value = Number(r.importe); c('H').value = r.factura;
    c('I').value = /^\d+$/.test(lqPar.centro) ? Number(lqPar.centro) : lqPar.centro; c('J').value = lqPar.denominacion; c('K').value = /^\d+$/.test(lqPar.cuenta) ? Number(lqPar.cuenta) : lqPar.cuenta;
  }
  ws.getCell('G61').value = { formula: 'SUM(G14:G60)', result: filas.reduce((x, r) => x + Number(r.importe), 0) };
  return wb.xlsx.writeBuffer();
}
// agrega al ZIP el Excel de una liquidación (suelto, afuera) y sus fotos (todas juntas en la carpeta FACTURAS)
async function lqCarpeta(zip, filas, L) {
  const fotos = zip.folder('FACTURAS'); zip.__usados = zip.__usados || new Set();
  // nombre del Excel: FORMATO RENDICIÓN - AÑO MES - DÍA (fecha del despacho); si hay otro igual en el ZIP, se numera
  const [fy, fm, fd] = String(filas[0].fecha).slice(0, 10).split('-'), mes = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'][Number(fm) - 1];
  const base = `FORMATO RENDICIÓN - ${fy} ${mes} - ${fd}`, previos = Object.keys(zip.files).filter(n => !n.includes('/') && n.endsWith('.xlsx') && n.startsWith(base)).length;
  zip.file(base + (previos ? ` (${previos + 1})` : '') + '.xlsx', await lqExcel(filas, L));
  const cola = [...filas]; let falla = 0;
  await Promise.all(Array.from({ length: 5 }, async () => {
    while (cola.length) {
      const r = cola.shift(), { data, error } = await sb.storage.from(BUCKET).download(r.factura_archivo);
      if (error || !data) { falla++; continue; }
      fotos.file(lqNombreFoto(r, zip.__usados), data);
    }
  }));
  return falla;
}
function lqBajar(blob, nombre) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
async function lqDescargar() {
  const elegidos = lqElegidos(); if (!elegidos.length) return;
  const grupos = lqPartir(elegidos), total = elegidos.reduce((s, r) => s + Number(r.importe), 0);
  if (!await confirmar('Descargar liquidación', `${elegidos.length} ${elegidos.length === 1 ? 'registro' : 'registros'} por ${fmt(total)} en ${grupos.length} ${grupos.length === 1 ? 'archivo' : 'archivos'}. Quedarán marcados como enviados.`, 'Descargar')) return;
  lqOcupado = true; lqPintar();
  const creadas = [];
  try {
    await lqLibs();
    const zip = new JSZip(); let fallas = 0;
    for (const g of grupos) {
      const modos = [...new Set(g.map(r => r.modo_pago))];
      const { data, error } = await sb.from('liquidaciones').insert({ fecha: today(), modo_pago: modos.length === 1 ? modos[0] : 'TODOS', total: Math.round(g.reduce((s, r) => s + Number(r.importe) * 100, 0)) / 100, items: g.length }).select().single();
      if (error) throw error; creadas.push(data);
      fallas += await lqCarpeta(zip, g, data);
    }
    for (let i = 0; i < grupos.length; i++) {
      const ids = grupos[i].map(r => r.id), { data: marcados, error } = await sb.from('envios').update({ liquidacion_id: creadas[i].id }).in('id', ids).is('liquidacion_id', null).select('id'); if (error) throw error;
      if (marcados.length !== ids.length) throw new Error('Algunos registros ya pertenecen a otra liquidación. Volver a buscar.');
    }
    lqBajar(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), `Liquidaciones ${fdmy(today()).replace(/\//g, '-')}.zip`);
    toast(`${grupos.length} ${grupos.length === 1 ? 'liquidación descargada' : 'liquidaciones descargadas'}` + (fallas ? `. ${fallas} ${fallas === 1 ? 'foto no se pudo incluir' : 'fotos no se pudieron incluir'}` : ''), fallas ? 'err' : undefined);
  } catch (e) {
    for (const L of creadas) { await sb.from('envios').update({ liquidacion_id: null }).eq('liquidacion_id', L.id); await sb.from('liquidaciones').delete().eq('id', L.id); }
    toast('No se pudo generar la liquidación: ' + (e.message || e), 'err');
  }
  lqOcupado = false; memoLimpiar('liq'); lqBuscar(); lqHistorial();
}
async function lqHistorial() {
  const { data, error } = await sb.from('liquidaciones').select('*').eq('anulada', false).order('numero', { ascending: false }).limit(50);
  if (error) return;
  lqHistData = data; lqPintar();
}
async function lqAnular(id, cod) {
  if (!await confirmar(`¿Anular ${cod}?`, 'Sus registros vuelven a estar disponibles para una nueva liquidación.', 'Anular', true)) return;
  const r1 = await sb.from('envios').update({ liquidacion_id: null }).eq('liquidacion_id', id); if (r1.error) return toast(r1.error.message, 'err');
  const r2 = await sb.from('liquidaciones').update({ anulada: true }).eq('id', id); if (r2.error) return toast(r2.error.message, 'err');
  toast(cod + ' anulada'); memoLimpiar('liq'); lqBuscar(); lqHistorial();
}
const lqMenu = $('#lqMenu'), lqCerrarMenu = () => { lqMenu.hidden = true; lqMenu.dataset.id = ''; };
$('#lqHist').addEventListener('click', async e => {
  const mn = e.target.closest('[data-act=lqmenu]');
  if (mn) {
    const id = mn.closest('tr').dataset.id; if (!lqMenu.hidden && lqMenu.dataset.id === id) return lqCerrarMenu();
    lqMenu.dataset.id = id; lqMenu.dataset.cod = mn.closest('tr').querySelector('b').textContent; lqMenu.hidden = false;
    const rc = mn.getBoundingClientRect(); lqMenu.style.top = Math.max(8, Math.min(rc.bottom + 4, innerHeight - lqMenu.offsetHeight - 8)) + 'px'; lqMenu.style.right = Math.max(8, innerWidth - rc.right) + 'px'; return;
  }
  const b = e.target.closest('button[data-a=bajar]'); if (!b) return;
  const id = b.closest('tr').dataset.id, cod = b.closest('tr').querySelector('b').textContent;
  b.disabled = true; b.textContent = 'Preparando…';
  try {
    await lqLibs();
    const { data: L } = await sb.from('liquidaciones').select('*').eq('id', id).single();
    const { data: filas, error } = await sb.from('envios').select('*').eq('liquidacion_id', id).order('fecha').order('transporte').order('nodo'); if (error) throw error;
    const zip = new JSZip(), fallas = await lqCarpeta(zip, filas, L);
    lqBajar(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), cod + '.zip');
    if (fallas) toast(`${fallas} ${fallas === 1 ? 'foto no se pudo incluir' : 'fotos no se pudieron incluir'}`, 'err');
  } catch (err) { toast('No se pudo descargar: ' + (err.message || err), 'err'); }
  b.disabled = false; b.textContent = 'Descargar';
});
lqMenu.addEventListener('click', async e => { const b = e.target.closest('button[data-a=anular]'); if (!b) return; const id = lqMenu.dataset.id, cod = lqMenu.dataset.cod; lqCerrarMenu(); await lqAnular(id, cod); });
document.addEventListener('click', e => { if (!lqMenu.hidden && !e.target.closest('#lqMenu, [data-act=lqmenu]')) lqCerrarMenu(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !lqMenu.hidden) lqCerrarMenu(); });
window.addEventListener('resize', lqCerrarMenu); window.addEventListener('scroll', lqCerrarMenu, true);
// datos fijos del formato
const LP = { liquidador: '#lpLiquidador', area: '#lpArea', moneda: '#lpMoneda', concepto: '#lpConcepto', concepto_recojo: '#lpConceptoRec', tipo_doc: '#lpTipo', cuenta: '#lpCuenta', centro: '#lpCentro', denominacion: '#lpDenom', limite: '#lpLimite' };
$('#lqPar').onclick = () => { for (const k in LP) $(LP[k]).value = lqPar[k]; $('#lpE').className = 'msg'; $('#lpOk').disabled = false; $('#lqParDlg').showModal(); };
$('#lpX').onclick = () => $('#lqParDlg').close();
$('#lqParF').onsubmit = async e => {
  e.preventDefault(); $('#lpOk').disabled = true;
  const nuevo = {}; for (const k in LP) nuevo[k] = $(LP[k]).value.trim();
  const { error } = await sb.from('parametros').upsert(Object.entries(nuevo).map(([k, v]) => ({ clave: 'lq_' + k, valor: v, updated_at: new Date().toISOString() })));
  $('#lpOk').disabled = false;
  if (error) { $('#lpE').textContent = error.message; $('#lpE').className = 'msg err'; return; }
  lqPar = { ...lqPar, ...nuevo }; $('#lqParDlg').close(); toast('Datos del formato guardados'); lqPintar();
};

// ---------- tiempo real: los cambios de otros usuarios aparecen solos ----------
let rtCanal = null, rtCaido = false, rtUltimo = Date.now(); const rtTimers = {};
const rtEspera = (clave, fn, ms) => { clearTimeout(rtTimers[clave]); rtTimers[clave] = setTimeout(fn, ms); };
const pestanaActual = () => ['cargar', 'registros', 'liquidaciones', 'kpis', 'detalle'].find(t => !$('#tab-' + t).classList.contains('hide'));
// ¿el registro entra en los filtros que tengo puestos en Registros?
const rtCumple = r => { const d = $('#rDesde').value, h = $('#rHasta').value, t = $('#rTrans').value, p = $('#rPago').value, m = $('#rMot').value, n = $('#rNodo').value.trim();
  return (!d || r.fecha >= d) && (!h || r.fecha <= h) && (!t || r.transporte === t) && (!p || r.modo_pago === p) && (!m || (r.motivo || 'DESPACHO') === m) && (!n || [r.nodo, r.agencia, r.placa, r.factura].some(v => norm(v).includes(norm(n)))); };
function rtRefrescarVista() {
  const t = pestanaActual();
  if (t === 'kpis') rtEspera('kpi', () => calcularKpis(true), 3000);
  if (t === 'liquidaciones' && !lqOcupado) rtEspera('lq', () => { lqBuscar({ conservar: true }); lqHistorial(); }, 800);
}
function rtEnvio(p) {
  rtUltimo = Date.now(); memoLimpiar('prev'); memoLimpiar('liq');
  const nuevo = p.new || {}, id = nuevo.id || (p.old || {}).id, r = regBase.find(x => x.id === id);
  if (p.eventType === 'UPDATE' && r) {
    const igual = Object.keys(nuevo).every(k => String(nuevo[k] ?? '') === String(r[k] ?? ''));
    if (!igual) {
      const enModal = $('#regDlg').open && rm.id === id, sucio = enModal && rmSucio();   // antes de actualizar el registro
      Object.assign(r, nuevo);
      const tr = document.querySelector(`#rTabla tr[data-id="${id}"]`);
      if (tr && !tr.contains(document.activeElement)) repintarFila(r);
      contarEstados(); pintarResumenRegistros();
      if (enModal && !esMio(id)) { if (!sucio) abrirReg(id); else toast('Otro usuario modificó este registro. Guardar o descartar los cambios propios.', 'err'); }
    }
  } else if (p.eventType === 'DELETE') {
    if (r) { rtEspera('reg', buscarRegistros, 400); if ($('#regDlg').open && rm.id === id) { $('#regDlg').close(); toast('Este registro fue eliminado por otro usuario', 'err'); } }
  } else if (!r && rtCumple(nuevo)) rtEspera('reg', buscarRegistros, 500);   // registro nuevo o que ahora entra en mis filtros
  rtRefrescarVista();
}
function rtLiq() { rtUltimo = Date.now(); memoLimpiar('liq'); rtEspera('reg', buscarRegistros, 500); rtRefrescarVista(); }
function iniciarTiempoReal() {
  if (rtCanal) return;
  rtCanal = sb.channel('red-troncal')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'envios' }, rtEnvio)
    .on('postgres_changes', { event: '*', schema: 'public', table: 'liquidaciones' }, rtLiq)
    .subscribe(estado => {
      if (estado === 'SUBSCRIBED') { if (rtCaido) { rtCaido = false; buscarRegistros(); rtRefrescarVista(); } }
      else if (['CHANNEL_ERROR', 'TIMED_OUT', 'CLOSED'].includes(estado)) rtCaido = true;
    });
}
// al volver a la pestaña después de un rato (celular, otra ventana) se pone al día
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState !== 'visible' || !rtCanal || Date.now() - rtUltimo < 60000) return;
  rtUltimo = Date.now(); buscarRegistros(); rtRefrescarVista();
});

{
  const svg = d => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="${d}"/></svg>`;
  const IC = { buscar: svg('M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z'), bajar: svg('M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4'), refrescar: svg('M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15'), mas: svg('M12 4v16m8-8H4'), enlace: svg('M10 14a5 5 0 007.07 0l3-3a5 5 0 00-7.07-7.07l-1.5 1.5M14 10a5 5 0 00-7.07 0l-3 3a5 5 0 007.07 7.07l1.5-1.5'), pdf: svg('M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2zM9 13h6m-6 4h6') };
  [['qBuscar', 'buscar'], ['rExport', 'bajar'], ['kExport', 'bajar'], ['kLink', 'enlace'], ['kPdf', 'pdf'], ['qExport', 'bajar'], ['rRefrescar', 'refrescar'], ['rNuevo', 'mas']].forEach(([id, k]) => $('#' + id).insertAdjacentHTML('afterbegin', IC[k]));
}
iniciar();
