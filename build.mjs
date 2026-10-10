// Genera la carpeta dist con los archivos de la aplicación; el código y los estilos salen reducidos.
// Vercel lo ejecuta en cada publicación. Los archivos .sql y las maquetas no se publican.
import { transform } from 'esbuild';
import fs from 'node:fs';

fs.rmSync('dist', { recursive: true, force: true });
fs.mkdirSync('dist', { recursive: true });

for (const [f, loader] of [['app.js', 'js'], ['ocr.js', 'js'], ['estilos.css', 'css']]) {
  const src = fs.readFileSync(f, 'utf8');
  const r = await transform(src, { loader, minify: true, target: 'es2020', legalComments: 'none' });
  fs.writeFileSync('dist/' + f, r.code);
  console.log(f, (src.length / 1024).toFixed(0) + ' KB ->', (r.code.length / 1024).toFixed(0) + ' KB');
}
for (const f of ['index.html', 'config.js', 'logo.png', 'camion.webp', 'camion_hoja2.webp', 'costos_arte2.webp', 'mapa_peru.json', 'plantilla_liquidacion.xlsx', 'plantilla_despacho.xlsx']) {
  if (fs.existsSync(f)) fs.copyFileSync(f, 'dist/' + f); else console.warn('no existe, se omite:', f);
}
