import { build } from 'esbuild';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

const outfile = fileURLToPath(new URL('../dist/js/lib/sei-pro-exportar-doc.js', import.meta.url));

await build({
  entryPoints: [fileURLToPath(new URL('src/main.ts', import.meta.url))],
  outfile,
  bundle: true,
  minify: true,
  format: 'iife',
  target: ['chrome116', 'firefox115'],
  charset: 'ascii',
  legalComments: 'none',
  banner: { js: '/* GERADO por exportar-documento/build.mjs. NAO EDITE. Rode: npm run build */' },
});

if ((await readFile(outfile)).some((byte) => byte > 127)) {
  throw new Error('Bundle contem bytes nao ASCII');
}
console.log('Exportar documento: bundle gerado.');
