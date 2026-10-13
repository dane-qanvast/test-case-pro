// Bundles test-case-collaborator.jsx and inlines it into shell.html -> index.html
import { build } from 'esbuild';
import fs from 'fs';
const r = await build({
  entryPoints: ['test-case-collaborator.jsx'], bundle: true, minify: true, format: 'esm',
  jsx: 'automatic', write: false, legalComments: 'none',
  define: { 'process.env.NODE_ENV': '"production"' },
});
const js = r.outputFiles[0].text.replace(/<\/script/gi, '<\\/script');
const shell = fs.readFileSync('shell.html', 'utf8');
const html = shell.replace('/*__APP_BUNDLE__*/', () => js);
fs.writeFileSync('index.html', html); // served by GitHub Pages from the repo root
// Vercel looks for a "public" folder after building, so put a copy there too.
fs.mkdirSync('public', { recursive: true });
fs.writeFileSync('public/index.html', html);
console.log('Built index.html and public/index.html');
