// Builds index.html (the whole app in one file) from test-case-collaborator.jsx + shell.html.
// Run: npm run build
import { build } from 'esbuild';
import fs from 'fs';
const r = await build({
  entryPoints: ['test-case-collaborator.jsx'], bundle: true, minify: true, format: 'iife',
  jsx: 'automatic', write: false, legalComments: 'none',
  // Icons: every lucide-react import is drawn with Google's Material Symbols instead.
  alias: { 'lucide-react': './lucide-shim.jsx' },
  define: { 'process.env.NODE_ENV': '"production"' },
});
const js = r.outputFiles[0].text.replace(/<\/script/gi, '<\\/script');
const shell = fs.readFileSync('shell.html', 'utf8');
fs.writeFileSync('index.html', shell.replace('/*__APP_BUNDLE__*/', () => js));
console.log('Built index.html');
