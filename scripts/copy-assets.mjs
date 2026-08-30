// Copies the static renderer assets next to the compiled renderer bundle.
import { cp, mkdir } from 'node:fs/promises';
import path from 'node:path';

const from = path.resolve('src/renderer');
const to = path.resolve('dist/web/renderer');

await mkdir(to, { recursive: true });
for (const file of ['index.html', 'styles.css', 'paper-stack.svg']) {
  await cp(path.join(from, file), path.join(to, file));
}
console.log(`copied renderer assets -> ${to}`);
