import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join, dirname } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

export const readFixture = (name) => readFileSync(join(root, 'tests', 'fixtures', name), 'utf8');
export const readSample = (name) => readFileSync(join(root, 'samples', name), 'utf8');
export const sampleNames = () =>
  readdirSync(join(root, 'samples')).filter((f) => f.endsWith('.txt')).sort();
