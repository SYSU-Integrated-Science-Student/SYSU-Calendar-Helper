#!/usr/bin/env node

import { mkdir, readdir, rm, stat, copyFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.join(__dirname, '..');
const sourceDir = path.join(projectRoot, 'frontend');
const targetDir = path.join(projectRoot, 'dist');

// Shallow Vite-style copy without transformations to keep deployment simple.
async function copyDirectory(src, dest) {
  await mkdir(dest, { recursive: true });
  const entries = await readdir(src, { withFileTypes: true });

  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);

    if (entry.isDirectory()) {
      await copyDirectory(srcPath, destPath);
    } else if (entry.isFile()) {
      await copyFile(srcPath, destPath);
    }
  }
}

async function ensureSourceExists() {
  try {
    const stats = await stat(sourceDir);
    if (!stats.isDirectory()) {
      throw new Error(`Expected ${sourceDir} to be a directory.`);
    }
  } catch (error) {
    throw new Error(`Unable to locate frontend source directory at ${sourceDir}.`);
  }
}

async function main() {
  await ensureSourceExists();
  await rm(targetDir, { recursive: true, force: true });
  await copyDirectory(sourceDir, targetDir);
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
