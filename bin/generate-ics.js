#!/usr/bin/env node
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseScheduleFromPackage, buildIcs } from '../src/schedule.js';

function parseArgs(argv) {
  const args = {
    input: 'samples/sample-package.xml',
    startDate: '2025-09-08',
    timeZone: 'Asia/Shanghai',
    format: 'ics'
  };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    switch (arg) {
      case '--input':
      case '-i':
        args.input = argv[++i];
        break;
      case '--start':
        args.startDate = argv[++i];
        break;
      case '--tz':
        args.timeZone = argv[++i];
        break;
      case '--format':
        args.format = argv[++i];
        break;
      case '--help':
      case '-h':
        args.help = true;
        break;
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }
  return args;
}

function printHelp() {
  console.log(`Usage: generate-ics [options]\n\n` +
    `Options:\n` +
    `  -i, --input <file>   Flat OPC XML package path (default: samples/sample-package.xml)\n` +
    `      --start <date>   First-week Monday in YYYY-MM-DD (default: 2025-09-08)\n` +
    `      --tz <zone>      IANA timezone identifier (default: Asia/Shanghai)\n` +
    `      --format <fmt>   Output format: ics|json (default: ics)\n` +
    `  -h, --help           Show this message`);
}

function loadPackage(inputPath) {
  const resolved = resolve(process.cwd(), inputPath);
  return readFileSync(resolved);
}

async function main() {
  const argv = process.argv.slice(2);
  let options;
  try {
    options = parseArgs(argv);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }

  if (options.help) {
    printHelp();
    return;
  }

  try {
    const pkgData = loadPackage(options.input);
    const { events, timeZone } = await parseScheduleFromPackage(pkgData, {
      startDate: options.startDate,
      timeZone: options.timeZone
    });

    if (options.format === 'json') {
      console.log(JSON.stringify(events, null, 2));
      return;
    }

    if (options.format !== 'ics') {
      throw new Error(`Unsupported output format: ${options.format}`);
    }

    const ics = buildIcs(events, timeZone);
    console.log(ics);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
