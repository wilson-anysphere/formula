import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import {
  ArrowColumnarSheet,
  arrowTableFromColumns,
  arrowTableFromIPC,
  arrowTableToGridBatches,
  arrowTableToIPC,
  arrowTableToParquet,
  parquetFileToArrowTable,
  parquetToArrowTable,
} from '../src/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test('Parquet import -> Arrow table -> grid batches', async () => {
  const parquetPath = path.join(__dirname, 'fixtures', 'simple.parquet');
  const parquetBytes = new Uint8Array(await readFile(parquetPath));

  const table = await parquetToArrowTable(parquetBytes, { batchSize: 2 });

  assert.equal(table.numRows, 3);
  assert.deepEqual(
    table.schema.fields.map((f) => f.name),
    ['id', 'name', 'active', 'score']
  );

  assert.equal(table.getChildAt(0).get(0), 1);
  assert.equal(table.getChildAt(1).get(2), 'Carla');
  assert.equal(table.getChildAt(2).get(1), false);

  const grid = [];
  for await (const batch of arrowTableToGridBatches(table, {
    batchSize: 2,
    includeHeader: true,
  })) {
    for (let i = 0; i < batch.values.length; i++) {
      grid[batch.rowOffset + i] = batch.values[i];
    }
  }

  assert.deepEqual(grid[0], ['id', 'name', 'active', 'score']);
  assert.deepEqual(grid[1], [1, 'Alice', true, 1.5]);
  assert.deepEqual(grid[2], [2, 'Bob', false, 2.25]);
  assert.deepEqual(grid[3], [3, 'Carla', true, 3.75]);
});

test('arrowTableToGridBatches can omit header row', async () => {
  const parquetPath = path.join(__dirname, 'fixtures', 'simple.parquet');
  const parquetBytes = new Uint8Array(await readFile(parquetPath));

  const table = await parquetToArrowTable(parquetBytes, { batchSize: 2 });

  const grid = [];
  for await (const batch of arrowTableToGridBatches(table, {
    batchSize: 2,
    includeHeader: false,
  })) {
    for (let i = 0; i < batch.values.length; i++) {
      grid[batch.rowOffset + i] = batch.values[i];
    }
  }

  assert.deepEqual(grid[0], [1, 'Alice', true, 1.5]);
  assert.deepEqual(grid[2], [3, 'Carla', true, 3.75]);
});

test('Parquet import from Blob streams and can emit grid batches', async () => {
  const parquetPath = path.join(__dirname, 'fixtures', 'simple.parquet');
  const parquetBytes = new Uint8Array(await readFile(parquetPath));

  const blob = new Blob([parquetBytes]);

  const grid = [];
  const table = await parquetFileToArrowTable(blob, {
    batchSize: 2,
    gridBatchSize: 2,
    includeHeader: true,
    onGridBatch: async (batch) => {
      for (let i = 0; i < batch.values.length; i++) {
        grid[batch.rowOffset + i] = batch.values[i];
      }
    },
  });

  assert.equal(table.numRows, 3);
  assert.deepEqual(grid[0], ['id', 'name', 'active', 'score']);
  assert.deepEqual(grid[1], [1, 'Alice', true, 1.5]);
  assert.deepEqual(grid[3], [3, 'Carla', true, 3.75]);
});

test('Parquet export produces a readable parquet file', async () => {
  const parquetPath = path.join(__dirname, 'fixtures', 'simple.parquet');
  const parquetBytes = new Uint8Array(await readFile(parquetPath));

  const table = await parquetToArrowTable(parquetBytes, { batchSize: 2 });

  const exportedBytes = await arrowTableToParquet(table, {
    compression: 'UNCOMPRESSED',
  });

  const magic = String.fromCharCode(...exportedBytes.slice(0, 4));
  const magicTail = String.fromCharCode(...exportedBytes.slice(-4));
  assert.equal(magic, 'PAR1');
  assert.equal(magicTail, 'PAR1');

  const tableRoundTrip = await parquetToArrowTable(exportedBytes, {
    batchSize: 2,
  });

  assert.equal(tableRoundTrip.numRows, 3);
  assert.equal(tableRoundTrip.getChildAt(1).get(0), 'Alice');
  assert.equal(tableRoundTrip.getChildAt(3).get(2), 3.75);
});

test('ArrowColumnarSheet provides a columnar backing store', async () => {
  const parquetPath = path.join(__dirname, 'fixtures', 'simple.parquet');
  const parquetBytes = new Uint8Array(await readFile(parquetPath));

  const table = await parquetToArrowTable(parquetBytes, { batchSize: 2 });
  const sheet = new ArrowColumnarSheet(table);

  assert.equal(sheet.rowCount, 4);
  assert.equal(sheet.columnCount, 4);
  assert.equal(sheet.getCell(0, 0), 'id');
  assert.equal(sheet.getCell(3, 1), 'Carla');
});

test('Arrow IPC roundtrip preserves schema and values', () => {
  const table = arrowTableFromColumns({
    id: [1, 2, 3],
    name: ['Alice', 'Bob', 'Carla'],
    occurredAt: [new Date('2024-01-01T00:00:00.000Z'), null, new Date('2024-01-03T12:34:56.000Z')],
  });

  const bytes = arrowTableToIPC(table);
  const roundTrip = arrowTableFromIPC(bytes);

  assert.equal(roundTrip.numRows, table.numRows);
  assert.deepEqual(
    roundTrip.schema.fields.map((f) => f.name),
    table.schema.fields.map((f) => f.name),
  );
  assert.equal(roundTrip.getChildAt(0).get(0), 1);
  assert.equal(roundTrip.getChildAt(1).get(2), 'Carla');
});
