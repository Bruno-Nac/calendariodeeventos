'use strict';

// Shared read-modify-write access to the single dashboard-data.json blob.
// Both the xlsx-upload path (api/process.js, full overwrite) and the manual
// CRUD path (api/turmas.js, single-turma read-modify-write) go through this
// so there's one place that knows how the blob is stored (private access,
// fixed pathname, seed fallback).

const { get, put } = require('@vercel/blob');
const seed = require('../data/dashboard-data.seed.json');

async function readDashboard() {
  try {
    const result = await get('dashboard-data.json', { access: 'private' });
    if (!result || !result.stream) throw new Error('blob not found');
    const text = await new Response(result.stream).text();
    return JSON.parse(text);
  } catch (e) {
    // Deep clone so callers can mutate freely without touching the shared seed object.
    return JSON.parse(JSON.stringify(seed));
  }
}

async function writeDashboard(data) {
  await put('dashboard-data.json', JSON.stringify(data), {
    access: 'private',
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: 'application/json',
  });
}

module.exports = { readDashboard, writeDashboard };
