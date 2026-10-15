'use strict';

// Serves the current processed dashboard data to the front-end. The page URL
// never changes and the front-end never needs to know the blob's actual URL.
// See lib/turmaStore.js for the read path (private Blob, seed fallback).

const { readDashboard } = require('../lib/turmaStore');

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  const data = await readDashboard();
  res.status(200).send(JSON.stringify(data));
};
