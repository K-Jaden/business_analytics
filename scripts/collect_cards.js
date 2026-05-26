const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
const cheerio = require('cheerio');
const fs = require('fs');
const path = require('path');

const HEADERS = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' };
const BASE = 'https://www.pricecharting.com';

const CARDS = {
  cards_charizard1: '/game/pokemon-shining-fates/charizard-vmax-sv107',
  cards_umbreon:    '/game/pokemon-evolving-skies/umbreon-vmax-215',
  cards_rayquaza:   '/game/pokemon-evolving-skies/rayquaza-vmax-218',
  cards_pikachu:    '/game/pokemon-vivid-voltage/pikachu-vmax-188',
  cards_charizard2: '/game/pokemon-hidden-fates/charizard-gx-sv49',
};

const sleep = ms => new Promise(r => setTimeout(r, ms));

function parseChartData(html) {
  // Extract VGPC.chart_data = {...};
  const m = html.match(/VGPC\.chart_data\s*=\s*(\{[\s\S]*?\});\s*\n/);
  if (!m) return null;
  try { return JSON.parse(m[1]); } catch { return null; }
}

async function collect(itemId, urlPath) {
  console.log(`[${itemId}] fetching...`);
  let res;
  try {
    res = await fetch(BASE + urlPath, { headers: HEADERS });
  } catch (e) {
    console.log(`[${itemId}] fetch error: ${e.message}`); return;
  }
  if (res.status !== 200) {
    console.log(`[${itemId}] HTTP ${res.status}`); return;
  }
  const html = await res.text();

  // Check redirect
  const finalUrl = res.url;
  if (!finalUrl.includes(urlPath.split('/').pop())) {
    console.log(`[${itemId}] WARNING: redirected to ${finalUrl}`);
  }

  const data = parseChartData(html);
  if (!data) {
    console.log(`[${itemId}] chart_data not found`); return;
  }

  // graded = PSA 10 price series [[timestamp_ms, price_cents], ...]
  const graded = (data.graded || []).filter(([, v]) => v > 0);
  if (graded.length === 0) {
    console.log(`[${itemId}] no graded data`); return;
  }

  // Group by year_month, average price per month
  const monthly = {};
  for (const [ts, priceCents] of graded) {
    const d = new Date(ts);
    const ym = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
    if (ym < '2022-01' || ym > '2025-12') continue;
    if (!monthly[ym]) monthly[ym] = [];
    monthly[ym].push(priceCents / 100);
  }

  const rows = Object.keys(monthly).sort().map(ym => {
    const prices = monthly[ym];
    const mean = prices.reduce((a, b) => a + b, 0) / prices.length;
    return `${ym},${mean.toFixed(2)},${prices.length}`;
  });

  fs.mkdirSync('data/raw/prices', { recursive: true });
  fs.writeFileSync(`data/raw/prices/${itemId}.csv`,
    'year_month,mean_price,tx_count\n' + rows.join('\n') + '\n');
  console.log(`[${itemId}] ${rows.length}/48 months saved`);
}

async function main() {
  for (const [itemId, urlPath] of Object.entries(CARDS)) {
    await collect(itemId, urlPath);
    await sleep(3000);
  }
}

main().catch(console.error);
