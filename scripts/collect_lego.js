const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
const fs = require('fs');

const HEADERS = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' };
const BASE = 'https://brickranker.com';

const LEGO = {
  // 확정 (48/48)
  lego_falcon:    '/rankings/set/75192-1/millennium-falcon',
  lego_hogwarts:  '/rankings/set/71043-1/hogwarts-castle',
  // 결측 3개월 — 유지 후보
  lego_delorean:  '/rankings/set/10300-1/back-to-the-future',
  // 교체 후보 테스트 (2021년 이전 출시, 리셀 활발)
  lego_titanic:   '/rankings/set/10294-1/titanic',
  lego_colosseum: '/rankings/set/10276-1/colosseum',
  lego_razor:     '/rankings/set/75292-1/razor-crest',
  lego_shuttle:   '/rankings/set/10283-1/nasa-space-shuttle-discovery',
};

const sleep = ms => new Promise(r => setTimeout(r, ms));

function labelToYm(label) {
  // e.g. "Jan 22" -> "2022-01"
  const months = { Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12 };
  const m = label.trim().match(/^([A-Za-z]{3})\s*(\d{2})$/);
  if (!m) return null;
  const mon = months[m[1]];
  if (!mon) return null;
  const yr = 2000 + parseInt(m[2]);
  return `${yr}-${String(mon).padStart(2, '0')}`;
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

  const mlMatch = html.match(/labels\s*=\s*\[([^\]]+)\]/);
  const mpMatch = html.match(/(?:var\s+)?prices\s*=\s*\[([^\]]+)\]/);

  if (!mlMatch || !mpMatch) {
    console.log(`[${itemId}] labels/prices not found`);
    // Debug: show what variables are present
    const vars = html.match(/var\s+\w+\s*=/g) || [];
    console.log(`  variables found: ${vars.slice(0,10).join(', ')}`);
    return;
  }

  const labels = [...mlMatch[1].matchAll(/'([^']+)'/g)].map(m => m[1]);
  const prices = [...mpMatch[1].matchAll(/[\d.]+/g)].map(m => parseFloat(m[0]));

  if (labels.length === 0 || prices.length === 0) {
    console.log(`[${itemId}] empty labels or prices`); return;
  }

  const rows = [];
  for (let i = 0; i < Math.min(labels.length, prices.length); i++) {
    const ym = labelToYm(labels[i]);
    if (!ym || ym < '2022-01' || ym > '2025-12') continue;
    rows.push(`${ym},${prices[i].toFixed(2)},1`);
  }

  fs.mkdirSync('data/raw/prices', { recursive: true });
  fs.writeFileSync(`data/raw/prices/${itemId}.csv`,
    'year_month,mean_price,tx_count\n' + rows.join('\n') + '\n');
  console.log(`[${itemId}] ${rows.length}/48 months saved`);
}

async function main() {
  for (const [itemId, urlPath] of Object.entries(LEGO)) {
    await collect(itemId, urlPath);
    await sleep(2000);
  }
}

main().catch(console.error);
