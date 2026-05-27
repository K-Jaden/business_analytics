const fs = require('fs');

const ALL_MONTHS = [];
for (let y = 2022; y <= 2025; y++)
  for (let m = 1; m <= 12; m++)
    ALL_MONTHS.push(`${y}-${String(m).padStart(2,'0')}`);

const ITEMS = [
  'sneakers_jordan1','sneakers_panda','sneakers_yeezy','sneakers_travis','sneakers_nb550',
  'cards_charizard1','cards_umbreon','cards_rayquaza','cards_pikachu','cards_charizard2',
  'lego_atat','lego_taj','lego_homealone','lego_stranger','lego_haunted',
];

let passCount = 0;
for (const item of ITEMS) {
  const p = `data/raw/prices/${item}.csv`;
  if (!fs.existsSync(p)) { console.log(`${item}: FILE MISSING`); continue; }

  const lines = fs.readFileSync(p, 'utf8').trim().split('\n').slice(1);
  const present = new Set(lines.map(l => l.split(',')[0]));
  const prices = lines.map(l => parseFloat(l.split(',')[1]));
  const txs    = lines.map(l => parseInt(l.split(',')[2]));

  let maxGap = 0, gap = 0;
  for (const m of ALL_MONTHS) {
    if (!present.has(m)) { gap++; maxGap = Math.max(maxGap, gap); }
    else gap = 0;
  }

  const n = lines.length;
  const lowTx = txs.filter(t => t < 3).length;
  const priceMin = Math.min(...prices).toFixed(2);
  const priceMax = Math.max(...prices).toFixed(2);
  const ok = n >= 40 && maxGap < 3;
  if (ok) passCount++;
  const flag = ok ? 'PASS' : 'FAIL';
  console.log(`${flag}  ${item}: ${n}/48  max_gap=${maxGap}  tx_count<3: ${lowTx}개  $${priceMin}~$${priceMax}`);
}
console.log(`\n${passCount}/15 PASS`);
