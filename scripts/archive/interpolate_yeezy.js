const fs = require('fs');

const path = 'data/raw/prices/sneakers_yeezy.csv';
const lines = fs.readFileSync(path, 'utf8').trim().split('\n');
const header = lines[0];
const rows = lines.slice(1).map(l => {
  const [ym, price, tx] = l.split(',');
  return { ym, price: parseFloat(price), tx: parseInt(tx) };
});

// 월별 맵 구성
const map = {};
for (const r of rows) map[r.ym] = r;

const allMonths = [];
for (let y = 2022; y <= 2025; y++)
  for (let m = 1; m <= 12; m++)
    allMonths.push(`${y}-${String(m).padStart(2,'0')}`);

// 결측월 4개: 2022-10, 2023-09, 2024-08, 2025-06
// 각각 앞뒤 월 사이 선형보간
const missing = allMonths.filter(m => !map[m]);
console.log('결측월:', missing);

for (const ym of missing) {
  const idx = allMonths.indexOf(ym);
  const prev = allMonths[idx - 1];
  const next = allMonths[idx + 1];
  if (!map[prev] || !map[next]) {
    console.log(`  ${ym}: 앞뒤 데이터 없음 — 보간 불가`);
    continue;
  }
  const interpolated = (map[prev].price + map[next].price) / 2;
  map[ym] = { ym, price: Math.round(interpolated * 100) / 100, tx: 0 };
  console.log(`  ${ym}: ${map[prev].price} ~ ${map[next].price} → ${map[ym].price} (보간, tx_count=0)`);
}

// 정렬 후 저장
const output = [header];
for (const m of allMonths) {
  if (map[m]) output.push(`${m},${map[m].price},${map[m].tx}`);
}
fs.writeFileSync(path, output.join('\n') + '\n');
console.log(`\n저장 완료: ${output.length - 1}/48 months`);
