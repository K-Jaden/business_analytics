// StockX salesGraph API 직접 호출 테스트
const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
const fs = require('fs');

const saved = JSON.parse(fs.readFileSync('stockx_cookies.json'));

// cf_clearance 쿠키 확인
const cfCookie = saved.cookies.find(c => c.name === 'cf_clearance');
console.log('cf_clearance expires:', cfCookie ? new Date(cfCookie.expiry * 1000).toISOString() : 'not found');

const cookieStr = saved.cookies.map(c => `${c.name}=${c.value}`).join('; ');

const HEADERS = {
  'Cookie': cookieStr,
  'Content-Type': 'application/json',
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36',
  'Accept': '*/*',
  'Origin': 'https://stockx.com',
  'Referer': 'https://stockx.com/air-jordan-1-retro-high-bordeaux',
  'x-stockx-device-id': saved.headers['x-stockx-device-id'],
  'x-stockx-session-id': saved.headers['x-stockx-session-id'],
  'apollographql-client-name': saved.headers['apollographql-client-name'],
};

// fetchSalesGraph GraphQL 쿼리 (StockX 공통 패턴)
const QUERY = `
query fetchSalesGraph($productId: String!, $countryCode: String!, $currencyCode: CurrencyCode, $startDate: String, $endDate: String, $intervals: Int) {
  product(id: $productId) {
    id
    title
    salesChart(countryCode: $countryCode, currencyCode: $currencyCode, startDate: $startDate, endDate: $endDate, intervals: $intervals) {
      series {
        xValue
        yValue
      }
    }
  }
}`;

async function testApi(slug) {
  const body = JSON.stringify({
    operationName: 'fetchSalesGraph',
    variables: {
      productId: slug,
      countryCode: 'KR',
      currencyCode: 'KRW',
      startDate: '2022-01-01',
      endDate: '2025-12-31',
      intervals: 100
    },
    query: QUERY
  });

  console.log(`\nTrying slug: ${slug}`);
  try {
    const res = await fetch('https://stockx.com/api/p/e', {
      method: 'POST', headers: HEADERS, body
    });
    console.log('Status:', res.status, res.statusText);
    const text = await res.text();
    console.log('Response (200 chars):', text.slice(0, 200));
    return res.status === 200 ? JSON.parse(text) : null;
  } catch (e) {
    console.log('Error:', e.message);
    return null;
  }
}

testApi('air-jordan-1-retro-high-bordeaux').catch(console.error);
