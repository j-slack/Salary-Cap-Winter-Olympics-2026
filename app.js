// Loads OlympicTracker.xlsx, reads sheet "Display Points", sorts by Points (column B), and renders.

const EXCEL_FILE = 'OlympicTracker.xlsx';
const SHEET_NAME = 'Display Points';

const leaderBody = document.getElementById('leaderBody');
const updatedAt = document.getElementById('updatedAt');
const refreshBtn = document.getElementById('refreshBtn');
const medalsUpdatedAt = document.getElementById('medalsUpdatedAt');
const medalSourceEl = document.getElementById('medalSource');

function setUpdatedNow(){
  const now = new Date();
  updatedAt.textContent = now.toLocaleString();
}

function setMedalsUpdatedNow(){
  const now = new Date();
  if(medalsUpdatedAt) medalsUpdatedAt.textContent = now.toLocaleString();
}

function setMedalSource(label){
  if(medalSourceEl) medalSourceEl.textContent = label;
}

function rankDotClass(rank){
  if(rank === 1) return 'gold';
  if(rank === 2) return 'silver';
  if(rank === 3) return 'bronze';
  return '';
}

function safeNumber(v){
  if(v == null) return 0;
  if(typeof v === 'number') return v;
  const n = Number(String(v).trim());
  return Number.isFinite(n) ? n : 0;
}

function escapeHTML(str){
  return String(str)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

// Friendly display names (keeps your Excel values, but shows full names on the page)
const COUNTRY_DISPLAY_MAP = {
  'USA': 'United States',
  'US': 'United States',
  'U.S.': 'United States',
  'U.S.A.': 'United States',
  'Swiss': 'Switzerland',
  'Czech': 'Czech Republic',
  'UK': 'Great Britain',
  'AIN': 'Individual Neutral Athletes'
};

function normalizeCountryName(name){
  const raw = String(name ?? '').trim();
  if(!raw) return raw;
  const hitKey = Object.keys(COUNTRY_DISPLAY_MAP).find(k => k.toLowerCase() === raw.toLowerCase());
  return hitKey ? COUNTRY_DISPLAY_MAP[hitKey] : raw;
}

// Flag support (emoji flags from ISO-3166 alpha-2 codes)
const COUNTRY_FLAG_CODE = {
  'United States': 'US',
  'Canada': 'CA',
  'Great Britain': 'GB',
  'Germany': 'DE',
  'France': 'FR',
  'Italy': 'IT',
  'Japan': 'JP',
  'Netherlands': 'NL',
  'Switzerland': 'CH',
  'Austria': 'AT',
  'Belgium': 'BE',
  'Norway': 'NO',
  'Sweden': 'SE',
  'Finland': 'FI',
  'Poland': 'PL',
  'Slovakia': 'SK',
  'Slovenia': 'SI',
  'Estonia': 'EE',
  'Latvia': 'LV',
  'Lithuania': 'LT',
  'Ukraine': 'UA',
  'China': 'CN',
  'Spain': 'ES',
  'Hungary': 'HU',
  'Romania': 'RO',
  'Australia': 'AU',
  'New Zealand': 'NZ',
  'Jamaica': 'JM',
  'Argentina': 'AR',
  'South Korea': 'KR',
  'Czech Republic': 'CZ'
  // 'Individual Neutral Athletes' intentionally has no flag
};

function isoToFlagEmoji(iso2){
  const code = String(iso2 ?? '').toUpperCase();
  if(!/^[A-Z]{2}$/.test(code)) return '';
  const A = 0x1F1E6;
  const first = code.codePointAt(0) - 65 + A;
  const second = code.codePointAt(1) - 65 + A;
  return String.fromCodePoint(first, second);
}

function flagForCountry(countryName){
  const name = String(countryName ?? '').trim();
  if(!name) return '';
  const iso = COUNTRY_FLAG_CODE[name];
  return iso ? isoToFlagEmoji(iso) : '';
}

// --- Live medal table (Wikipedia via MediaWiki Action API) ---
// MediaWiki CORS requires an origin=* parameter for anonymous requests.
const WIKI_MEDALS_ENDPOINT = 'https://en.wikipedia.org/w/api.php?action=parse&page=2026_Winter_Olympics_medal_table&prop=text&format=json&origin=*';

// Name variants between different tables/sources.
const COUNTRY_MEDAL_ALIASES = {
  'Republic of Korea': 'South Korea',
  'Korea': 'South Korea',
  "People's Republic of China": 'China',
  'Czechia': 'Czech Republic'
};

function canonicalMedalCountryName(name){
  const n = String(name ?? '').trim();
  if(!n) return n;
  const hit = Object.keys(COUNTRY_MEDAL_ALIASES).find(k => k.toLowerCase() === n.toLowerCase());
  return hit ? COUNTRY_MEDAL_ALIASES[hit] : n;
}

function parseIntSafe(x){
  const n = Number(String(x ?? '').replace(/[^0-9]/g,'').trim());
  return Number.isFinite(n) ? n : 0;
}

async function fetchMedalTable(){
  // Returns: Map<countryName, {gold, silver, bronze, total}>
  const map = new Map();
  try{
    const res = await fetch(`${WIKI_MEDALS_ENDPOINT}&v=${Date.now()}`);
    if(!res.ok) throw new Error(`Medal fetch failed (${res.status})`);
    const data = await res.json();
    const html = data?.parse?.text?.['*'];
    if(!html) throw new Error('Medal HTML not found');

    const doc = new DOMParser().parseFromString(html, 'text/html');
    const tables = Array.from(doc.querySelectorAll('table.wikitable'));
    const medalTable = tables.find(t => {
      const headerText = t.querySelector('tr')?.innerText ?? '';
      return /Gold/i.test(headerText) && /Silver/i.test(headerText) && /Bronze/i.test(headerText);
    }) || tables[0];

    if(!medalTable) throw new Error('No medal table found');

    const rows = Array.from(medalTable.querySelectorAll('tr')).slice(1);
    for(const r of rows){
      const cells = r.querySelectorAll('th, td');
      if(cells.length < 5) continue;

      const countryCell = r.querySelector('th') || cells[1];
      let country = (countryCell?.innerText ?? '').trim();
      country = country.replace(/\[[^\]]*\]/g,'').trim();
      if(!country) continue;

      const td = Array.from(r.querySelectorAll('td'));
      if(td.length < 4) continue;

      const gold = parseIntSafe(td[0]?.innerText);
      const silver = parseIntSafe(td[1]?.innerText);
      const bronze = parseIntSafe(td[2]?.innerText);
      const total = parseIntSafe(td[3]?.innerText);

      map.set(canonicalMedalCountryName(country), {gold, silver, bronze, total});
    }

    setMedalSource('Wikipedia medal table');
    setMedalsUpdatedNow();
    return map;
  } catch (err){
    console.warn('Medal table unavailable:', err);
    setMedalSource('Medals unavailable');
    setMedalsUpdatedNow();
    return map;
  }
}

function parseTeamsCell(text){
  // Expected format: "Name (Country, Country, ...)".
  const s = (text ?? '').toString().trim();
  const open = s.indexOf('(');
  const close = s.lastIndexOf(')');

  if(open > 0 && close > open){
    const name = s.slice(0, open).trim();
    const inside = s.slice(open + 1, close);
    const countries = inside
      .split(',')
      .map(x => normalizeCountryName(x.trim()))
      .filter(Boolean)
      .sort((a,b) => a.localeCompare(b));
    return { name, countries, raw: s };
  }
  return { name: s, countries: [], raw: s };
}

function normalizeRow(row){
  const teamsCell = (row.Teams ?? row.teams ?? row.Team ?? row.Participant ?? '').toString();
  const points = safeNumber(row.Points ?? row.points ?? row.POINTS ?? row['Total Points'] ?? row['Points ']);

  const parsed = parseTeamsCell(teamsCell);
  return {
    participant: parsed.name,
    countries: parsed.countries,
    teamsRaw: parsed.raw,
    points
  };
}

function chipsHTML(items, limit = Infinity, medalMap = null){
  if(!items || !items.length) return '<span class="chip secondary">No picks listed</span>';
  const shown = items.slice(0, limit);
  const chips = shown.map(c => {
    const flag = flagForCountry(c);
    const flagSpan = flag ? `<span class="flag" aria-hidden="true">${flag}</span>` : '';

    const canonical = canonicalMedalCountryName(c);
    const m = medalMap?.get?.(canonical);
    const badge = (m && Number.isFinite(m.total))
      ? `<span class="medalBadge ${m.total===0?'zero':''}">${m.total}</span>`
      : `<span class="medalBadge unknown">—</span>`;

    const title = m
      ? `🥇${m.gold}  🥈${m.silver}  🥉${m.bronze}  • Total ${m.total}`
      : 'Medal count not available';

    return `<span class="chip" title="${escapeHTML(title)}">${flagSpan}${escapeHTML(c)}${badge}</span>`;
  }).join('');

  const extra = items.length > limit ? `<span class="chip secondary">+${items.length - limit} more</span>` : '';
  return chips + extra;
}

function renderPodium(sorted, medalMap){
  const p1 = sorted[0] || {participant:'—', countries:[], points:'—'};
  const p2 = sorted[1] || {participant:'—', countries:[], points:'—'};
  const p3 = sorted[2] || {participant:'—', countries:[], points:'—'};

  document.getElementById('p1name').textContent = p1.participant || '—';
  document.getElementById('p1pts').textContent = p1.points ?? '—';
  document.getElementById('p1picks').innerHTML = chipsHTML(p1.countries, 6, medalMap);

  document.getElementById('p2name').textContent = p2.participant || '—';
  document.getElementById('p2pts').textContent = p2.points ?? '—';
  document.getElementById('p2picks').innerHTML = chipsHTML(p2.countries, 6, medalMap);

  document.getElementById('p3name').textContent = p3.participant || '—';
  document.getElementById('p3pts').textContent = p3.points ?? '—';
  document.getElementById('p3picks').innerHTML = chipsHTML(p3.countries, 6, medalMap);
}

function renderTable(sorted, medalMap){
  leaderBody.innerHTML = '';

  sorted.forEach((row, idx) => {
    const rank = idx + 1;
    const tr = document.createElement('tr');

    const tdRank = document.createElement('td');
    tdRank.innerHTML = `
      <span class="rankBadge">
        <span class="rankDot ${rankDotClass(rank)}"></span>
        #${rank}
      </span>
    `;

    const tdTeam = document.createElement('td');
    if(row.countries?.length){
      tdTeam.innerHTML = `
        <div class="who">${escapeHTML(row.participant)}</div>
        <div class="picksLabel">Selected countries (A→Z):</div>
        <div class="chips">${chipsHTML(row.countries, Infinity, medalMap)}</div>
      `;
    } else {
      tdTeam.innerHTML = `<div class="who">${escapeHTML(row.teamsRaw)}</div>`;
    }

    const tdPoints = document.createElement('td');
    tdPoints.className = 'points';
    tdPoints.textContent = row.points;

    tr.appendChild(tdRank);
    tr.appendChild(tdTeam);
    tr.appendChild(tdPoints);
    leaderBody.appendChild(tr);
  });
}

async function loadExcelAndRender(){
  leaderBody.innerHTML = '<tr><td colspan="3" class="loading">Loading Excel…</td></tr>';

  try {
    // Cache-bust to pick up updates quickly
    const url = `${EXCEL_FILE}?v=${Date.now()}`;

    // Fetch Excel and medal table in parallel
    const [excelRes, medalMap] = await Promise.all([
      fetch(url),
      fetchMedalTable()
    ]);

    if(!excelRes.ok) throw new Error(`Could not fetch ${EXCEL_FILE} (${excelRes.status})`);
    const arrayBuffer = await excelRes.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheet = workbook.Sheets[SHEET_NAME];
    if(!sheet) {
      const available = workbook.SheetNames.join(', ');
      throw new Error(`Sheet "${SHEET_NAME}" not found. Available: ${available}`);
    }

    const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const rows = raw
      .map(normalizeRow)
      .filter(r => (r.participant || '').length);

    // Sort largest → smallest by points
    rows.sort((a,b) => b.points - a.points);

    renderPodium(rows, medalMap);
    renderTable(rows, medalMap);
    setUpdatedNow();
  } catch(err){
    console.error(err);
    leaderBody.innerHTML = `
      <tr>
        <td colspan="3" class="loading">
          ⚠️ ${escapeHTML(err.message)}<br/>
          Make sure <code>${EXCEL_FILE}</code> is beside <code>index.html</code> and has a sheet named <code>${SHEET_NAME}</code>.
        </td>
      </tr>
    `;
    setUpdatedNow();
  }
}

refreshBtn?.addEventListener('click', loadExcelAndRender);

// Initial load
loadExcelAndRender();

// Auto-refresh every 60 seconds
setInterval(loadExcelAndRender, 60_000);
