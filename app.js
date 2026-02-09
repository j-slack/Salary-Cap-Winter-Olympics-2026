// v11: Edge hover fix: position tooltips using the pointer (clientX) on hover
const EXCEL_FILE = 'OlympicTracker.xlsx';
const SHEET_POINTS = 'Display Points';
const SHEET_MEDALS = 'Medal Count';

const leaderBody = document.getElementById('leaderBody');
const updatedAt = document.getElementById('updatedAt');
const refreshBtn = document.getElementById('refreshBtn');
const medalsUpdatedAt = document.getElementById('medalsUpdatedAt');
const medalSourceEl = document.getElementById('medalSource');

function setUpdatedNow(){ updatedAt.textContent = new Date().toLocaleString(); }
function setMedalsUpdatedNow(){ if(medalsUpdatedAt) medalsUpdatedAt.textContent = new Date().toLocaleString(); }
function setMedalSource(label){ if(medalSourceEl) medalSourceEl.textContent = label; }

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

const COUNTRY_DISPLAY_MAP = { 'USA':'United States','US':'United States','U.S.':'United States','U.S.A.':'United States','Swiss':'Switzerland','Czech':'Czech Republic','UK':'Great Britain','AIN':'Individual Neutral Athletes' };
function normalizeCountryName(name){
  const raw = String(name ?? '').trim();
  if(!raw) return raw;
  const hitKey = Object.keys(COUNTRY_DISPLAY_MAP).find(k => k.toLowerCase() === raw.toLowerCase());
  return hitKey ? COUNTRY_DISPLAY_MAP[hitKey] : raw;
}

const COUNTRY_FLAG_CODE = {
  'United States':'US','Canada':'CA','Great Britain':'GB','Germany':'DE','France':'FR','Italy':'IT','Japan':'JP','Netherlands':'NL','Switzerland':'CH','Austria':'AT','Belgium':'BE','Norway':'NO','Sweden':'SE','Finland':'FI','Poland':'PL','Slovakia':'SK','Slovenia':'SI','Estonia':'EE','Latvia':'LV','Lithuania':'LT','Ukraine':'UA','China':'CN','Spain':'ES','Hungary':'HU','Romania':'RO','Australia':'AU','New Zealand':'NZ','Jamaica':'JM','Argentina':'AR','South Korea':'KR','Czech Republic':'CZ'
};
function isoToFlagEmoji(iso2){
  const code = String(iso2 ?? '').toUpperCase();
  if(!/^[A-Z]{2}$/.test(code)) return '';
  const A = 0x1F1E6;
  return String.fromCodePoint(code.codePointAt(0) - 65 + A, code.codePointAt(1) - 65 + A);
}
function flagForCountry(countryName){
  const iso = COUNTRY_FLAG_CODE[String(countryName ?? '').trim()];
  return iso ? isoToFlagEmoji(iso) : '';
}

const COUNTRY_MEDAL_ALIASES = {"People's Republic of China":'China','Czechia':'Czech Republic','Republic of Korea':'South Korea'};
function canonicalMedalCountryName(name){
  const n = normalizeCountryName(name);
  const hit = Object.keys(COUNTRY_MEDAL_ALIASES).find(k => k.toLowerCase() === String(n).toLowerCase());
  return hit ? COUNTRY_MEDAL_ALIASES[hit] : n;
}

function parseTeamsCell(text){
  const s = (text ?? '').toString().trim();
  const open = s.indexOf('(');
  const close = s.lastIndexOf(')');
  if(open > 0 && close > open){
    const name = s.slice(0, open).trim();
    const inside = s.slice(open + 1, close);
    const countries = inside.split(',').map(x => normalizeCountryName(x.trim())).filter(Boolean).sort((a,b) => a.localeCompare(b));
    return { name, countries, raw: s };
  }
  return { name: s, countries: [], raw: s };
}

function normalizeRow(row){
  const teamsCell = (row.Teams ?? row.teams ?? row.Team ?? row.Participant ?? '').toString();
  const points = safeNumber(row.Points ?? row.points ?? row.POINTS ?? row['Total Points'] ?? row['Points ']);
  const parsed = parseTeamsCell(teamsCell);
  return { participant: parsed.name, countries: parsed.countries, teamsRaw: parsed.raw, points };
}

function buildMedalMap(workbook){
  const map = new Map();
  const sheet = workbook.Sheets[SHEET_MEDALS];
  if(!sheet){ setMedalSource('Missing Medal Count sheet'); setMedalsUpdatedNow(); return map; }
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  raw.forEach(r => {
    const country = canonicalMedalCountryName(r.Country ?? r.country ?? r['Country'] ?? r.NOC);
    if(!country) return;
    const gold = safeNumber(r['Gold Medals'] ?? r.Gold ?? r['Gold']);
    const silver = safeNumber(r['Silver Medals'] ?? r.Silver ?? r['Silver']);
    const bronze = safeNumber(r['Bronze Medals'] ?? r.Bronze ?? r['Bronze']);
    const total = safeNumber(r['Total Medals'] ?? r.Total ?? r['Total']);
    map.set(country, {gold, silver, bronze, total});
  });
  setMedalSource('Excel (Medal Count)'); setMedalsUpdatedNow();
  return map;
}

function chipsHTML(items, limit = Infinity, medalMap = null){
  if(!items || !items.length) return '<span class="chip secondary">No picks listed</span>';
  const shown = items.slice(0, limit);
  return shown.map(c => {
    const flag = flagForCountry(c);
    const flagSpan = flag ? `<span class="flag" aria-hidden="true">${flag}</span>` : '';
    const m = medalMap?.get?.(canonicalMedalCountryName(c));
    const badge = (m && Number.isFinite(m.total)) ? `<span class="medalBadge">${m.total}</span>` : `<span class="medalBadge">—</span>`;
    const tooltip = m
      ? `<div class="tooltip" role="tooltip"><div><strong>${escapeHTML(c)}</strong></div><div class="row"><span>🥇</span><span>${m.gold}</span></div><div class="row"><span>🥈</span><span>${m.silver}</span></div><div class="row"><span>🥉</span><span>${m.bronze}</span></div><div class="row"><span>Total</span><span>${m.total}</span></div></div>`
      : `<div class="tooltip" role="tooltip"><div><strong>${escapeHTML(c)}</strong></div><div class="row"><span>Medals</span><span>—</span></div></div>`;
    return `<span class="chip" tabindex="0" aria-label="${escapeHTML(c)} medal details">${flagSpan}${escapeHTML(c)}${badge}${tooltip}</span>`;
  }).join('') + (items.length > limit ? `<span class="chip secondary">+${items.length - limit} more</span>` : '');
}

function positionTooltip(chip, evt=null){
  const tip = chip.querySelector('.tooltip');
  if(!tip) return;
  const rect = chip.getBoundingClientRect();

  // Force open briefly for measurement
  const wasOpen = chip.classList.contains('open');
  chip.classList.add('open');
  const tipRect = tip.getBoundingClientRect();
  if(!wasOpen) chip.classList.remove('open');

  const margin = 10;
  const anchorX = evt && typeof evt.clientX === 'number' ? evt.clientX : (rect.left + rect.width/2);
  let x = anchorX - (tipRect.width/2);
  let y = rect.bottom + 8;

  const maxX = window.innerWidth - tipRect.width - margin;
  x = Math.max(margin, Math.min(x, maxX));
  if(y + tipRect.height + margin > window.innerHeight){ y = rect.top - tipRect.height - 8; }
  y = Math.max(margin, Math.min(y, window.innerHeight - tipRect.height - margin));

  tip.style.setProperty('--tip-x', `${Math.round(x)}px`);
  tip.style.setProperty('--tip-y', `${Math.round(y)}px`);
}

function wireChipTooltips(scope=document){
  const chips = scope.querySelectorAll('.chip:not(.secondary)');
  const hoverFine = window.matchMedia('(hover: hover) and (pointer: fine)').matches;

  chips.forEach(chip => {
    if(chip.dataset.bound === '1') return;
    chip.dataset.bound = '1';

    const toggle = (e) => {
      e.stopPropagation();
      const isOpen = chip.classList.contains('open');
      document.querySelectorAll('.chip.open').forEach(c => c.classList.remove('open'));
      if(!isOpen){ chip.classList.add('open'); positionTooltip(chip, e); }
    };

    chip.addEventListener('click', toggle);

    // Desktop hover: open + follow pointer (fixes Edge podium #1 drift)
    chip.addEventListener('mouseenter', (e) => {
      if(!hoverFine) return;
      chip.classList.add('open');
      positionTooltip(chip, e);
    });
    chip.addEventListener('mousemove', (e) => {
      if(!hoverFine) return;
      if(chip.classList.contains('open')) positionTooltip(chip, e);
    });
    chip.addEventListener('mouseleave', () => { if(hoverFine) chip.classList.remove('open'); });

    chip.addEventListener('keydown', (e) => {
      if(e.key === 'Enter' || e.key === ' ') toggle(e);
      if(e.key === 'Escape') chip.classList.remove('open');
    });
  });

  if(!document.body.dataset.tooltipCloser){
    document.body.dataset.tooltipCloser = '1';
    document.addEventListener('click', () => document.querySelectorAll('.chip.open').forEach(c => c.classList.remove('open')));
    document.addEventListener('scroll', () => document.querySelectorAll('.chip.open').forEach(c => c.classList.remove('open')), {passive:true});
    window.addEventListener('resize', () => document.querySelectorAll('.chip.open').forEach(c => positionTooltip(c)));
  }
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
    tdRank.innerHTML = `<span class="rankBadge">#${rank}</span>`;

    const tdTeam = document.createElement('td');
    if(row.countries?.length){
      tdTeam.innerHTML = `<div class="who">${escapeHTML(row.participant)}</div><div class="picksLabel">Selected countries (A→Z):</div><div class="chips">${chipsHTML(row.countries, Infinity, medalMap)}</div>`;
    } else {
      tdTeam.innerHTML = `<div class="who">${escapeHTML(row.teamsRaw)}</div>`;
    }

    const tdPoints = document.createElement('td');
    tdPoints.className = 'points';
    tdPoints.textContent = row.points;

    tr.appendChild(tdRank); tr.appendChild(tdTeam); tr.appendChild(tdPoints);
    leaderBody.appendChild(tr);
  });
}

async function loadExcelAndRender(){
  leaderBody.innerHTML = '<tr><td colspan="3" class="loading">Loading Excel…</td></tr>';
  try{
    const url = `${EXCEL_FILE}?v=${Date.now()}`;
    const res = await fetch(url);
    if(!res.ok) throw new Error(`Could not fetch ${EXCEL_FILE} (${res.status})`);
    const buf = await res.arrayBuffer();
    const workbook = XLSX.read(buf, { type: 'array' });

    const pointsSheet = workbook.Sheets[SHEET_POINTS];
    if(!pointsSheet) throw new Error(`Sheet "${SHEET_POINTS}" not found. Available: ${workbook.SheetNames.join(', ')}`);

    const medalMap = buildMedalMap(workbook);
    const raw = XLSX.utils.sheet_to_json(pointsSheet, { defval: '' });
    const rows = raw.map(normalizeRow).filter(r => (r.participant || '').length);
    rows.sort((a,b) => b.points - a.points);

    renderPodium(rows, medalMap);
    renderTable(rows, medalMap);
    wireChipTooltips(document);
    setUpdatedNow();
  } catch(err){
    console.error(err);
    leaderBody.innerHTML = `<tr><td colspan="3" class="loading">⚠️ ${escapeHTML(err.message)}</td></tr>`;
    setUpdatedNow();
    setMedalSource('Error');
    setMedalsUpdatedNow();
  }
}

refreshBtn?.addEventListener('click', loadExcelAndRender);
loadExcelAndRender();
setInterval(loadExcelAndRender, 60_000);
