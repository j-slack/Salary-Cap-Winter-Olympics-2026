// Loads OlympicTracker.xlsx, reads sheet "Display Points", sorts by column B (Points), and renders.

const EXCEL_FILE = 'OlympicTracker.xlsx';
const SHEET_NAME = 'Display Points';

const leaderBody = document.getElementById('leaderBody');
const updatedAt = document.getElementById('updatedAt');
const refreshBtn = document.getElementById('refreshBtn');

function setUpdatedNow(){
  const now = new Date();
  updatedAt.textContent = now.toLocaleString();
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

function normalizeRow(row){
  // Expecting: {Teams: '...', Points: 18}
  const teams = (row.Teams ?? row.teams ?? row.Team ?? row.Participant ?? '').toString().trim();
  const points = safeNumber(row.Points ?? row.points ?? row.POINTS ?? row['Total Points'] ?? row['Points ']);
  return { teams, points };
}

function renderPodium(sorted){
  const p1 = sorted[0] || {teams:'—', points:'—'};
  const p2 = sorted[1] || {teams:'—', points:'—'};
  const p3 = sorted[2] || {teams:'—', points:'—'};

  document.getElementById('p1name').textContent = p1.teams || '—';
  document.getElementById('p1pts').textContent = p1.points ?? '—';

  document.getElementById('p2name').textContent = p2.teams || '—';
  document.getElementById('p2pts').textContent = p2.points ?? '—';

  document.getElementById('p3name').textContent = p3.teams || '—';
  document.getElementById('p3pts').textContent = p3.points ?? '—';
}

function renderTable(sorted){
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
    tdTeam.textContent = row.teams;

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
    // Bust caches so updates to the xlsx propagate quickly (important on static hosts)
    const url = `${EXCEL_FILE}?v=${Date.now()}`;
    const res = await fetch(url);
    if(!res.ok) throw new Error(`Could not fetch ${EXCEL_FILE} (${res.status})`);
    const arrayBuffer = await res.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheet = workbook.Sheets[SHEET_NAME];
    if(!sheet) {
      const available = workbook.SheetNames.join(', ');
      throw new Error(`Sheet "${SHEET_NAME}" not found. Available: ${available}`);
    }

    const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const rows = raw
      .map(normalizeRow)
      .filter(r => r.teams.length);

    // Sort largest → smallest by points (column B)
    rows.sort((a,b) => b.points - a.points);

    renderPodium(rows);
    renderTable(rows);
    setUpdatedNow();
  } catch(err){
    console.error(err);
    leaderBody.innerHTML = `
      <tr>
        <td colspan="3" class="loading">
          ⚠️ ${err.message}<br/>
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

// Optional: auto-refresh every 60 seconds
setInterval(loadExcelAndRender, 60_000);
