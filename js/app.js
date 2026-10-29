/* APAAR Creation Monitor - SSM, Purba Bardhaman
 * Monitors APAAR ID generation by block, school, management and category, from the
 * "APAAR Entry Status - School Wise (Only Operational)" export of the UDISE+ portal.
 * Static app (GitHub Pages). Daily snapshots are stored in IndexedDB and can
 * also be published for everyone through data/manifest.json in the repo. */
'use strict';

const HEADER = 'APAAR Creation status :: SSM, Purba Bardhaman';
const CREDIT = 'An in-house production of SSM, Purba Bardhaman';
const CAT = { full: '100%', part: '1-99%', zero: '0%', na: 'No students' };
const NUM_TYPES = new Set(['idx', 'n', 'p', 'd', 'dp']);

const state = { snaps: [], cur: '', prev: '', charts: [], cache: {} };

// ---------------------------------------------------------------- helpers
const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
const num = v => { const x = Number(String(v ?? '').replace(/,/g, '')); return isFinite(x) ? x : 0; };
const pct = (o, t) => (t ? (o * 100) / t : 0);
const round2 = n => Math.round(n * 100) / 100;
const fmtN = n => (n == null || n === '' ? '' : Number(n).toLocaleString('en-IN'));
const fmtP = n => (n == null ? '' : round2(n).toFixed(2));
const fmtD = d => (d ? d.split('-').reverse().join('-') : '');
const signed = (n, dec) => { if (n == null) return ''; const v = dec ? round2(n).toFixed(2) : fmtN(n); return n > 0 ? '+' + v : v; };
const pctCls = p => (p >= 80 ? 'good' : p >= 50 ? 'mid' : 'bad');
const slug = s => s.replace(/[^a-z0-9]+/gi, '_').replace(/^_|_$/g, '');
const byDate = (a, b) => a.date.localeCompare(b.date);

function deltaHtml(n, dec, invert) {
    if (n == null) return '<span class="muted">&ndash;</span>';
    const r = round2(n);
    if (r === 0) return '<span class="muted">0</span>';
    const good = (r > 0) !== !!invert;
    return `<span class="${good ? 'up' : 'down'}">${r > 0 ? '&#9650;' : '&#9660;'} ${signed(n, dec)}</span>`;
}

function toast(msg, err) {
    const d = document.createElement('div');
    d.textContent = msg;
    if (err) d.className = 'err';
    $('#toast').append(d);
    setTimeout(() => d.remove(), err ? 7000 : 3500);
}

function dateFromName(name) {
    let m = name.match(/(20\d{2})[-_.](\d{2})[-_.](\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = name.match(/(\d{2})[-_.](\d{2})[-_.](20\d{2})/);
    return m ? `${m[3]}-${m[2]}-${m[1]}` : null;
}

function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.append(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

// ---------------------------------------------------------------- storage (IndexedDB, falls back to memory)
// "snapshots" is keyed by date; "meta" holds the uploaded school master (key 'master')
const STORES = { snap: 'snapshots', meta: 'meta' };
const DB = {
    db: null,
    mem: Object.fromEntries(Object.keys(STORES).map(k => [k, new Map()])),
    open() {
        return new Promise(res => {
            setTimeout(res, 3000); // some browsers/private modes never answer; fall back to memory
            try {
                const r = indexedDB.open('apaar-monitor', 1);
                r.onupgradeneeded = () => Object.entries(STORES).forEach(([kind, n]) => {
                    if (!r.result.objectStoreNames.contains(n)) r.result.createObjectStore(n, { keyPath: kind === 'meta' ? 'key' : 'date' });
                });
                r.onsuccess = () => { this.db = r.result; res(); };
                r.onerror = () => res();
            } catch (e) { res(); }
        });
    },
    tx(kind, mode, fn) {
        return new Promise((res, rej) => {
            const t = this.db.transaction(STORES[kind], mode);
            const req = fn(t.objectStore(STORES[kind]));
            t.oncomplete = () => res(req && req.result);
            t.onerror = () => rej(t.error);
        });
    },
    all(kind) { return this.db ? this.tx(kind, 'readonly', s => s.getAll()) : Promise.resolve([...this.mem[kind].values()]); },
    put(kind, v) { if (!this.db) { this.mem[kind].set(v.date ?? v.key, v); return Promise.resolve(); } return this.tx(kind, 'readwrite', s => s.put(v)); },
    del(kind, k) { if (!this.db) { this.mem[kind].delete(k); return Promise.resolve(); } return this.tx(kind, 'readwrite', s => s.delete(k)); },
    clear(kind) { if (!this.db) { this.mem[kind].clear(); return Promise.resolve(); } return this.tx(kind, 'readwrite', s => s.clear()); },
};

// ---------------------------------------------------------------- Excel parsing
// APAAR Entry Status - School Wise: Block Name, School Name, UDISE Code, School Management, School Category,
// Total Student, Total Generated, Total Requested, Total Failed. The title rows above the header carry
// "Data as on : DD-MM-YYYY", which is used as the data date.
const isApaarHeader = H => H.some(x => /generated/i.test(x)) && H.some(x => /total\s*student/i.test(x));

function parseWorkbook(buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    for (const name of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true });
        const hi = rows.findIndex(r => r.some(c => /school\s*name/i.test(String(c))));
        if (hi < 0) continue;
        const H = rows[hi].map(c => String(c).trim());
        if (!isApaarHeader(H)) throw new Error(`Sheet "${name}": "Total Student" / "Total Generated" columns not found. Upload the APAAR Entry Status - School Wise file.`);
        const find = (...tests) => H.findIndex(x => tests.every(t => t(x)));
        const iBlock = find(x => /block\s*name/i.test(x));
        const iSchool = find(x => /school\s*name/i.test(x));
        const iUdise = find(x => /udise/i.test(x));
        const iMg = find(x => /management/i.test(x));
        const iCt = find(x => /category/i.test(x));
        const iTotal = find(x => /total\s*student/i.test(x));
        const iGen = find(x => /generated/i.test(x));
        const iReq = find(x => /requested/i.test(x));
        const iFail = find(x => /failed/i.test(x));
        let asOn = null;
        rows.slice(0, hi).forEach(r => r.forEach(c => {
            const m = String(c).match(/data\s*as\s*on\s*:?\s*(\d{2})[-/.](\d{2})[-/.](20\d{2})/i);
            if (m) asOn = `${m[3]}-${m[2]}-${m[1]}`;
        }));
        const txt = (r, i) => (i >= 0 ? String(r[i] ?? '').trim() : '');
        const schools = [];
        for (const r of rows.slice(hi + 1)) {
            const n = txt(r, iSchool);
            if (!n || /^\(\d+\)$/.test(n)) continue; // skip the "(1) (2) (3)" column-number row
            schools.push({
                b: txt(r, iBlock) || 'UNKNOWN', n, u: txt(r, iUdise), mc: txt(r, iMg), cc: txt(r, iCt),
                t: num(r[iTotal]), o: num(r[iGen]), rq: iReq >= 0 ? num(r[iReq]) : 0, fl: iFail >= 0 ? num(r[iFail]) : 0,
            });
        }
        if (schools.length) return { asOn, schools };
    }
    throw new Error('No sheet with a "School Name" column was found');
}

// School master (School Search export): UDISE Code, Management, Category, Type.
// Returns [{ u, mg, ct, ty }] or null when the file is not a master file (the APAAR file also has
// UDISE and management columns, so it is told apart by its generated / total student columns).
function parseMaster(buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    for (const name of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true });
        const hi = rows.findIndex(r => r.some(c => /udise/i.test(String(c))) && r.some(c => /management/i.test(String(c))));
        if (hi < 0) continue;
        const H = rows[hi].map(c => String(c).trim());
        if (isApaarHeader(H)) return null;
        const iU = H.findIndex(x => /udise/i.test(x)), iM = H.findIndex(x => /management/i.test(x));
        const iC = H.findIndex(x => /category/i.test(x)), iT = H.findIndex(x => /^(school\s*)?type$/i.test(x));
        const txt = (r, i) => (i >= 0 ? String(r[i] ?? '').trim() : '');
        const out = rows.slice(hi + 1).map(r => ({ u: txt(r, iU), mg: txt(r, iM), ct: txt(r, iC), ty: txt(r, iT) })).filter(x => x.u);
        if (out.length) return out;
    }
    return null;
}

// ---------------------------------------------------------------- calculations
function category(x) {
    if (!x.t) return 'na';
    if (x.o >= x.t) return 'full';
    if (x.o <= 0) return 'zero';
    return 'part';
}
const SUMS = ['t', 'o', 'rq', 'fl', 'full', 'part', 'zero', 'na'];
const blank = block => ({ block, schools: 0, t: 0, o: 0, rq: 0, fl: 0, full: 0, part: 0, zero: 0, na: 0 });
function addTo(a, x) { a.schools++; a.t += x.t; a.o += x.o; a.rq += x.rq; a.fl += x.fl; a[x.c]++; }
// pend = students without an APAAR ID; nr = of those, not yet requested (and not failed)
function finish(a) { a.p = pct(a.o, a.t); a.pend = Math.max(0, a.t - a.o); a.nr = Math.max(0, a.pend - a.rq - a.fl); return a; }

function enrich(s) {
    if (!s || s._e) return s;
    const blocks = new Map();
    const T = blank('District Total');
    s.schools.forEach(x => {
        x.c = category(x);
        x.p = pct(x.o, x.t);
        x.pend = Math.max(0, x.t - x.o);
        x.nr = Math.max(0, x.pend - x.rq - x.fl);
        x.k = x.u || x.b + '|' + x.n;
        if (!blocks.has(x.b)) blocks.set(x.b, blank(x.b));
        addTo(blocks.get(x.b), x);
        addTo(T, x);
    });
    s.blocks = [...blocks.values()].map(finish).sort((a, b) => a.block.localeCompare(b.block));
    s.total = finish(T);
    s.byKey = new Map(s.schools.map(x => [x.k, x]));
    s.byBlock = new Map(s.blocks.map(b => [b.block, b]));
    Object.defineProperty(s, '_e', { value: true, enumerable: false });
    return s;
}

const snap = d => enrich(state.snaps.find(s => s.date === d));
const cur = () => snap(state.cur);
const prev = () => (state.prev ? snap(state.prev) : null);
const blockNames = () => cur().blocks.map(b => b.block);

function memo(key, fn) {
    if (!(key in state.cache)) state.cache[key] = fn();
    return state.cache[key];
}

// Blocks ranked by generated %, with change vs P and Leading/Moderate/Lagging status (top/bottom quarter)
function rankedBlocks(S, P) {
    return memo(`rank|${S.date}|${P ? P.date : ''}`, () => {
        const rows = S.blocks.map(b => {
            const pb = P && P.byBlock.get(b.block);
            return {
                ...b,
                dO: pb ? b.o - pb.o : null,
                dP: pb ? b.p - pb.p : null,
                dFull: pb ? b.full - pb.full : null,
                dZero: pb ? b.zero - pb.zero : null,
            };
        }).sort((a, b) => b.p - a.p || b.o - a.o);
        const n = rows.length, q = Math.ceil(n / 4);
        rows.forEach((r, i) => {
            r.rank = i + 1;
            r.status = i < q ? 'Leading' : i >= n - q ? 'Lagging' : 'Moderate';
        });
        return rows;
    });
}

function schoolRows(S, P) {
    return memo(`sch|${S.date}|${P ? P.date : ''}`, () => S.schools.map(x => {
        const a = P && P.byKey.get(x.k);
        return { ...x, oPrev: a ? a.o : null, pPrev: a ? a.p : null, dO: a ? x.o - a.o : null, cPrev: a ? a.c : null };
    }));
}

function totalsRow(rows, label = 'District Total') {
    const T = blank(label);
    rows.forEach(r => { T.schools += r.schools ?? 1; SUMS.forEach(k => { T[k] += r[k] || 0; }); });
    finish(T);
    T.dO = rows.some(r => r.dO != null) ? rows.reduce((s, r) => s + (r.dO || 0), 0) : null;
    return T;
}

// ---------------------------------------------------------------- Management / Category / Type groups
// Management and Category come from the codes in the APAAR file (UDISE+ codes, named below); a code not
// listed here takes its name from the school master. School Type is only in the school master (by UDISE).
const GROUPS = { mg: 'Management', ct: 'Category', ty: 'School Type' };
const NO_MASTER = 'Not in school master';
const MG_CODES = {
    1: 'Department of Education', 3: 'Local Body', 5: 'Private Unaided (Recognized)', 8: 'Unrecognized',
    12: 'SSKs & MSKs', 13: 'Madrasa Siksha Kendra', 15: 'Govt. Sponsored', 16: 'Mass Education',
    89: 'Minority Affairs Department', 97: 'Madrasa Private and Aided (Recognized)', 99: 'Madrasa Aided (Recognized)',
};
const CT_CODES = {
    1: 'Primary only with grades 1 to 5 (PRY)', 2: 'Upper Primary with grades 1 to 8 (PRY-UPR)',
    3: 'Higher Secondary with grades 1 to 12 (PRY-UPR-SEC-HSEC)', 4: 'Upper Primary only with grades 6 to 8 (UPR)',
    5: 'Higher Secondary with grades 6 to 12 (UPR-SEC-HSEC)', 6: 'Secondary with grades 1 to 10 (PRY-UPR-SEC)',
    7: 'Secondary with grades 6 to 10 (UPR-SEC)', 8: 'Secondary only with grades 9 & 10 (SEC)',
    10: 'Higher Secondary with grades 9 to 12 (SEC-HSEC)', 11: 'Higher Secondary only with grades 11 & 12 (HSEC)',
};
function groupOf(x, by) {
    const m = state.master && state.master.get(x.u);
    if (by === 'ty') return (m && m.ty) || NO_MASTER;
    const code = by === 'mg' ? x.mc : x.cc;
    const named = (by === 'mg' ? MG_CODES : CT_CODES)[code];
    return named || (m && m[by]) || (code ? `${GROUPS[by]} code ${code}` : 'Not specified');
}
const shortCat = c => { const m = String(c || '').match(/\(([^)]+)\)\s*$/); return m ? m[1] : c || ''; };
const groupKeys = () => Object.keys(GROUPS).filter(k => k !== 'ty' || state.master);

// Rows shaped like block rows (group name in "block") so the block columns and totals can be reused
function groupRows(S, P, by, block) {
    return memo(`grp|${by}|${block}|${S.date}|${P ? P.date : ''}`, () => {
        const agg = X => {
            const map = new Map();
            X.schools.forEach(x => {
                if (block && x.b !== block) return;
                const g = groupOf(x, by);
                if (!map.has(g)) map.set(g, blank(g));
                addTo(map.get(g), x);
            });
            map.forEach(finish);
            return map;
        };
        const now = agg(S), before = P ? agg(P) : null;
        return [...now.values()].map(g => {
            const pg = before && before.get(g.block);
            return { ...g, dO: pg ? g.o - pg.o : null, dP: pg ? g.p - pg.p : null };
        }).sort((a, b) => b.schools - a.schools);
    });
}

// School list filters by group (?mg=..&ct=..&ty=..)
const groupParams = r => groupKeys().filter(k => r.q.get(k)).map(k => [k, r.q.get(k)]);
const inGroups = (x, gf) => gf.every(([k, val]) => groupOf(x, k) === val);
const groupTitle = gf => gf.map(([, val]) => ` - ${val}`).join('');
function listHash(cat, block, gf) {
    const q = new URLSearchParams();
    if (block) q.set('block', block);
    gf.forEach(([k, val]) => { if (val) q.set(k, val); });
    const s = q.toString();
    return `#/schools${cat && cat !== 'all' ? '/' + cat : ''}${s ? '?' + s : ''}`;
}
function groupSelects(S, gf) {
    return groupKeys().map(k => {
        const sel = (gf.find(g => g[0] === k) || [])[1] || '';
        const vals = [...new Set(S.schools.map(x => groupOf(x, k)))].sort();
        return `<label>${GROUPS[k]}<select data-grp="${k}"><option value="">All</option>${vals.map(x => `<option${x === sel ? ' selected' : ''}>${esc(x)}</option>`).join('')}</select></label>`;
    }).join('');
}
function wireGroupSelects(f, cat, block) {
    f.querySelectorAll('[data-grp]').forEach(s => {
        s.onchange = () => {
            const gf = [...f.querySelectorAll('[data-grp]')].map(x => [x.dataset.grp, x.value]).filter(g => g[1]);
            location.hash = listHash(cat, block, gf);
        };
    });
}
const groupCols = () => [
    { k: 'mg', h: 'Management', cls: 'wrap', v: x => groupOf(x, 'mg') },
    { k: 'ct', h: 'Category', cls: 'nowrap', v: x => shortCat(groupOf(x, 'ct')) },
];

// ---------------------------------------------------------------- column definitions
const C = {
    idx: { k: '_i', h: '#', type: 'idx' },
    rank: { k: 'rank', h: 'Rank', type: 'n' },
    block: { k: 'block', h: 'Block', cls: 'nowrap' },
    sb: { k: 'b', h: 'Block', cls: 'nowrap' },
    school: { k: 'n', h: 'School Name', cls: 'wrap' },
    udise: { k: 'u', h: 'UDISE Code', cls: 'nowrap' },
    schools: { k: 'schools', h: 'Schools', type: 'n' },
    t: { k: 't', h: 'Total Students', type: 'n' },
    o: { k: 'o', h: 'APAAR Generated', type: 'n' },
    pend: { k: 'pend', h: 'Not Generated', type: 'n', invert: true },
    p: { k: 'p', h: 'Generated %', type: 'p' },
    rq: { k: 'rq', h: 'Requested', type: 'n' },
    fl: { k: 'fl', h: 'Failed', type: 'n' },
    nr: { k: 'nr', h: 'Not Requested', type: 'n' },
    dO: { k: 'dO', h: 'Change (Gen.)', type: 'd' },
    dP: { k: 'dP', h: 'Change (% pt)', type: 'dp' },
    full: { k: 'full', h: '100% Sch.', type: 'n' },
    part: { k: 'part', h: '1-99% Sch.', type: 'n' },
    zero: { k: 'zero', h: '0% Sch.', type: 'n' },
    status: { k: 'status', h: 'Status', type: 'status' },
    cat: { k: 'c', h: 'Level', type: 'cat' },
};
const col = (base, extra) => ({ ...base, ...extra });

const val = (c, r) => (c.v ? c.v(r) : r[c.k]);

function cellHtml(c, r, i) {
    if (c.type === 'idx') return i + 1;
    const v = val(c, r);
    switch (c.type) {
        case 'n': return fmtN(v);
        case 'p': return v == null ? '' : `<span class="pcell"><span class="pbar"><span class="c-${pctCls(v)}" style="width:${Math.min(100, v)}%"></span></span>${fmtP(v)}</span>`;
        case 'd': return deltaHtml(v, false, c.invert);
        case 'dp': return deltaHtml(v, true, c.invert);
        case 'cat': return v ? `<span class="badge ${v}">${CAT[v]}</span>` : '';
        case 'status': return v ? `<span class="badge ${v}">${v}</span>` : '';
        default: return esc(c.show ? c.show(v) : v);
    }
}

// Value written to Excel (raw numbers) or PDF (formatted text)
function cellOut(c, r, i, pdf) {
    if (c.type === 'idx') return i + 1;
    const v = val(c, r);
    if (v == null) return '';
    switch (c.type) {
        case 'p': return pdf ? fmtP(v) : round2(v);
        case 'dp': return pdf ? signed(v, true) : round2(v);
        case 'd': return pdf ? signed(v) : v;
        case 'n': return pdf ? fmtN(v) : v;
        case 'cat': return CAT[v] || v;
        default: return c.show ? c.show(v) : v;
    }
}

// ---------------------------------------------------------------- table component
function makeTable({ cols, rows, sort = null, dir = 'desc', pageSize = 0, onRow = null, foot = null, empty = 'No records found' }) {
    const wrap = h('<div><div class="table-wrap"><table><thead></thead><tbody></tbody><tfoot></tfoot></table></div><div class="pager"></div></div>');
    const thead = $('thead', wrap), tbody = $('tbody', wrap), tfoot = $('tfoot', wrap), pager = $('.pager', wrap);
    let sk = sort, sd = dir, page = 0;

    const sorted = () => {
        if (!sk) return rows;
        const c = cols.find(x => x.k === sk);
        return [...rows].sort((a, b) => {
            const x = val(c, a), y = val(c, b);
            const r = NUM_TYPES.has(c.type)
                ? (x ?? -1e15) - (y ?? -1e15)
                : String(x ?? '').localeCompare(String(y ?? ''), 'en', { numeric: true });
            return sd === 'asc' ? r : -r;
        });
    };
    const tdCls = c => [NUM_TYPES.has(c.type) ? 'num' : '', c.cls || ''].join(' ').trim();

    function draw() {
        const all = sorted();
        const pages = pageSize ? Math.max(1, Math.ceil(all.length / pageSize)) : 1;
        page = Math.min(page, pages - 1);
        const start = pageSize ? page * pageSize : 0;
        const list = pageSize ? all.slice(start, start + pageSize) : all;

        thead.innerHTML = '<tr>' + cols.map(c => {
            const cls = [NUM_TYPES.has(c.type) ? 'num' : '', c.type === 'idx' ? '' : 'sortable', sk === c.k ? 'sorted ' + sd : ''].join(' ');
            return `<th data-k="${c.k}" class="${cls}">${esc(c.h)}</th>`;
        }).join('') + '</tr>';

        tbody.innerHTML = list.length
            ? list.map((r, i) => `<tr data-i="${start + i}" class="${onRow ? 'clickable' : ''}">${cols.map(c => `<td class="${tdCls(c)}">${cellHtml(c, r, start + i)}</td>`).join('')}</tr>`).join('')
            : `<tr><td colspan="${cols.length}" class="empty">${empty}</td></tr>`;

        tfoot.innerHTML = foot && list.length
            ? '<tr>' + cols.map(c => `<td class="${tdCls(c)}">${c.type === 'idx' ? '' : cellHtml(c, foot, 0)}</td>`).join('') + '</tr>'
            : '';

        pager.innerHTML = `<span>${fmtN(all.length)} record${all.length === 1 ? '' : 's'}</span>` + (pages > 1
            ? `<span class="pg"><button class="btn sm" data-p="-1" ${page === 0 ? 'disabled' : ''}>&lsaquo; Prev</button> Page ${page + 1} of ${pages} <button class="btn sm" data-p="1" ${page >= pages - 1 ? 'disabled' : ''}>Next &rsaquo;</button></span>`
            : '');
    }

    thead.addEventListener('click', e => {
        const th = e.target.closest('th.sortable');
        if (!th) return;
        const k = th.dataset.k;
        if (sk === k) sd = sd === 'asc' ? 'desc' : 'asc';
        else { sk = k; sd = NUM_TYPES.has(cols.find(c => c.k === k).type) ? 'desc' : 'asc'; }
        page = 0;
        draw();
    });
    pager.addEventListener('click', e => {
        const b = e.target.closest('[data-p]');
        if (!b) return;
        page += Number(b.dataset.p);
        draw();
        wrap.scrollIntoView({ block: 'start', behavior: 'smooth' });
    });
    if (onRow) tbody.addEventListener('click', e => {
        const tr = e.target.closest('tr[data-i]');
        if (tr) onRow(sorted()[Number(tr.dataset.i)]);
    });

    draw();
    return {
        el: wrap,
        cols,
        get foot() { return foot; },
        rows: sorted,
        setRows(r, f) { rows = r; foot = f === undefined ? foot : f; page = 0; draw(); },
    };
}

// ---------------------------------------------------------------- export (Excel / PDF)
// The four title rows and the blank row above every Excel report
const excelHead = (title, sub) => [
    [HEADER],
    [title],
    [`Data as on: ${fmtD(state.cur)}${sub ? '  |  ' + sub : ''}`],
    [`Generated on: ${new Date().toLocaleString('en-IN')}  |  ${CREDIT}`],
    [],
];

function exportExcel(spec) {
    const { title, cols, rows, foot, sub } = spec();
    const n = cols.length;
    const aoa = [
        ...excelHead(title, sub),
        cols.map(c => c.h),
        ...rows.map((r, i) => cols.map(c => cellOut(c, r, i, false))),
    ];
    if (foot) aoa.push(cols.map((c, i) => (c.type === 'idx' ? (i === 0 ? 'Total' : '') : cellOut(c, foot, 0, false))));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [0, 1, 2, 3].map(r => ({ s: { r, c: 0 }, e: { r, c: Math.max(0, n - 1) } }));
    ws['!cols'] = cols.map((c, j) => {
        let w = c.h.length;
        for (let i = 6; i < Math.min(aoa.length, 400); i++) w = Math.max(w, String(aoa[i][j] ?? '').length);
        return { wch: Math.min(60, w + 2) };
    });
    if (rows.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 5, c: 0 }, e: { r: 5 + rows.length, c: n - 1 } }) };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, slug(title).slice(0, 31) || 'Report');
    XLSX.writeFile(wb, `${slug(title)}_${state.cur}.xlsx`);
}

function exportPDF(spec) {
    const { title, cols, rows, foot, sub } = spec();
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: cols.length > 6 ? 'landscape' : 'portrait', unit: 'mm', format: 'a4' });
    const TP = '{total_pages}';
    const generated = new Date().toLocaleString('en-IN');
    const colStyles = {};
    cols.forEach((c, i) => { if (NUM_TYPES.has(c.type)) colStyles[i] = { halign: 'right' }; });

    doc.autoTable({
        head: [cols.map(c => c.h)],
        body: rows.map((r, i) => cols.map(c => cellOut(c, r, i, true))),
        foot: foot ? [cols.map((c, i) => (c.type === 'idx' ? (i === 0 ? 'Total' : '') : cellOut(c, foot, 0, true)))] : undefined,
        showFoot: 'lastPage',
        startY: 34,
        margin: { top: 34, left: 8, right: 8, bottom: 14 },
        ...PDF_STYLE,
        styles: { ...PDF_STYLE.styles, fontSize: cols.length > 10 ? 7 : 8 },
        columnStyles: colStyles,
        didParseCell: d => {
            if (d.section !== 'body' && NUM_TYPES.has(cols[d.column.index].type)) d.cell.styles.halign = 'right';
        },
        didDrawPage: () => pdfFrame(doc, title, sub, generated, TP),
    });
    if (typeof doc.putTotalPages === 'function') doc.putTotalPages(TP);
    doc.save(`${slug(title)}_${state.cur}.pdf`);
}

const PDF_STYLE = {
    styles: { fontSize: 8, cellPadding: 1.6, overflow: 'linebreak', lineColor: [226, 232, 240], lineWidth: 0.1, textColor: [30, 41, 59] },
    headStyles: { fillColor: [30, 58, 138], textColor: 255, fontStyle: 'bold' },
    footStyles: { fillColor: [224, 233, 255], textColor: [15, 23, 42], fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [248, 250, 252] },
};

// Header (portal name, report title, data date) and footer drawn on every PDF page
function pdfFrame(doc, title, sub, generated, TP) {
    const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13);
    doc.setTextColor(30, 58, 138);
    doc.text(HEADER, W / 2, 11, { align: 'center' });
    doc.setFontSize(11);
    doc.setTextColor(15, 23, 42);
    doc.text(title, W / 2, 18, { align: 'center' });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.text(`Data as on: ${fmtD(state.cur)}${sub ? '   |   ' + sub : ''}`, W / 2, 24, { align: 'center' });
    doc.setDrawColor(30, 58, 138);
    doc.setLineWidth(0.5);
    doc.line(8, 28, W - 8, 28);
    doc.setFontSize(7.5);
    doc.setTextColor(100, 116, 139);
    doc.text(`Generated on ${generated}`, 8, H - 6);
    doc.text(CREDIT, W / 2, H - 6, { align: 'center' });
    doc.text(`Page ${doc.internal.getNumberOfPages()} of ${TP}`, W - 8, H - 6, { align: 'right' });
}

function exportBtns(spec) {
    const d = h('<div class="actions"><button class="btn xls">&#128215; Excel</button><button class="btn pdf">&#128213; PDF</button></div>');
    const run = fn => () => {
        try { fn(spec); } catch (e) { console.error(e); toast('Export failed: ' + e.message, true); }
    };
    d.children[0].onclick = run(exportExcel);
    d.children[1].onclick = run(exportPDF);
    return d;
}
// Exported reports are always ordered by block name, then school name
const cmpText = (a, b) => String(a ?? '').localeCompare(String(b ?? ''), 'en', { numeric: true, sensitivity: 'base' });
function exportOrder(rows) {
    if (!rows.length) return rows;
    if ('b' in rows[0]) return [...rows].sort((x, y) => cmpText(x.b, y.b) || cmpText(x.n, y.n));
    if ('block' in rows[0]) return [...rows].sort((x, y) => cmpText(x.block, y.block));
    return rows;
}
const tableSpec = (tbl, title, sub) => () => ({
    title: typeof title === 'function' ? title() : title,
    sub: typeof sub === 'function' ? sub() : sub,
    cols: tbl.cols,
    rows: exportOrder(tbl.rows()),
    foot: tbl.foot,
});

// ---------------------------------------------------------------- layout helpers
function pageHead(v, title, sub, actions) {
    const d = h(`<div class="page-head"><div><h1></h1>${sub ? `<p class="sub">${sub}</p>` : ''}</div></div>`);
    $('h1', d).textContent = title;
    if (actions) d.append(actions);
    v.append(d);
    return d;
}
function card(parent, title, actions, flush) {
    const c = h(`<section class="card"><div class="card-head"><h2></h2></div><div class="card-body${flush ? ' flush' : ''}"></div></section>`);
    $('h2', c).textContent = title;
    if (actions) $('.card-head', c).append(actions);
    parent.append(c);
    return $('.card-body', c);
}
function tableCard(parent, title, opts, sub) {
    const tbl = makeTable(opts);
    const body = card(parent, title, exportBtns(tableSpec(tbl, title, sub)), true);
    body.append(tbl.el);
    return tbl;
}
const subLine = (S, P) => `Data as on <b>${fmtD(S.date)}</b>${P ? ` &middot; changes compared with <b>${fmtD(P.date)}</b>` : ''}`;
const plainSub = (S, P) => (P ? `Compared with: ${fmtD(P.date)}` : '');

function chart(parent, height, config) {
    const box = h(`<div class="chart-box" style="height:${height}px"><canvas></canvas></div>`);
    parent.append(box);
    config.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: { duration: 300 } }, config.options);
    state.charts.push(new Chart($('canvas', box), config));
}
const COLORS = { good: '#16a34a', mid: '#f59e0b', bad: '#dc2626', na: '#94a3b8', brand: '#1e3a8a', brand2: '#2563eb' };

const ICONS = {
    student: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
    check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    school: '<path d="M3 21h18M5 21V10l7-5 7 5v11M10 21v-5h4v5"/>',
    full: '<path d="M20 6 9 17l-5-5"/>',
    part: '<circle cx="12" cy="12" r="9"/><path d="M12 3v9h9"/>',
    zero: '<circle cx="12" cy="12" r="9"/><path d="M5.6 5.6l12.8 12.8"/>',
    up: '<path d="m3 17 6-6 4 4 8-8"/><path d="M14 7h7v7"/>',
    down: '<path d="m3 7 6 6 4-4 8 8"/><path d="M14 17h7v-7"/>',
    info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>',
    flag: '<path d="M12 3 2 20h20L12 3Z"/><path d="M12 10v4M12 17h.01"/>',
};
const icon = n => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[n]}</svg>`;

function statCard({ label, value, delta, invert, pctDelta, foot, ic, tone, href }) {
    const d = delta == null ? '' : `<span class="stat-delta">${deltaHtml(delta, pctDelta, invert)} <small>since ${fmtD(state.prev)}</small></span>`;
    return `<a class="stat ${tone}"${href ? ` href="${href}"` : ''}>
        <span class="stat-icon">${icon(ic)}</span>
        <span class="stat-body"><span class="stat-label">${label}</span><span class="stat-val">${value}</span>${d}<span class="stat-foot">${foot}</span></span>
    </a>`;
}

// ---------------------------------------------------------------- pages
function pageDashboard(v) {
    const S = cur(), P = prev(), T = S.total, PT = P && P.total;
    const d = k => (PT ? T[k] - PT[k] : null);
    const ranked = rankedBlocks(S, P);
    const count = s => ranked.filter(b => b.status === s).length;
    pageHead(v, 'District Dashboard', subLine(S, P));

    // hero: district progress
    const chg = PT ? `<span class="hero-chip ${T.p >= PT.p ? 'pos' : 'neg'}">${icon(T.p >= PT.p ? 'up' : 'down')} ${signed(T.p - PT.p, true)} pt &middot; ${signed(T.o - PT.o)} APAAR IDs since ${fmtD(P.date)}</span>` : '';
    v.append(h(`<section class="hero">
        <div class="ring" style="--p:${Math.min(100, T.p)}"><div class="ring-in"><b>${fmtP(T.p)}%</b><small>generated</small></div></div>
        <div class="hero-main">
            <div class="hero-label">District APAAR creation &middot; Purba Bardhaman</div>
            <div class="hero-title"><b>${fmtN(T.o)}</b> of ${fmtN(T.t)} students have an APAAR ID</div>
            <div class="hero-bar"><span style="width:${Math.min(100, T.p)}%"></span></div>
            <div class="hero-meta">${fmtN(T.pend)} students still without an APAAR ID across ${fmtN(T.schools)} schools</div>
            ${chg}
        </div>
        <div class="hero-stats">
            <div><b>${S.blocks.length}</b><span>Blocks</span></div>
            <div><b>${fmtN(T.schools)}</b><span>Schools</span></div>
            <div><b class="g">${count('Leading')}</b><span>Leading blocks</span></div>
            <div><b class="r">${count('Lagging')}</b><span>Lagging blocks</span></div>
        </div>
    </section>`));

    // student KPIs
    v.append(h(`<div class="stat-grid">
        ${statCard({ label: 'Total Students', value: fmtN(T.t), delta: d('t'), foot: `in ${fmtN(T.schools)} schools`, ic: 'student', tone: 'blue' })}
        ${statCard({ label: 'APAAR Generated', value: fmtN(T.o), delta: d('o'), foot: `${fmtP(T.p)}% of all students`, ic: 'check', tone: 'green' })}
        ${statCard({ label: 'Not Generated', value: fmtN(T.pend), delta: d('pend'), invert: true, foot: `${fmtP(100 - T.p)}% yet to get an APAAR ID`, ic: 'clock', tone: 'red' })}
        ${statCard({ label: 'Generated %', value: fmtP(T.p) + '%', delta: d('p'), pctDelta: true, foot: `best block ${fmtP(ranked[0].p)}%`, ic: 'up', tone: 'indigo' })}
    </div>`));
    const k = (label, key, o = {}) => `<a class="kpi ${o.cls || ''}"${o.href ? ` href="${o.href}"` : ''}><span class="kpi-label">${label}</span><span class="kpi-val">${fmtN(T[key])}</span><span class="kpi-delta">${PT ? deltaHtml(T[key] - PT[key], false, o.invert) + ` <small>since ${fmtD(P.date)}</small>` : `<small>${o.foot || ''}</small>`}</span></a>`;
    v.append(h(`<div class="kpis">
        ${k('Requested (in process)', 'rq', { cls: 'mid', foot: `in ${fmtN(S.schools.filter(x => x.rq > 0).length)} schools`, href: '#/schools/requested' })}
        ${k('Failed', 'fl', { cls: 'bad', invert: true, foot: `in ${fmtN(S.schools.filter(x => x.fl > 0).length)} schools`, href: '#/schools/failed' })}
        ${k('Not yet requested', 'nr', { invert: true, foot: 'without ID, request or failure' })}
        ${k('Schools with no students', 'na', { foot: 'Total Student = 0', href: '#/schools/na' })}
    </div>`));

    // school categories
    const catCard = (key, label, ic, invert) => {
        const n = T[key], share = pct(n, T.schools);
        const dd = d(key);
        return `<a class="cat-card ${key}" href="#/schools/${key}">
            <div class="cat-top"><span class="cat-icon">${icon(ic)}</span><span class="cat-label">${label}</span><span class="cat-link">View list &rsaquo;</span></div>
            <div class="cat-val">${fmtN(n)} <small>schools</small></div>
            <div class="cat-bar"><span style="width:${share}%"></span></div>
            <div class="cat-foot"><span>${fmtP(share)}% of schools</span>${dd == null ? '' : deltaHtml(dd, false, invert)}</div>
        </a>`;
    };
    v.append(h(`<div class="cat-grid">
        ${catCard('full', '100% APAAR Generated', 'full')}
        ${catCard('part', 'Partially Generated (1-99%)', 'part')}
        ${catCard('zero', '0% APAAR Generated', 'zero', true)}
    </div>`));

    // key highlights
    const best = ranked[0], worst = ranked[ranked.length - 1];
    const most = [...ranked].sort((a, b) => b.pend - a.pend)[0];
    const notes = [
        ['up', `Highest generation: <b>${esc(best.block)}</b> at ${fmtP(best.p)}%; lowest: <b>${esc(worst.block)}</b> at ${fmtP(worst.p)}%.`],
        ['zero', `<b>${fmtN(T.zero)}</b> schools (${fmtP(pct(T.zero, T.schools))}%) have not generated any APAAR ID. <a href="#/schools/zero">See list &rsaquo;</a>`],
        ['flag', `Most students without APAAR: <b>${esc(most.block)}</b> (${fmtN(most.pend)} students).`],
    ];
    if (T.fl) notes.push(['info', `<b>${fmtN(T.fl)}</b> APAAR requests have failed in ${fmtN(S.schools.filter(x => x.fl > 0).length)} schools and need correction. <a href="#/schools/failed">See list &rsaquo;</a>`]);
    if (P) {
        const sr = schoolRows(S, P);
        const newFull = sr.filter(x => x.c === 'full' && x.cPrev && x.cPrev !== 'full').length;
        const outZero = sr.filter(x => x.cPrev === 'zero' && x.c !== 'zero').length;
        const improved = sr.filter(x => x.dO > 0).length;
        const regressed = sr.filter(x => x.dO < 0).length;
        const still = ranked.filter(b => b.dO === 0).map(b => b.block);
        const gainer = [...ranked].sort((a, b) => (b.dO || 0) - (a.dO || 0))[0];
        notes.unshift(['check', `Since ${fmtD(P.date)}: <b>${fmtN(improved)}</b> schools improved, <b>${fmtN(newFull)}</b> became 100% and <b>${fmtN(outZero)}</b> moved out of 0%. <a href="#/compare">Details &rsaquo;</a>`]);
        if (gainer && gainer.dO > 0) notes.push(['up', `Top gainer since ${fmtD(P.date)}: <b>${esc(gainer.block)}</b> (${signed(gainer.dO)} APAAR IDs).`]);
        if (still.length) notes.push(['info', `No progress since ${fmtD(P.date)} in ${still.length} block(s): ${still.map(esc).join(', ')}.`]);
        if (regressed) notes.push(['down', `<span class="down">${fmtN(regressed)} school(s) show fewer APAAR IDs than on ${fmtD(P.date)}</span> &mdash; please verify.`]);
    }
    card(v, 'Key Highlights').append(h(`<ul class="highlights">${notes.map(([i, t]) => `<li><span class="hl-icon">${icon(i)}</span><span>${t}</span></li>`).join('')}</ul>`));

    // charts
    const g = h('<div class="grid-2"></div>');
    v.append(g);
    chart(card(g, 'Schools by APAAR Generation Level'), 260, {
        type: 'doughnut',
        data: {
            labels: ['100%', '1-99%', '0%', 'No students'],
            datasets: [{ data: [T.full, T.part, T.zero, T.na], backgroundColor: [COLORS.good, COLORS.mid, COLORS.bad, COLORS.na], borderWidth: 2, borderColor: '#fff' }],
        },
        options: { plugins: { legend: { position: window.innerWidth < 600 ? 'bottom' : 'right', labels: { usePointStyle: true, padding: 14 } } }, cutout: '64%' },
    });
    const bands = [['Below 20%', 0, 20], ['20-40%', 20, 40], ['40-60%', 40, 60], ['60-80%', 60, 80], ['80% and above', 80, 101]];
    chart(card(g, 'Blocks by Generation Level'), 260, {
        type: 'bar',
        data: {
            labels: bands.map(b => b[0]),
            datasets: [{ label: 'Blocks', data: bands.map(([, lo, hi]) => ranked.filter(b => b.p >= lo && b.p < hi).length), backgroundColor: [COLORS.bad, '#ef6c00', COLORS.mid, COLORS.brand2, COLORS.good], borderRadius: 6, maxBarThickness: 56 }],
        },
        options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 }, title: { display: true, text: 'Number of blocks' } }, x: { grid: { display: false } } } },
    });
    const byP = [...ranked];
    chart(card(v, 'Block-wise APAAR Generated %'), Math.max(320, byP.length * 20 + 60), {
        type: 'bar',
        data: {
            labels: byP.map(b => b.block),
            datasets: [{ label: 'Generated %', data: byP.map(b => round2(b.p)), backgroundColor: byP.map(b => COLORS[pctCls(b.p)]), borderRadius: 3, maxBarThickness: 18 }],
        },
        options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, max: 100, ticks: { callback: x => x + '%' } }, y: { ticks: { autoSkip: false, font: { size: 10 } } } } },
    });
    if (state.snaps.length > 1) {
        const all = state.snaps.map(s => enrich(s));
        chart(card(v, 'District Daily Trend'), 280, {
            type: 'line',
            data: {
                labels: all.map(s => fmtD(s.date)),
                datasets: [
                    { label: 'Generated %', data: all.map(s => round2(s.total.p)), borderColor: COLORS.brand, backgroundColor: 'rgba(26,35,126,.10)', fill: true, tension: .3, yAxisID: 'y', pointRadius: 3 },
                    { label: '0% schools', data: all.map(s => s.total.zero), borderColor: COLORS.bad, tension: .3, yAxisID: 'y1', pointRadius: 3 },
                    { label: '100% schools', data: all.map(s => s.total.full), borderColor: COLORS.good, tension: .3, yAxisID: 'y1', pointRadius: 3 },
                ],
            },
            options: { interaction: { mode: 'index', intersect: false }, plugins: { legend: { labels: { usePointStyle: true } } }, scales: { y: { ticks: { callback: x => x + '%' } }, y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Schools' } } } },
        });
    }

    const g2 = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g2);
    const mini = [C.rank, C.block, C.o, C.p, col(C.dO, { h: 'Change' })];
    const go = r => { location.hash = '#/schools?block=' + encodeURIComponent(r.block); };
    tableCard(g2, 'Top 5 Leading Blocks', { cols: mini, rows: ranked.slice(0, 5), onRow: go }, plainSub(S, P));
    tableCard(g2, 'Bottom 5 Lagging Blocks', { cols: mini, rows: ranked.slice(-5).reverse(), onRow: go }, plainSub(S, P));
    groupSummary(v, S, P);
}

function pageBlocks(v) {
    const S = cur(), P = prev();
    const rows = rankedBlocks(S, P);
    const cols = [C.rank, C.block, C.schools, C.t, C.o, C.pend, C.p, C.dO, C.dP, C.rq, C.fl, C.full, C.part, C.zero, C.status];
    const foot = Object.assign(totalsRow(rows), { dP: P ? S.total.p - P.total.p : null });
    const tbl = makeTable({ cols, rows, foot, onRow: r => { location.hash = '#/schools?block=' + encodeURIComponent(r.block); } });
    pageHead(v, 'Block-wise APAAR Status', subLine(S, P) + ' &middot; tap a block to see its schools', exportBtns(tableSpec(tbl, 'Block-wise APAAR Status', plainSub(S, P))));
    v.append(h(`<p class="help">Status: <span class="badge Leading">Leading</span> = top quarter of blocks by generated %, <span class="badge Lagging">Lagging</span> = bottom quarter, <span class="badge Moderate">Moderate</span> = the rest.
        <b>Not Generated</b> = Total Students &minus; APAAR Generated; it includes the Requested and Failed students.</p>`));
    card(v, 'All Blocks', null, true).append(tbl.el);

    const byName = [...rows].sort((a, b) => a.block.localeCompare(b.block));
    chart(card(v, 'School generation levels by block'), Math.max(320, rows.length * 20 + 60), {
        type: 'bar',
        data: {
            labels: byName.map(b => b.block),
            datasets: [
                { label: '100%', data: byName.map(b => b.full), backgroundColor: COLORS.good },
                { label: '1-99%', data: byName.map(b => b.part), backgroundColor: COLORS.mid },
                { label: '0%', data: byName.map(b => b.zero), backgroundColor: COLORS.bad },
                { label: 'No students', data: byName.map(b => b.na), backgroundColor: COLORS.na },
            ],
        },
        options: { indexAxis: 'y', scales: { x: { stacked: true }, y: { stacked: true, ticks: { autoSkip: false, font: { size: 10 } } } }, interaction: { mode: 'index', intersect: false } },
    });
}

// School lists: title and filter of each list
const SCHOOL_LISTS = {
    all: ['All Schools', 'All schools', () => true],
    full: ['100% APAAR Generated Schools', '100% generated', x => x.c === 'full'],
    part: ['Partially Generated Schools (1-99%)', '1-99% generated', x => x.c === 'part'],
    zero: ['0% APAAR Generated Schools', '0% generated', x => x.c === 'zero'],
    na: ['Schools with No Students', 'No students', x => x.c === 'na'],
    requested: ['Schools with APAAR Requests in Process', 'Has requests in process', x => x.rq > 0],
    failed: ['Schools with Failed APAAR Requests', 'Has failed requests', x => x.fl > 0],
};

function pageSchools(v, r) {
    const S = cur(), P = prev();
    const catKey = SCHOOL_LISTS[r.arg] ? r.arg : 'all';
    const inList = SCHOOL_LISTS[catKey][2];
    const block = r.q.get('block') || '';
    let search = '';
    const all = schoolRows(S, P);
    const gf = groupParams(r);
    const filtered = () => all.filter(x => (!block || x.b === block) && inList(x) && inGroups(x, gf) &&
        (!search || (x.n + ' ' + x.u).toLowerCase().includes(search)));
    const title = SCHOOL_LISTS[catKey][0] + (block ? ` - ${block} Block` : '') + groupTitle(gf);
    const go = (c, b) => { location.hash = listHash(c, b, gf); };

    const cols = [C.idx];
    if (!block) cols.push(C.sb);
    cols.push(C.school, C.udise, ...groupCols(), C.t, C.o, C.pend, C.p, C.rq, C.fl);
    if (P) cols.push({ k: 'oPrev', h: `Gen. on ${fmtD(P.date)}`, type: 'n' }, C.dO);
    if (catKey === 'all') cols.push(C.cat);

    const defaultSort = { zero: ['t', 'desc'], part: ['p', 'asc'], failed: ['fl', 'desc'], requested: ['rq', 'desc'] }[catKey] || [null];
    const tbl = makeTable({ cols, rows: filtered(), sort: defaultSort[0], dir: defaultSort[1], pageSize: 50 });
    pageHead(v, title, subLine(S, P), exportBtns(tableSpec(tbl, () => title + (search ? ` (search: ${search})` : ''), plainSub(S, P))));

    const f = h(`<div class="filters">
        <label>Block<select id="fBlock"><option value="">All blocks</option>${blockNames().map(b => `<option${b === block ? ' selected' : ''}>${esc(b)}</option>`).join('')}</select></label>
        <label>APAAR status<select id="fCat">${Object.entries(SCHOOL_LISTS).map(([k, x]) => `<option value="${k}"${k === catKey ? ' selected' : ''}>${x[1]}</option>`).join('')}</select></label>
        ${groupSelects(S, gf)}
        <label>Search school / UDISE<input type="search" id="fSearch" placeholder="Type to search..."></label>
    </div>`);
    v.append(f);
    $('#fBlock', f).onchange = e => go(catKey, e.target.value);
    $('#fCat', f).onchange = e => go(e.target.value, block);
    wireGroupSelects(f, catKey, block);

    const chips = h('<div class="chips"></div>');
    v.append(chips);
    const drawChips = rows => {
        const t = totalsRow(rows);
        chips.innerHTML = `<span class="chip">Schools <b>${fmtN(rows.length)}</b></span><span class="chip">Students <b>${fmtN(t.t)}</b></span>
            <span class="chip">Generated <b>${fmtN(t.o)}</b></span><span class="chip">Not generated <b>${fmtN(t.pend)}</b></span><span class="chip">Generated % <b>${fmtP(t.p)}</b></span>
            <span class="chip">Requested <b>${fmtN(t.rq)}</b></span><span class="chip">Failed <b>${fmtN(t.fl)}</b></span>`;
    };
    drawChips(tbl.rows());
    let timer;
    $('#fSearch', f).oninput = e => {
        clearTimeout(timer);
        timer = setTimeout(() => { search = e.target.value.trim().toLowerCase(); const rows = filtered(); tbl.setRows(rows); drawChips(rows); }, 200);
    };
    card(v, 'Schools', null, true).append(tbl.el);
}

function pageCompare(v, r) {
    if (state.snaps.length < 2) {
        pageHead(v, 'Day-wise Comparison');
        v.append(h('<div class="card"><div class="card-body">Upload data for at least two dates to compare.</div></div>'));
        return;
    }
    const dates = state.snaps.map(s => s.date);
    const bDate = r.q.get('b') || state.cur;
    const aDate = r.q.get('a') || state.prev || dates[Math.max(0, dates.indexOf(bDate) - 1)];
    const A = snap(aDate), B = snap(bDate);
    const opts = sel => dates.slice().reverse().map(d => `<option value="${d}"${d === sel ? ' selected' : ''}>${fmtD(d)}</option>`).join('');
    const sub = `${fmtD(A.date)} vs ${fmtD(B.date)}`;
    pageHead(v, 'Day-wise Comparison', `Changes from <b>${fmtD(A.date)}</b> to <b>${fmtD(B.date)}</b>`);

    const f = h(`<div class="filters"><label>From date<select id="cA">${opts(aDate)}</select></label><label>To date<select id="cB">${opts(bDate)}</select></label></div>`);
    v.append(f);
    const go = () => { location.hash = `#/compare?a=${$('#cA').value}&b=${$('#cB').value}`; };
    $('#cA', f).onchange = go;
    $('#cB', f).onchange = go;

    const TA = A.total, TB = B.total;
    const k = (label, key, o = {}) => `<div class="kpi ${o.cls || ''}"><span class="kpi-label">${label}</span><span class="kpi-val">${o.pct ? fmtP(TB[key]) + '%' : fmtN(TB[key])}</span><span class="kpi-delta">${deltaHtml(TB[key] - TA[key], o.pct, o.invert)} <small>from ${o.pct ? fmtP(TA[key]) + '%' : fmtN(TA[key])}</small></span></div>`;
    v.append(h(`<div class="kpis">${k('APAAR Generated', 'o', { cls: 'good' })}${k('Generated %', 'p', { pct: true })}${k('Total Students', 't')}${k('Failed', 'fl', { invert: true, cls: 'bad' })}${k('100% Schools', 'full', { cls: 'good' })}${k('1-99% Schools', 'part', { cls: 'mid' })}${k('0% Schools', 'zero', { invert: true, cls: 'bad' })}</div>`));

    // block comparison
    const names = [...new Set([...A.blocks, ...B.blocks].map(b => b.block))];
    const rows = names.map(n => {
        const a = A.byBlock.get(n) || finish(blank(n)), b = B.byBlock.get(n) || finish(blank(n));
        return { block: n, t: b.t, oA: a.o, o: b.o, dO: b.o - a.o, pA: a.p, p: b.p, dP: b.p - a.p, fullA: a.full, full: b.full, zeroA: a.zero, zero: b.zero, dZero: b.zero - a.zero };
    });
    const tot = { block: 'District Total', t: TB.t, oA: TA.o, o: TB.o, dO: TB.o - TA.o, pA: TA.p, p: TB.p, dP: TB.p - TA.p, fullA: TA.full, full: TB.full, zeroA: TA.zero, zero: TB.zero, dZero: TB.zero - TA.zero };
    const dA = fmtD(A.date), dB = fmtD(B.date);
    tableCard(v, 'Block-wise Comparison', {
        cols: [C.idx, C.block, C.t, { k: 'oA', h: `Gen. ${dA}`, type: 'n' }, { k: 'o', h: `Gen. ${dB}`, type: 'n' }, C.dO,
            { k: 'pA', h: `% ${dA}`, type: 'p' }, { k: 'p', h: `% ${dB}`, type: 'p' }, C.dP,
            { k: 'fullA', h: `100% ${dA}`, type: 'n' }, { k: 'full', h: `100% ${dB}`, type: 'n' },
            { k: 'zeroA', h: `0% ${dA}`, type: 'n' }, { k: 'zero', h: `0% ${dB}`, type: 'n' }, { k: 'dZero', h: 'Change (0% sch.)', type: 'd', invert: true }],
        rows, sort: 'dO', foot: tot,
    }, sub);

    // school movements
    const sr = schoolRows(B, A);
    const groups = [
        ['improved', 'Schools improved', x => x.dO > 0],
        ['newfull', 'Became 100%', x => x.c === 'full' && x.cPrev && x.cPrev !== 'full'],
        ['outzero', 'Moved out of 0%', x => x.cPrev === 'zero' && x.c !== 'zero'],
        ['stillzero', 'Still 0% on both dates', x => x.c === 'zero' && x.cPrev === 'zero'],
        ['regressed', 'Generated count decreased', x => x.dO < 0],
        ['new', 'New in later file', x => x.cPrev == null],
    ].map(([key, label, fn]) => ({ key, label, rows: sr.filter(fn) }));
    let active = groups[0];
    const cols = [C.idx, C.sb, C.school, C.udise, C.t, { k: 'oPrev', h: `Gen. ${dA}`, type: 'n' }, { k: 'o', h: `Gen. ${dB}`, type: 'n' }, C.dO, C.p, C.cat];
    const tbl = makeTable({ cols, rows: active.rows, sort: 'dO', pageSize: 50 });
    const body = card(v, 'School-level Movement', exportBtns(() => ({ ...tableSpec(tbl, `School Movement - ${active.label}`, sub)() })), true);
    const tabs = h(`<div class="tabs">${groups.map((g, i) => `<button class="tab${i === 0 ? ' active' : ''}" data-i="${i}">${g.label}<span class="count">${fmtN(g.rows.length)}</span></button>`).join('')}</div>`);
    tabs.onclick = e => {
        const b = e.target.closest('.tab');
        if (!b) return;
        active = groups[Number(b.dataset.i)];
        tabs.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t === b));
        tbl.setRows(active.rows);
    };
    body.append(tabs, tbl.el);
}

function pageTrend(v) {
    const all = state.snaps.map(s => enrich(s));
    const rows = all.map((s, i) => {
        const p = i ? all[i - 1].total : null, t = s.total;
        return { date: s.date, schools: t.schools, t: t.t, o: t.o, dO: p ? t.o - p.o : null, p: t.p, dP: p ? t.p - p.p : null, rq: t.rq, fl: t.fl, full: t.full, part: t.part, zero: t.zero, pend: t.pend };
    }).reverse();
    pageHead(v, 'Daily Trend', `${all.length} day(s) of data, from <b>${fmtD(all[0].date)}</b> to <b>${fmtD(all[all.length - 1].date)}</b>`);

    if (all.length > 1) {
        chart(card(v, 'District progress by date'), 280, {
            type: 'bar',
            data: {
                labels: all.map(s => fmtD(s.date)),
                datasets: [
                    { type: 'line', label: 'Generated %', data: all.map(s => round2(s.total.p)), borderColor: COLORS.brand, backgroundColor: COLORS.brand, yAxisID: 'y1', tension: .25 },
                    { label: 'APAAR generated', data: all.map(s => s.total.o), backgroundColor: 'rgba(57,73,171,.35)', yAxisID: 'y' },
                ],
            },
            options: { interaction: { mode: 'index', intersect: false }, scales: { y: { beginAtZero: true }, y1: { position: 'right', grid: { drawOnChartArea: false }, ticks: { callback: x => x + '%' } } } },
        });
    }

    tableCard(v, 'District Date-wise Summary', {
        cols: [{ k: 'date', h: 'Date', cls: 'nowrap', show: fmtD }, C.schools, C.t, C.o, col(C.dO, { h: 'Daily change' }), C.p, C.dP, C.rq, C.fl, C.full, C.part, C.zero, C.pend],
        rows,
    });

    // block x date matrix
    const metricBox = card(v, 'Block-wise Trend (by date)');
    const sel = h('<div class="filters"><label>Show<select><option value="p">Generated %</option><option value="o">APAAR generated</option><option value="pend">Not generated</option><option value="fl">Failed</option><option value="zero">0% schools</option><option value="full">100% schools</option></select></label></div>');
    metricBox.append(sel);
    const holder = h('<div></div>');
    metricBox.append(holder);
    const draw = metric => {
        holder.innerHTML = '';
        const last = all[all.length - 1], first = all[0];
        const names = last.blocks.map(b => b.block);
        const type = metric === 'p' ? 'p' : 'n';
        const cols = [C.idx, C.block, ...all.map(s => ({ k: s.date, h: fmtD(s.date), type })),
            { k: 'chg', h: 'Change (first to last)', type: metric === 'p' ? 'dp' : 'd', invert: ['zero', 'pend', 'fl'].includes(metric) }];
        const rows = names.map(n => {
            const r = { block: n };
            all.forEach(s => { const b = s.byBlock.get(n); r[s.date] = b ? b[metric] : null; });
            r.chg = r[last.date] != null && r[first.date] != null ? r[last.date] - r[first.date] : null;
            return r;
        });
        const label = sel.querySelector('option:checked').textContent;
        const tbl = makeTable({ cols, rows, sort: last.date });
        const bar = h('<div class="actions" style="margin-bottom:8px"></div>');
        bar.append(exportBtns(tableSpec(tbl, `Block-wise Trend - ${label}`, `${fmtD(first.date)} to ${fmtD(last.date)}`)));
        holder.append(bar, tbl.el);
    };
    $('select', sel).onchange = e => draw(e.target.value);
    draw('p');
}

function pageRank(v) {
    const S = cur(), P = prev();
    const ranked = rankedBlocks(S, P);
    const sub = plainSub(S, P);
    pageHead(v, 'Leading & Lagging', subLine(S, P));
    const go = r => { location.hash = '#/schools?block=' + encodeURIComponent(r.block || r.b); };
    const bcols = [C.rank, C.block, C.t, C.o, C.pend, C.p, C.dO];

    const g1 = h('<div class="grid-2"></div>');
    v.append(g1);
    tableCard(g1, 'Top 10 Leading Blocks (by generated %)', { cols: bcols, rows: ranked.slice(0, 10), onRow: go }, sub);
    tableCard(g1, 'Bottom 10 Lagging Blocks (by generated %)', { cols: bcols, rows: ranked.slice(-10).reverse(), onRow: go }, sub);

    const g2 = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g2);
    if (P) {
        tableCard(g2, `Top Gainers since ${fmtD(P.date)}`, { cols: [C.idx, C.block, C.dO, C.dP, C.p, { k: 'dFull', h: 'New 100% sch.', type: 'd' }], rows: [...ranked].sort((a, b) => b.dO - a.dO).slice(0, 10), onRow: go }, sub);
        tableCard(g2, `Least Progress since ${fmtD(P.date)}`, { cols: [C.idx, C.block, C.dO, C.dP, C.p, C.pend], rows: [...ranked].sort((a, b) => a.dO - b.dO || a.p - b.p).slice(0, 10), onRow: go }, sub);
    }
    tableCard(g2, 'Blocks with Most Students without APAAR', { cols: [C.idx, C.block, C.t, C.o, C.pend, C.p], rows: [...ranked].sort((a, b) => b.pend - a.pend).slice(0, 10), onRow: go }, sub);
    tableCard(g2, 'Blocks with Most 0% Schools', { cols: [C.idx, C.block, C.schools, C.zero, { k: 'zp', h: '% of schools at 0%', type: 'p', v: r => pct(r.zero, r.schools) }, C.p], rows: [...ranked].sort((a, b) => b.zero - a.zero).slice(0, 10), onRow: go }, sub);
    tableCard(g2, 'Blocks with Most Failed Requests', { cols: [C.idx, C.block, C.fl, C.rq, C.o, C.p], rows: [...ranked].filter(b => b.fl > 0).sort((a, b) => b.fl - a.fl).slice(0, 10), onRow: go }, sub);

    const sr = schoolRows(S, P);
    const scols = [C.idx, C.sb, C.school, C.udise, C.t, C.o, C.pend, C.p];
    const g3 = h('<div class="stack" style="margin-top:16px"></div>');
    v.append(g3);
    tableCard(g3, 'Priority Schools - Most Students without APAAR (Top 50)', { cols: scols, rows: [...sr].sort((a, b) => b.pend - a.pend).slice(0, 50) }, sub);
    tableCard(g3, 'Large Schools at 0% (Top 50 by students)', { cols: scols, rows: sr.filter(x => x.c === 'zero').sort((a, b) => b.t - a.t).slice(0, 50) }, sub);
    tableCard(g3, 'Schools with Most Failed Requests (Top 50)', { cols: [...scols, C.fl], rows: sr.filter(x => x.fl > 0).sort((a, b) => b.fl - a.fl).slice(0, 50) }, sub);
    if (P) tableCard(g3, `Best Performing Schools since ${fmtD(P.date)} (Top 50)`, { cols: [...scols, C.dO], rows: sr.filter(x => x.dO > 0).sort((a, b) => b.dO - a.dO).slice(0, 50) }, sub);
}

// ---------------------------------------------------------------- Management / Category-wise report
const groupLink = (by, block) => r => {
    const q = new URLSearchParams({ [by]: r.block });
    if (block) q.set('block', block);
    location.hash = `#/schools?${q}`;
};

function pageGroups(v, r) {
    const S = cur(), P = prev();
    const by = groupKeys().includes(r.q.get('by')) ? r.q.get('by') : 'mg';
    const block = r.q.get('block') || '';
    const label = GROUPS[by];
    const title = `APAAR - ${label}-wise Report${block ? ` - ${block} Block` : ''}`;
    pageHead(v, `${label}-wise Report`, subLine(S, P) + ' &middot; tap a row to see its schools');

    const f = h(`<div class="filters">
        <label>Report by<select id="gBy">${groupKeys().map(k => `<option value="${k}"${k === by ? ' selected' : ''}>${GROUPS[k]}</option>`).join('')}</select></label>
        <label>Block<select id="gBlock"><option value="">All blocks</option>${blockNames().map(b => `<option${b === block ? ' selected' : ''}>${esc(b)}</option>`).join('')}</select></label>
    </div>`);
    v.append(f);
    const go = (b, bl) => { location.hash = `#/groups?by=${b}${bl ? '&block=' + encodeURIComponent(bl) : ''}`; };
    $('#gBy', f).onchange = e => go(e.target.value, block);
    $('#gBlock', f).onchange = e => go(by, e.target.value);

    const rows = groupRows(S, P, by, block);
    const foot = totalsRow(rows);
    if (P) foot.dP = foot.p - totalsRow(groupRows(P, null, by, block)).p;
    const g = col(C.block, { h: label, cls: 'wrap' });
    tableCard(v, title, { cols: [C.idx, g, C.schools, C.t, C.o, C.pend, C.p, C.dO, C.dP, C.rq, C.fl, C.full, C.part, C.zero], rows, foot, sort: 'schools', onRow: groupLink(by, block) }, plainSub(S, P));

    chart(card(v, `Generated % by ${label}`), Math.max(220, rows.length * 34 + 60), {
        type: 'bar',
        data: {
            labels: rows.map(x => (by === 'ct' ? shortCat(x.block) : x.block)),
            datasets: [{ label: 'Generated %', data: rows.map(x => round2(x.p)), backgroundColor: rows.map(x => COLORS[pctCls(x.p)] || COLORS.brand2), borderRadius: 4, maxBarThickness: 26 }],
        },
        options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, max: 100, ticks: { callback: x => x + '%' } }, y: { ticks: { autoSkip: false } } } },
    });

    // Block x group matrix
    const groups = rows.map(x => x.block);
    const metrics = [['p', 'Generated %'], ['o', 'APAAR generated'], ['pend', 'Not generated'], ['t', 'Total students'], ['fl', 'Failed'], ['zero', '0% schools'], ['full', '100% schools'], ['schools', 'Total schools']];
    const box = card(v, `Block-wise by ${label}`);
    const sel = h(`<div class="filters"><label>Show<select>${metrics.map(([k, t]) => `<option value="${k}">${t}</option>`).join('')}</select></label></div>`);
    const holder = h('<div></div>');
    box.append(sel, holder);
    const cells = new Map(); // "block|group" -> aggregate
    S.schools.forEach(x => {
        const key = x.b + '|' + groupOf(x, by);
        if (!cells.has(key)) cells.set(key, blank(key));
        addTo(cells.get(key), x);
    });
    cells.forEach(finish);
    const draw = metric => {
        holder.innerHTML = '';
        const type = metric === 'p' ? 'p' : 'n';
        const head = x => (by === 'ct' ? shortCat(x) : x);
        const cols = [C.idx, C.block, ...groups.map((x, i) => ({ k: 'g' + i, h: head(x), type })), { k: 'all', h: 'All', type }];
        const mrows = S.blocks.map(b => {
            const row = { block: b.block, all: b[metric] };
            groups.forEach((x, i) => { const c = cells.get(b.block + '|' + x); row['g' + i] = c ? c[metric] : null; });
            return row;
        });
        const dist = groupRows(S, null, by, '');
        const mfoot = { block: 'District Total', all: S.total[metric] };
        groups.forEach((x, i) => { const d = dist.find(y => y.block === x); mfoot['g' + i] = d ? d[metric] : null; });
        const tbl = makeTable({ cols, rows: mrows, foot: mfoot });
        const bar = h('<div class="actions" style="margin-bottom:8px"></div>');
        const mLabel = sel.querySelector('option:checked').textContent;
        bar.append(exportBtns(tableSpec(tbl, `APAAR - Block-wise by ${label} - ${mLabel}`, plainSub(S, P))));
        holder.append(bar, tbl.el);
    };
    $('select', sel).onchange = e => draw(e.target.value);
    draw('p');
    if (by === 'ct') v.append(h(`<p class="help">Category codes: ${groups.filter(x => shortCat(x) !== x).map(x => `<b>${esc(shortCat(x))}</b> = ${esc(x)}`).join(' &middot; ')}</p>`));
}

// Management and Category summary tables shown at the end of the dashboard
function groupSummary(v, S, P) {
    const g = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g);
    ['mg', 'ct'].forEach(by => {
        const label = GROUPS[by];
        const gc = col(C.block, { h: label, cls: 'wrap', v: by === 'ct' ? x => shortCat(x.block) : null });
        const tbl = tableCard(g, `APAAR by ${label}`, { cols: [gc, C.t, C.o, C.p, col(C.dO, { h: 'Change' })], rows: groupRows(S, P, by, ''), sort: 't', onRow: () => { location.hash = `#/groups?by=${by}`; } }, plainSub(S, P));
        tbl.el.append(h(`<p class="help" style="padding:0 16px 12px"><a href="#/groups?by=${by}">Full ${label.toLowerCase()}-wise report &rsaquo;</a></p>`));
    });
}

// ---------------------------------------------------------------- Block x Management report (fits two slides)
// Dept. of Education, SSK & MSK and Govt. Sponsored schools side by side, every other management together
const MG4 = [
    { k: 'doe', h: 'Dept. of Education', names: ['Department of Education'], tone: '#2563eb' },
    { k: 'ssk', h: 'SSK & MSK', names: ['SSKs & MSKs'], tone: '#7c3aed' },
    { k: 'gs', h: 'Govt. Sponsored', names: ['Govt. Sponsored'], tone: '#0891b2' },
    { k: 'oth', h: 'Other Managements', names: null, tone: '#d97706' },
    { k: 'all', h: 'All Managements', names: null, tone: '#0b1f44' },
];
const MG_SUB = [['schools', 'Schools'], ['t', 'Enrolment'], ['o', 'APAAR'], ['pend', 'Pending'], ['p', '%']];
const MG_SHORT = ['Sch.', 'Enrol.', 'APAAR', 'Pending', '%']; // headings for the narrow PDF / slide columns
const MG_TITLE = 'APAAR Creation Status - Block & Management-wise';
const mg4Of = x => { const n = groupOf(x, 'mg'); const g = MG4.find(m => m.names && m.names.includes(n)); return g ? g.k : 'oth'; };

function mgReport(S) {
    return memo(`mg4|${S.date}`, () => {
        const mk = name => Object.fromEntries(MG4.map(g => [g.k, blank(name)]));
        const rows = S.blocks.map(b => ({ block: b.block, g: mk(b.block) }));
        const byB = new Map(rows.map(r => [r.block, r]));
        const tot = { block: 'District Total', g: mk('District Total') };
        const others = new Map();
        S.schools.forEach(x => {
            const k = mg4Of(x);
            [byB.get(x.b), tot].forEach(r => { addTo(r.g[k], x); addTo(r.g.all, x); });
            if (k === 'oth') { const n = groupOf(x, 'mg'); others.set(n, (others.get(n) || 0) + 1); }
        });
        [...rows, tot].forEach(r => Object.values(r.g).forEach(finish));
        return { rows, tot, others: [...others].sort((a, b) => b[1] - a[1]) };
    });
}

// One cell: { text, cls } where cls marks a % above (hi) / below (lo) the district % of the same management
function mgCell(R, r, gk, sk, total) {
    const a = r.g[gk];
    if (!a.schools) return { text: '-', cls: '' };
    if (sk !== 'p') return { text: fmtN(a[sk]), cls: '' };
    const d = R.tot.g[gk].p;
    return { text: fmtP(a.p), cls: total || !a.t ? '' : a.p >= d ? 'hi' : 'lo' };
}
const othersNote = R => R.others.map(([n, c]) => `${n} (${fmtN(c)})`).join(', ');

function pageMgmt(v) {
    const S = cur(), R = mgReport(S);
    const btns = h('<div class="actions"><button class="btn xls">Excel</button><button class="btn pdf">PDF</button><button class="btn ppt">PowerPoint (2 slides)</button></div>');
    const run = fn => async e => {
        const b = e.currentTarget;
        b.disabled = true;
        try { await fn(R); } catch (err) { console.error(err); toast('Export failed: ' + err.message, true); }
        b.disabled = false;
    };
    btns.children[0].onclick = run(mgExcel);
    btns.children[1].onclick = run(mgPDF);
    btns.children[2].onclick = run(mgPPT);
    pageHead(v, 'Block & Management-wise Report', `Data as on <b>${fmtD(S.date)}</b> &middot; schools, enrolment, APAAR created, pending and % for each management`, btns);

    // district tiles, one per management group
    v.append(h(`<div class="mg-tiles">${MG4.map(g => {
        const a = R.tot.g[g.k];
        return `<div class="mg-tile ${g.k}" style="--tone:${g.tone}">
            <h3>${esc(g.h)}</h3>
            <div class="mg-pct">${fmtP(a.p)}% <small>created</small></div>
            <div class="cat-bar"><span style="width:${Math.min(100, a.p)}%"></span></div>
            <dl><dt>Schools</dt><dd>${fmtN(a.schools)}</dd><dt>Enrolment</dt><dd>${fmtN(a.t)}</dd><dt>APAAR created</dt><dd>${fmtN(a.o)}</dd><dt>Pending</dt><dd>${fmtN(a.pend)}</dd>
            <dt>Share of pending</dt><dd>${fmtP(pct(a.pend, R.tot.g.all.pend))}%</dd></dl>
        </div>`;
    }).join('')}</div>`));
    v.append(h(`<p class="help mg-note"><b>Other Managements:</b> ${esc(othersNote(R)) || 'none'}.</p>`));

    const head1 = `<tr><th rowspan="2" class="sticky">#</th><th rowspan="2" class="sticky2">Block</th>${MG4.map(g => `<th colspan="${MG_SUB.length}" class="g ${g.k === 'all' ? 'all' : ''}">${esc(g.h)}</th>`).join('')}</tr>`;
    const head2 = `<tr>${MG4.map(g => MG_SUB.map(([, t], i) => `<th class="num${i === 0 ? ' gs' : ''}${g.k === 'all' ? ' all' : ''}">${t}</th>`).join('')).join('')}</tr>`;
    const line = (r, i, total) => `<td class="sticky num">${total ? '' : i + 1}</td><td class="sticky2 nowrap">${esc(r.block)}</td>` + MG4.map(g => MG_SUB.map(([sk], j) => {
        const c = mgCell(R, r, g.k, sk, total);
        return `<td class="num${j === 0 ? ' gs' : ''}${g.k === 'all' ? ' all' : ''}${sk === 'p' ? ' pc ' + c.cls : ''}">${c.text}</td>`;
    }).join('')).join('');
    const body = card(v, `Block-wise by Management (${R.rows.length} blocks)`, null, true);
    body.append(h(`<div class="table-wrap"><table class="mg-table"><thead>${head1}${head2}</thead>
        <tbody>${R.rows.map((r, i) => `<tr class="clickable" data-b="${esc(r.block)}">${line(r, i)}</tr>`).join('')}</tbody>
        <tfoot><tr>${line(R.tot, 0, true)}</tr></tfoot></table></div>`));
    body.onclick = e => { const tr = e.target.closest('tr[data-b]'); if (tr) location.hash = '#/schools?block=' + encodeURIComponent(tr.dataset.b); };
    v.append(h(`<p class="help"><span class="legend"><i style="background:var(--good)"></i>% at or above the district % of the same management</span>
        <span class="legend"><i style="background:var(--bad)"></i>below it</span> &middot; Pending = Enrolment &minus; APAAR created &middot; tap a block to see its schools.</p>`));
}

// Flat rows for Excel / PDF: Sl, Block, then Schools .. % for each management group
const mgFlat = (R, r, i, total, pdf) => [total ? '' : i + 1, r.block, ...MG4.flatMap(g => MG_SUB.map(([sk]) => {
    const a = r.g[g.k];
    if (!a.schools) return pdf ? '-' : '';
    return sk === 'p' ? (pdf ? fmtP(a.p) : round2(a.p)) : pdf ? fmtN(a[sk]) : a[sk];
}))];

function mgExcel(R) {
    const n = 2 + MG4.length * MG_SUB.length;
    const aoa = [
        ...excelHead(MG_TITLE, 'Pending = Enrolment - APAAR created'),
        ['Sl', 'Block', ...MG4.flatMap(g => [g.h, ...Array(MG_SUB.length - 1).fill('')])],
        ['', '', ...MG4.flatMap(() => MG_SUB.map(s => s[1]))],
        ...R.rows.map((r, i) => mgFlat(R, r, i)),
        mgFlat(R, R.tot, 0, true),
        [],
        [`Other Managements: ${othersNote(R)}`],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [
        ...[0, 1, 2, 3].map(r => ({ s: { r, c: 0 }, e: { r, c: n - 1 } })),
        { s: { r: 5, c: 0 }, e: { r: 6, c: 0 } }, { s: { r: 5, c: 1 }, e: { r: 6, c: 1 } },
        ...MG4.map((g, i) => ({ s: { r: 5, c: 2 + i * MG_SUB.length }, e: { r: 5, c: 1 + (i + 1) * MG_SUB.length } })),
    ];
    ws['!cols'] = [{ wch: 5 }, { wch: 22 }, ...MG4.flatMap(() => [{ wch: 8 }, { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 7 }])];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Block x Management');
    XLSX.writeFile(wb, `${slug(MG_TITLE)}_${state.cur}.xlsx`);
}

function mgPDF(R) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const TP = '{total_pages}', generated = new Date().toLocaleString('en-IN');
    const pIdx = new Set(MG4.map((g, i) => 2 + i * MG_SUB.length + MG_SUB.length - 1));
    const gOf = c => Math.floor((c - 2) / MG_SUB.length);
    doc.autoTable({
        head: [
            [{ content: 'Sl', rowSpan: 2 }, { content: 'Block', rowSpan: 2 }, ...MG4.map(g => ({ content: g.h, colSpan: MG_SUB.length, styles: { halign: 'center' } }))],
            MG4.flatMap(() => MG_SHORT),
        ],
        body: R.rows.map((r, i) => mgFlat(R, r, i, false, true)),
        foot: [mgFlat(R, R.tot, 0, true, true)],
        showFoot: 'lastPage',
        startY: 34,
        margin: { top: 34, left: 6, right: 6, bottom: 14 },
        ...PDF_STYLE,
        styles: { ...PDF_STYLE.styles, fontSize: 6.6, cellPadding: 1.1, halign: 'right', overflow: 'visible' },
        columnStyles: {
            0: { halign: 'center', cellWidth: 7 }, 1: { halign: 'left', cellWidth: 26 },
            ...Object.fromEntries(MG4.flatMap((g, i) => [8.5, 11, 10.5, 11, 9.4].map((w, j) => [2 + i * MG_SUB.length + j, { cellWidth: w }]))),
        },
        didParseCell: d => {
            const c = d.column.index;
            if (d.section !== 'body' && c === 1) d.cell.styles.halign = 'left';
            if (d.section === 'head' && c > 1) d.cell.styles.halign = d.row.index === 0 ? 'center' : 'right';
            if (d.section === 'head' && d.row.index === 1) Object.assign(d.cell.styles, { fillColor: [224, 233, 255], textColor: [30, 58, 138] });
            if (c < 2 || d.section !== 'body') return;
            if (MG4[gOf(c)].k === 'all') d.cell.styles.fillColor = [238, 243, 255];
            if (pIdx.has(c)) {
                const r = R.rows[d.row.index], a = r.g[MG4[gOf(c)].k];
                if (a.schools && a.t) d.cell.styles.textColor = a.p >= R.tot.g[MG4[gOf(c)].k].p ? [21, 128, 61] : [185, 28, 28];
                d.cell.styles.fontStyle = 'bold';
            }
        },
        didDrawPage: () => pdfFrame(doc, MG_TITLE, 'Pending = Enrolment - APAAR created', generated, TP),
    });
    doc.setFontSize(7);
    doc.setTextColor(100, 116, 139);
    doc.text(doc.splitTextToSize(`Other Managements: ${othersNote(R)}.  Green / red %: at or above / below the district % of the same management.`, 280), 8, doc.lastAutoTable.finalY + 5);
    if (typeof doc.putTotalPages === 'function') doc.putTotalPages(TP);
    doc.save(`${slug(MG_TITLE)}_${state.cur}.pdf`);
}

const PPTX_SRC = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
const loadScript = src => new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = res;
    s.onerror = () => rej(new Error('could not load the PowerPoint library; check the internet connection'));
    document.head.append(s);
});

async function mgPPT(R) {
    if (typeof PptxGenJS === 'undefined') await loadScript(PPTX_SRC);
    await buildMgPPT(R).writeFile({ fileName: `${slug(MG_TITLE)}_${state.cur}.pptx` });
}

// Two 16:9 slides: the blocks split in half, the district total on the second slide
function buildMgPPT(R) {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 in
    Object.assign(pptx, { author: 'SSM, Purba Bardhaman', company: 'SSM, Purba Bardhaman', title: MG_TITLE });
    const F = 'Calibri', NAVY = '0B1F44', BLUE = '1E3A8A', LINE = 'D5DCE8';
    const T = R.tot.g.all;
    const half = Math.ceil(R.rows.length / 2);
    const parts = [R.rows.slice(0, half), R.rows.slice(half)];
    const border = { type: 'solid', pt: 0.5, color: LINE };
    const colW = [0.3, 1.3, ...MG4.flatMap(() => [0.36, 0.5, 0.49, 0.5, 0.392])];
    const hdr = (text, o = {}) => ({ text, options: { bold: true, color: 'FFFFFF', fill: { color: BLUE }, align: 'center', ...o } });
    const sub = text => ({ text, options: { bold: true, color: BLUE, fill: { color: 'E0E9FF' }, align: 'center', fontSize: 7 } });
    const generated = new Date().toLocaleString('en-IN');

    parts.forEach((rows, pi) => {
        const s = pptx.addSlide();
        s.background = { color: 'FFFFFF' };
        s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.62, fill: { color: NAVY }, line: { color: NAVY } });
        s.addText(MG_TITLE, { x: 0.25, y: 0.08, w: 9, h: 0.46, fontFace: F, fontSize: 18, bold: true, color: 'FFFFFF', margin: 0 });
        s.addText(`Purba Bardhaman  |  Data as on ${fmtD(state.cur)}  |  Part ${pi + 1} of ${parts.length}`, { x: 8.3, y: 0.08, w: 4.78, h: 0.46, fontFace: F, fontSize: 11, color: 'BFDBFE', align: 'right', margin: 0 });
        s.addText([
            { text: 'District: ', options: { bold: true, color: BLUE } },
            { text: `${fmtN(T.schools)} schools  |  Enrolment ${fmtN(T.t)}  |  APAAR created ${fmtN(T.o)}  |  Pending ${fmtN(T.pend)}  |  ` },
            { text: `${fmtP(T.p)}% created`, options: { bold: true, color: BLUE } },
            { text: '   (' + MG4.filter(g => g.k !== 'all').map(g => `${g.h} ${fmtP(R.tot.g[g.k].p)}%`).join(', ') + ')', options: { color: '64748B' } },
        ], { x: 0.25, y: 0.68, w: 12.83, h: 0.3, fontFace: F, fontSize: 10, color: '334155', margin: 0 });

        const body = rows.map((r, i) => {
            const n = pi * half + i;
            const fill = { color: n % 2 ? 'F8FAFC' : 'FFFFFF' };
            return [
                { text: String(n + 1), options: { align: 'center', fill } },
                { text: r.block, options: { align: 'left', fill } },
                ...MG4.flatMap(g => MG_SUB.map(([sk]) => {
                    const c = mgCell(R, r, g.k, sk);
                    const o = { align: 'right', fill: g.k === 'all' ? { color: n % 2 ? 'E8EEFF' : 'EEF3FF' } : fill };
                    if (sk === 'p') Object.assign(o, { bold: true, color: c.cls === 'hi' ? '15803D' : c.cls === 'lo' ? 'B91C1C' : '334155' });
                    return { text: c.text, options: o };
                })),
            ];
        });
        if (pi === parts.length - 1) body.push([
            { text: '', options: { fill: { color: 'DCE6FF' } } },
            { text: 'District Total', options: { align: 'left', bold: true, fill: { color: 'DCE6FF' } } },
            ...MG4.flatMap(g => MG_SUB.map(([sk]) => ({ text: mgCell(R, R.tot, g.k, sk, true).text, options: { align: 'right', bold: true, fill: { color: 'DCE6FF' } } }))),
        ]);
        s.addTable([
            [hdr('Sl', { rowspan: 2 }), hdr('Block', { rowspan: 2, align: 'left' }), ...MG4.map(g => hdr(g.h, { colspan: MG_SUB.length }))],
            MG4.flatMap(() => MG_SHORT.map(sub)),
            ...body,
        ], { x: 0.25, y: 1.04, w: 12.83, colW, rowH: 0.215, fontFace: F, fontSize: 7.5, color: '1E293B', border, valign: 'middle', margin: [0, 0.04, 0, 0.04], autoPage: false });

        const notes = ['Pending = Enrolment - APAAR created.  % in green: at or above the district % of the same management; red: below.'];
        if (pi === parts.length - 1) notes.unshift(`Other Managements: ${othersNote(R)}.`);
        s.addText(notes.join('\n'), { x: 0.25, y: 6.62, w: 9.6, h: 0.5, fontFace: F, fontSize: 8, color: '64748B', valign: 'top', margin: 0 });
        s.addText(`${CREDIT}\nGenerated on ${generated}`, { x: 9.9, y: 6.62, w: 3.18, h: 0.5, fontFace: F, fontSize: 8, color: '64748B', align: 'right', valign: 'top', margin: 0 });
        s.addShape(pptx.ShapeType.rect, { x: 0, y: 7.38, w: 13.333, h: 0.12, fill: { color: BLUE }, line: { color: BLUE } });
    });
    return pptx;
}

// ---------------------------------------------------------------- Upload / Manage Data password
// Static site: only the SHA-256 hash of the password is kept here; unlocking lasts for the browser tab session.
const ADMIN_HASH = 'ab05e582c7455107115fe4a1de5320598c3e9e879d7a243b3f944bb265ea6573';
const UNLOCK_KEY = 'apaar-admin-unlocked';
let adminUnlocked = false;
const isUnlocked = () => {
    if (adminUnlocked) return true;
    try { return sessionStorage.getItem(UNLOCK_KEY) === ADMIN_HASH; } catch (e) { return false; }
};
async function sha256(text) {
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
    return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('');
}
function lockAdmin() {
    adminUnlocked = false;
    try { sessionStorage.removeItem(UNLOCK_KEY); } catch (e) { /* storage blocked */ }
}

function pageLogin(v) {
    pageHead(v, 'Upload / Manage Data', 'This section is password protected.');
    const box = h(`<div class="card login-card"><div class="card-body">
        <form class="login">
            <label for="admPwd">Admin password</label>
            <input type="password" id="admPwd" autocomplete="current-password" required>
            <button class="btn primary" type="submit">Unlock</button>
            <p class="login-err" hidden>Incorrect password. Please try again.</p>
        </form>
    </div></div>`);
    v.append(box);
    const form = $('form', box), pwd = $('#admPwd', box), err = $('.login-err', box);
    pwd.focus();
    form.onsubmit = async e => {
        e.preventDefault();
        if (!window.crypto || !crypto.subtle) { err.textContent = 'This browser cannot check the password here. Open the portal over https.'; err.hidden = false; return; }
        if (await sha256(pwd.value) !== ADMIN_HASH) { err.hidden = false; pwd.select(); return; }
        adminUnlocked = true;
        try { sessionStorage.setItem(UNLOCK_KEY, ADMIN_HASH); } catch (e2) { /* storage blocked: unlocked for this page load only */ }
        render();
    };
}

function pageData(v) {
    if (!isUnlocked()) { pageLogin(v); return; }
    const lock = h('<div class="actions"><button class="btn">&#128274; Lock</button></div>');
    lock.firstElementChild.onclick = () => { lockAdmin(); render(); };
    pageHead(v, 'Upload / Manage Data', 'Upload the APAAR Entry Status - School Wise Excel file every day. One file is kept per date.', lock);

    const up = card(v, 'Upload daily Excel');
    up.append(h(`<div class="upload">
        <label class="drop" id="drop">
            <input type="file" id="file" accept=".xlsx,.xls,.csv" multiple hidden>
            <span class="drop-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3"/></svg></span>
            <strong>Drag &amp; drop the Excel file(s) here</strong>
            <span class="help">or</span>
            <span class="btn primary">Browse files</span>
            <span class="help drop-note">.xlsx &middot; <code>APAAR_Entry_Status_School_Wise_Only_Operational DD-MM-YYYY ....xlsx</code> as downloaded from the portal, and the School Search <code>.csv</code> (school master) &middot; several files can be uploaded together</span>
        </label>
        <div class="upload-opts">
            <div class="field">
                <label for="upDate">Data date <em>(optional)</em></label>
                <input type="date" id="upDate">
            </div>
            <div class="upload-hint">
                <span class="hl-icon">${icon('info')}</span>
                <p>The date is read from the <b>Data as on</b> line at the top of the file, or else from the file name. Set a date here only to override it when uploading a single file.</p>
            </div>
        </div>
    </div>`));
    const drop = $('#drop', up), file = $('#file', up);
    file.onchange = () => { handleFiles([...file.files]); file.value = ''; };
    drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
    drop.ondragleave = () => drop.classList.remove('over');
    drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); handleFiles([...e.dataTransfer.files]); };

    const delBtn = r => (r.src === 'repo' ? '' : `<button class="btn sm danger" data-del="${r.date}">Delete</button>`);
    const fileCell = r => `${esc(r.file)} ${r.src === 'repo' ? '<span class="badge repo">from repo</span>' : ''}`;
    const listed = state.snaps.map(s => { const e = enrich(s); return { date: s.date, file: s.fileName || '', src: s.source, ...e.total }; }).reverse();

    const list = card(v, `APAAR data (${listed.length} day${listed.length === 1 ? '' : 's'})`, null, true);
    list.append(h(`<div class="table-wrap"><table><thead><tr><th>Date</th><th>File</th><th class="num">Schools</th><th class="num">Students</th><th class="num">Generated</th><th class="num">%</th><th class="num">Requested</th><th class="num">Failed</th><th class="num">100%</th><th class="num">0%</th><th></th></tr></thead><tbody>
        ${listed.map(r => `<tr><td class="nowrap"><b>${fmtD(r.date)}</b></td><td>${fileCell(r)}</td><td class="num">${fmtN(r.schools)}</td><td class="num">${fmtN(r.t)}</td><td class="num">${fmtN(r.o)}</td><td class="num">${fmtP(r.p)}</td><td class="num">${fmtN(r.rq)}</td><td class="num">${fmtN(r.fl)}</td><td class="num">${fmtN(r.full)}</td><td class="num">${fmtN(r.zero)}</td>
        <td class="nowrap">${delBtn(r)}</td></tr>`).join('') || '<tr><td colspan="11" class="empty">No data uploaded yet</td></tr>'}
    </tbody></table></div>`));
    list.onclick = async e => {
        const b = e.target.closest('[data-del]');
        if (!b || !confirm(`Delete APAAR data of ${fmtD(b.dataset.del)}?`)) return;
        await DB.del('snap', b.dataset.del);
        await reload();
        toast('Deleted');
    };

    const mi = state.masterInfo;
    const mBox = card(v, 'School master (School Type, and names of unknown codes)');
    const last = state.snaps[state.snaps.length - 1];
    const unmatched = last ? last.schools.filter(x => !state.master || !state.master.has(x.u)).length : null;
    mBox.append(h(`<div class="help" style="font-size:13px">
        <p><b>Management</b> and <b>Category</b> are taken from the codes in the APAAR file itself. The <b>School Type</b> (and the name of any
        management / category code the portal does not yet know) come from the School Search export (columns UDISE Code, Management, Category, Type),
        matched by UDISE code. Upload a newer CSV here at any time; it replaces the one in use in this browser.</p>
        ${mi ? `<p>In use: <b>${esc(mi.fileName)}</b> ${mi.source === 'repo' ? '<span class="badge repo">from repo</span>' : `(uploaded ${new Date(mi.uploadedAt).toLocaleString('en-IN')})`} &middot; ${fmtN(mi.count)} schools</p>
            ${unmatched == null ? '' : `<p>Latest day: ${unmatched ? `<span class="down">${fmtN(unmatched)} school(s) not found in the master</span>` : 'all schools matched'}</p>`}
            ${mi.source === 'repo' ? '' : '<p><button class="btn sm danger" id="mDel">Remove uploaded master</button></p>'}`
        : '<p class="down">No school master loaded. Upload the School Search CSV to enable the School Type report.</p>'}
    </div>`));
    if ($('#mDel', mBox)) $('#mDel', mBox).onclick = async () => {
        if (!confirm('Remove the uploaded school master from this browser?')) return;
        await DB.del('meta', 'master');
        await reload();
        toast('School master removed');
    };

    const bk = card(v, 'Backup & restore');
    bk.append(h(`<div><p class="help">Uploaded data is saved only in this browser. Download a backup regularly, and restore it on another phone or computer.</p>
        <div class="actions"><button class="btn primary" id="bkDown">&#11015; Download backup (.json)</button>
        <label class="btn">&#11014; Restore backup<input type="file" id="bkUp" accept=".json" hidden></label>
        <button class="btn danger" id="bkClear">Delete all uploaded data</button></div></div>`));
    $('#bkDown', bk).onclick = async () => {
        const snaps = (await DB.all('snap')).sort(byDate);
        const data = {
            app: 'apaar-monitor', version: 1, exported: new Date().toISOString(),
            snapshots: snaps.map(s => ({ date: s.date, fileName: s.fileName, uploadedAt: s.uploadedAt, schools: s.schools.map(({ b, n, u, mc, cc, t, o, rq, fl }) => ({ b, n, u, mc, cc, t, o, rq, fl })) })),
            master: (await DB.all('meta')).find(x => x.key === 'master') || null,
        };
        download(new Blob([JSON.stringify(data)], { type: 'application/json' }), `apaar-monitor-backup_${new Date().toISOString().slice(0, 10)}.json`);
    };
    $('#bkUp', bk).onchange = async e => {
        const f = e.target.files[0];
        if (!f) return;
        try {
            const data = JSON.parse(await f.text());
            if (data.app !== 'apaar-monitor' || !Array.isArray(data.snapshots)) throw new Error('Not a backup file of this portal');
            for (const s of data.snapshots) await DB.put('snap', { ...s, source: 'upload' });
            if (data.master) await DB.put('meta', data.master);
            await reload();
            toast(`Restored ${data.snapshots.length} day(s)`);
        } catch (err) { toast('Restore failed: ' + err.message, true); }
    };
    $('#bkClear', bk).onclick = async () => {
        if (!confirm('Delete ALL uploaded APAAR data and the uploaded school master from this browser? Download a backup first if needed.')) return;
        await DB.clear('snap');
        await DB.clear('meta');
        await reload();
        toast('All uploaded data deleted');
    };

    const gh = card(v, 'Share data with everyone (GitHub)');
    gh.append(h(`<div class="help" style="font-size:13px">
        <p>Data uploaded here stays on this device. To let every officer who opens the portal see the same history, add the daily files to the repository:</p>
        <ol>
            <li>Put the Excel file in the <code>data/</code> folder of the repository. Rename it with the data date, for example <code>APAAR_Entry_Status_School_Wise_2026-10-28.xlsx</code>.</li>
            <li>Add the file name to <code>data/manifest.json</code>, for example:<br><code>["APAAR_Entry_Status_School_Wise_2026-10-28.xlsx", "APAAR_Entry_Status_School_Wise_2026-10-29.xlsx"]</code></li>
            <li>Commit. GitHub Pages updates in a minute or two; those dates then show as <span class="badge repo">from repo</span> for everyone.</li>
        </ol>
    </div>`));
}

function emptyState() {
    return h(`<div class="card empty-state"><h2>No APAAR data yet</h2>
        <p>Upload the <b>APAAR Entry Status - School Wise (Only Operational)</b> Excel file (Block Name, School Name, UDISE Code, School Management, School Category, Total Student, Total Generated, Total Requested, Total Failed) to start monitoring.</p>
        <p><a class="btn primary" href="#/data">&#11014; Upload Excel</a></p>
        <p class="help">${CREDIT}</p></div>`);
}

// ---------------------------------------------------------------- upload handling
async function handleFiles(files) {
    if (!files.length) return;
    const override = $('#upDate') && $('#upDate').value;
    if (override && files.length > 1) { toast('Pick a date only when uploading a single file.', true); return; }
    let changed = false, added = false;
    for (const f of files) {
        try {
            const buf = await f.arrayBuffer();
            const master = parseMaster(buf);
            if (master) {
                await DB.put('meta', { key: 'master', fileName: f.name, uploadedAt: new Date().toISOString(), rows: master });
                changed = true;
                toast(`School master ${f.name}: ${fmtN(master.length)} schools loaded`);
                continue;
            }
            const { asOn, schools } = parseWorkbook(buf);
            const date = override || asOn || dateFromName(f.name);
            if (!date) throw new Error('no date found in the file; choose the "Data date" and upload again');
            const exists = state.snaps.find(s => s.date === date && s.source !== 'repo');
            if (exists && !confirm(`APAAR data for ${fmtD(date)} already exists. Replace it with ${f.name}?`)) continue;
            await DB.put('snap', { date, fileName: f.name, uploadedAt: new Date().toISOString(), source: 'upload', schools });
            changed = added = true;
            toast(`APAAR ${fmtD(date)}: ${fmtN(schools.length)} schools loaded`);
        } catch (e) {
            console.error(e);
            toast(`${f.name}: ${e.message}`, true);
        }
    }
    if (changed) await reload(added);
}

// ---------------------------------------------------------------- data loading
async function loadRepo() {
    try {
        const res = await fetch('data/manifest.json', { cache: 'no-store' });
        if (!res.ok) return [];
        const list = await res.json();
        const out = [];
        for (const e of list) {
            const file = typeof e === 'string' ? e : e.file;
            if (!file) continue;
            try {
                const r = await fetch('data/' + encodeURIComponent(file), { cache: 'no-store' });
                if (!r.ok) throw new Error('HTTP ' + r.status);
                const { asOn, schools } = parseWorkbook(await r.arrayBuffer());
                const date = (typeof e === 'object' && e.date) || dateFromName(file) || asOn;
                if (date) out.push({ date, fileName: file, source: 'repo', schools });
            } catch (err) { toast(`data/${file}: ${err.message}`, true); }
        }
        return out;
    } catch (e) { return []; }
}

// School master: data/school_master.csv in the repo, overridden by an uploaded master in this browser
async function loadRepoMaster() {
    try {
        const r = await fetch('data/school_master.csv', { cache: 'no-store' });
        if (!r.ok) return null;
        const rows = parseMaster(await r.arrayBuffer());
        return rows ? { fileName: 'school_master.csv', source: 'repo', rows } : null;
    } catch (e) { return null; }
}

let repoSnaps = null, repoMaster;
// jumpToLatest: move the selected date to the latest day (after an upload)
async function reload(jumpToLatest) {
    if (repoSnaps === null) [repoSnaps, repoMaster] = await Promise.all([loadRepo(), loadRepoMaster()]);
    const m = (await DB.all('meta')).find(x => x.key === 'master') || repoMaster;
    state.master = m ? new Map(m.rows.map(x => [x.u, x])) : null;
    state.masterInfo = m ? { fileName: m.fileName, source: m.source || 'upload', uploadedAt: m.uploadedAt, count: m.rows.length } : null;
    const map = new Map(repoSnaps.map(s => [s.date, s]));
    (await DB.all('snap')).forEach(s => map.set(s.date, s)); // uploaded data overrides repo data of the same date
    state.snaps = [...map.values()].sort(byDate);
    const dates = state.snaps.map(s => s.date);
    if (jumpToLatest || !dates.includes(state.cur)) setCur(dates[dates.length - 1] || '');
    else if (state.prev && !dates.includes(state.prev)) setCur(state.cur);
    state.cache = {};
    fillDateSelects();
    render();
}

function setCur(d) {
    state.cur = d;
    const i = state.snaps.findIndex(s => s.date === d);
    state.prev = i > 0 ? state.snaps[i - 1].date : '';
}

function fillDateSelects() {
    const cs = $('#curDate'), ps = $('#prevDate');
    const dates = state.snaps.map(s => s.date).reverse();
    cs.innerHTML = dates.length ? dates.map(d => `<option value="${d}">${fmtD(d)}</option>`).join('') : '<option value="">No data</option>';
    cs.value = state.cur;
    ps.innerHTML = '<option value="">None</option>' + dates.filter(d => d !== state.cur).map(d => `<option value="${d}">${fmtD(d)}</option>`).join('');
    ps.value = state.prev;
}

// ---------------------------------------------------------------- routing
const routes = { dashboard: pageDashboard, blocks: pageBlocks, schools: pageSchools, compare: pageCompare, trend: pageTrend, rank: pageRank, groups: pageGroups, mgmt: pageMgmt, data: pageData };

function parseHash() {
    const raw = location.hash.replace(/^#\/?/, '');
    const [path, q = ''] = raw.split('?');
    const parts = path.split('/').filter(Boolean);
    return { name: routes[parts[0]] ? parts[0] : 'dashboard', arg: parts[1] || '', q: new URLSearchParams(q) };
}

function render() {
    state.charts.forEach(c => c.destroy());
    state.charts = [];
    const r = parseHash();
    const key = r.name + (r.arg ? '/' + r.arg : '');
    document.querySelectorAll('#sidebar nav a').forEach(a => a.classList.toggle('active', a.dataset.route === key));
    const link = $(`#sidebar nav a[data-route="${key}"]`) || $(`#sidebar nav a[data-route="${r.name}"]`);
    $('#crumbTitle').textContent = link ? link.textContent.trim() : 'Dashboard';
    const v = $('#view');
    v.innerHTML = '';
    document.body.classList.remove('menu-open');
    try {
        if (!state.snaps.length && r.name !== 'data') v.append(emptyState());
        else routes[r.name](v, r);
    } catch (e) {
        console.error(e);
        v.append(h(`<div class="card"><div class="card-body">Something went wrong: ${esc(e.message)}</div></div>`));
    }
}

async function init() {
    if (typeof XLSX === 'undefined' || !window.jspdf || typeof Chart === 'undefined') {
        $('#view').innerHTML = '<div class="card"><div class="card-body">Could not load required libraries. Please check the internet connection and reload.</div></div>';
        return;
    }
    $('#year').textContent = new Date().getFullYear();
    Object.assign(Chart.defaults, { color: '#64748b', borderColor: '#eef2f7' });
    Chart.defaults.font.family = 'Inter, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';
    Object.assign(Chart.defaults.plugins.tooltip, { backgroundColor: '#0f172a', padding: 10, cornerRadius: 6 });
    $('#menuBtn').onclick = () => document.body.classList.toggle('menu-open');
    $('#backdrop').onclick = () => document.body.classList.remove('menu-open');
    $('#curDate').onchange = e => { setCur(e.target.value); fillDateSelects(); render(); };
    $('#prevDate').onchange = e => { state.prev = e.target.value; render(); };
    window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
    await DB.open();
    await reload();
}

init();
