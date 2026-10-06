'use strict';
// Saxon I-REC portal: browser app. Plain JavaScript, no build step, no backend.
// Staff sign in with Microsoft. The page reads and writes SharePoint as the signed-in
// person. It never talks to I-REC: anything that touches the registry is queued in
// the "IREC Requests" list and carried out by the worker (GitHub Actions, hourly).

(function () {
  const C = window.IREC_PORTAL_CONFIG || {};
  const GRAPH = (C.graphBase || 'https://graph.microsoft.com/v1.0').replace(/\/+$/, '');
  const SCOPES = ['User.Read', 'Sites.ReadWrite.All'];
  const PORTAL_ROLES = ['viewer', 'operator', 'approver'];
  const ORDER = { viewer: 1, operator: 2, approver: 3 };
  const ROLE_LABEL = { viewer: 'Viewer', operator: 'Operator', approver: 'Approver' };
  const LIST_NAMES = { requests: 'IREC Requests', earmarks: 'IREC Earmarks', remarks: 'IREC Remarks', activity: 'IREC Activity' };
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const LABEL = {
    available: 'Available', earmarked: 'Earmarked', held: 'In a draft', redeemed: 'Redeemed',
    requested: 'Queued for the registry', processing: 'Being created', failed: 'Failed', draft: 'Draft',
    awaiting: 'Awaiting approval', 'delete-requested': 'Deletion queued', deleted: 'Deleted', completed: 'Completed',
    submitted: 'With the Issuer', issued: 'Issued', rejected: 'Rejected', withdrawn: 'Withdrawn'
  };
  const TITLES = { dashboard: 'Overview', inventory: 'Inventory', reservations: 'Redemptions and transfers', issuance: 'Issuance', audit: 'Activity log' };
  const OPEN = ['requested', 'processing', 'draft', 'awaiting', 'approved', 'delete-requested'];
  const HOME = location.origin + location.pathname.replace(/index\.html$/, '');

  const S = {
    me: null, data: null, screen: 'dashboard', sel: new Set(), detail: null, form: null, err: '',
    toast: null, f: { q: '', device: '', period: '', status: '' }, resFilter: 'open',
    remarks: [], remarksKey: '', fv: {}, busy: false
  };
  const $app = document.getElementById('app');
  let msal = null, ids = null;

  // ---------- helpers ----------
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = n => Number(n || 0).toLocaleString('en-MY', { maximumFractionDigits: 3 });
  const pad = n => String(n).padStart(2, '0');
  const dDate = v => { if (!v) return ''; const d = new Date(String(v).length === 10 ? v + 'T00:00:00' : v); return isNaN(d) ? esc(v) : `${d.getDate()} ${MON[d.getMonth()]} ${d.getFullYear()}`; };
  const dTime = v => { const d = new Date(v); return isNaN(d) ? '' : `${dDate(v)}, ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
  const period = (a, b) => (a && b ? `${dDate(a)} to ${dDate(b)}` : '–');
  const short = uid => (uid && uid.length > 12 ? `${uid.slice(0, 6)}…${uid.slice(-4)}` : uid || '–');
  const initials = name => String(name).split(' ').filter(Boolean).map(w => w[0]).slice(0, 2).join('').toUpperCase();
  const label = r => (typeof r === 'string' ? LABEL[r] || r : r.status === 'approved' ? (r.verified ? (r.kind === 'issue' ? 'Approved, submit in registry' : 'Approved, finish in registry') : 'Approved, being checked') : LABEL[r.status] || r.status || 'Unknown');
  const pill = r => { const k = typeof r === 'string' ? r : r.status; return `<span class="pill p-${LABEL[k] || k === 'approved' ? esc(k) : 'other'}">${esc(label(r))}</span>`; };
  const sum = (arr, f) => arr.reduce((t, x) => t + f(x), 0);
  const isTrade = t => /trade/i.test(t || '');
  const attr = o => Object.entries(o).map(([k, v]) => `data-${k}="${esc(v)}"`).join(' ');
  const opt = (v, l, cur) => `<option value="${esc(v)}"${String(v) === String(cur) ? ' selected' : ''}>${esc(l)}</option>`;
  const q = v => String(v).replace(/'/g, "''");
  const enc = s => s.split('/').map(encodeURIComponent).join('/');
  const refOf = (kind, id) => `${kind === 'issue' ? 'IR' : 'RSV'}-${1000 + Number(id)}`;

  function gate(title, text, button) {
    $app.innerHTML = `<div class="gate"><main class="gate-card"><div><div class="brand-name">Saxon Renewables</div><div class="brand-sub">I-REC portal</div></div>
      <h1>${esc(title)}</h1><p>${esc(text)}</p>${button || ''}</main></div>`;
  }

  // ---------- Microsoft Graph, as the signed-in person ----------
  async function token() {
    const account = msal.getActiveAccount();
    try { return (await msal.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e instanceof window.msal.InteractionRequiredAuthError) await msal.acquireTokenRedirect({ scopes: SCOPES, account });
      throw new Error('Your sign-in needs to be refreshed. Signing you in again…');
    }
  }

  async function graph(method, path, body, headers) {
    const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
      method,
      headers: Object.assign({ Authorization: `Bearer ${await token()}` }, body !== undefined ? { 'Content-Type': 'application/json' } : {}, headers || {}),
      body: body !== undefined ? JSON.stringify(body) : undefined
    });
    if (res.status === 204) return null;
    const json = await res.json().catch(() => null);
    if (!res.ok) {
      if (res.status === 412) throw new Error('Someone else changed this at the same moment. It has been reloaded; try again.');
      if (res.status === 403) throw new Error("SharePoint says you don't have permission to do that. Ask an administrator to check your access.");
      if (res.status === 404) { const e = new Error('Not found in SharePoint.'); e.notFound = true; throw e; }
      throw new Error((json && json.error && json.error.message) || `SharePoint request failed (HTTP ${res.status}).`);
    }
    return json;
  }

  async function resolve() {
    if (ids) return ids;
    const u = new URL(C.siteUrl);
    const site = await graph('GET', `/sites/${u.hostname}:${u.pathname}?$select=id`);
    const drives = await graph('GET', `/sites/${site.id}/drives?$select=id,name`);
    const drive = (drives.value || []).find(d => d.name === C.library);
    if (!drive) throw new Error(`The "${C.library}" library was not found, or you don't have access to it.`);
    const lists = await graph('GET', `/sites/${site.id}/lists?$select=id,displayName&$top=500`);
    const listIds = {};
    for (const [k, n] of Object.entries(LIST_NAMES)) {
      const l = (lists.value || []).find(x => x.displayName === n);
      if (!l) throw new Error(`The SharePoint list "${n}" was not found, or you don't have access to it.`);
      listIds[k] = l.id;
    }
    ids = { siteId: site.id, driveId: drive.id, lists: listIds };
    return ids;
  }

  const person = x => (x && x.user ? { id: x.user.id, name: x.user.displayName, email: (x.user.email || '').toLowerCase() } : { id: null, name: 'Unknown', email: '' });
  const flat = i => ({ id: Number(i.id), etag: i.fields['@odata.etag'], f: i.fields, createdAt: i.createdDateTime, createdBy: person(i.createdBy), modifiedBy: person(i.lastModifiedBy) });

  async function items(key, opts) {
    opts = opts || {};
    const { siteId, lists } = await resolve();
    const qs = ['expand=fields'];
    if (opts.filter) qs.push(`$filter=${encodeURIComponent(opts.filter)}`);
    if (opts.orderby) qs.push(`$orderby=${encodeURIComponent(opts.orderby)}`);
    qs.push(`$top=${opts.top || 999}`);
    let url = `/sites/${siteId}/lists/${lists[key]}/items?${qs.join('&')}`;
    const out = [];
    while (url) {
      const page = await graph('GET', url, undefined, { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' });
      (page.value || []).forEach(i => out.push(flat(i)));
      url = opts.top ? null : page['@odata.nextLink'];
    }
    return out;
  }
  async function addItem(key, fields) { const { siteId, lists } = await resolve(); return graph('POST', `/sites/${siteId}/lists/${lists[key]}/items`, { fields }); }
  async function setFields(key, id, fields, etag) { const { siteId, lists } = await resolve(); return graph('PATCH', `/sites/${siteId}/lists/${lists[key]}/items/${id}/fields`, fields, etag ? { 'If-Match': etag } : undefined); }
  async function removeItem(key, id) { const { siteId, lists } = await resolve(); return graph('DELETE', `/sites/${siteId}/lists/${lists[key]}/items/${id}`); }

  async function readFile(name) {
    const { driveId } = await resolve();
    let meta;
    try { meta = await graph('GET', `/drives/${driveId}/root:/${enc(`${C.folder}/${name}`)}?$select=id,@microsoft.graph.downloadUrl`); }
    catch (e) { if (e.notFound) return null; throw e; }
    const res = await fetch(meta['@microsoft.graph.downloadUrl']);
    if (!res.ok) throw new Error(`Could not download ${name} from SharePoint (HTTP ${res.status}).`);
    return res.json();
  }

  // ---------- loading ----------
  const EMPTY = { syncedAt: null, accounts: [], devices: [], issues: [], items: [], transactions: [], beneficiaries: [], reservations: [] };

  async function load() {
    const [snap, status, reqs, ems, acts] = await Promise.all([
      readFile('registry-snapshot.json'), readFile('sync-status.json'),
      items('requests'), items('earmarks'), items('activity', { orderby: 'fields/Created desc', top: 200 })
    ]);
    const snapshot = snap || EMPTY;
    const bens = Object.fromEntries(snapshot.beneficiaries.map(b => [b.uid, b]));
    const devs = Object.fromEntries(snapshot.devices.map(d => [d.code, d]));
    const requests = reqs.filter(r => String(r.f.Kind || '').toLowerCase() !== 'refresh').map(r => {
      const kindRaw = String(r.f.Kind || '').toLowerCase(), kind = kindRaw === 'issue' ? 'issue' : 'reservation';
      let p = {}, res = {};
      try { p = JSON.parse(r.f.Payload || '{}'); } catch (e) { p = {}; }
      try { res = JSON.parse(r.f.Result || '{}'); } catch (e) { res = {}; }
      const s = kind === 'issue'
        ? { deviceCode: p.deviceCode, deviceName: res.deviceName || (devs[p.deviceCode] || {}).name || p.deviceCode, startDate: p.startDate, endDate: p.endDate, volume: Number(p.volume), recipientAccount: p.recipientAccount }
        : { type: kindRaw === 'transfer' ? 'Transfer' : 'Redemption', beneficiaryName: res.beneficiaryName || (bens[p.beneficiaryUid] || {}).name || null,
            destinationAccount: res.destinationAccount || p.destinationAccount || null, sourceAccount: res.sourceAccount || null,
            volume: Number(p.volume), purpose: p.purpose, periodStart: p.periodStart, periodEnd: p.periodEnd, itemUids: p.itemUids || [] };
      return {
        id: r.id, etag: r.etag, kind, ref: refOf(kind, r.id), status: r.f.RequestStatus || 'requested', summary: s,
        registryUid: r.f.RegistryUid || null, error: r.f.ErrorMessage || '',
        verified: r.f.ApprovalVerified === true, approvedByName: r.f.ApprovalVerified === true ? r.f.ApprovedByName : null,
        createdBy: r.createdBy, createdAt: r.createdAt
      };
    }).filter(r => r.status !== 'deleted').sort((a, b) => b.id - a.id);
    const earmarks = {};
    ems.forEach(e => { (earmarks[e.f.Title] = earmarks[e.f.Title] || []).push({ id: e.id, client: e.f.Client, byName: e.createdBy.name }); });
    const refreshQueued = reqs.some(r => String(r.f.Kind || '').toLowerCase() === 'refresh' && r.f.RequestStatus === 'requested');
    S.data = Object.assign({}, snapshot, { status: status || {}, requests, earmarks, refreshQueued });
    S.audit = acts.map(a => ({ at: a.createdAt, name: a.f.PersonName, action: a.f.Title, ref: a.f.Ref, ok: a.f.Ok !== false }));
  }

  function say(msg, isErr) {
    S.toast = { msg, err: !!isErr };
    render();
    clearTimeout(say.t);
    say.t = setTimeout(() => { S.toast = null; render(); }, isErr ? 12000 : 7000);
  }
  async function run(fn) {
    if (S.busy) return;
    S.busy = true;
    document.body.setAttribute('aria-busy', 'true');
    try { await fn(); } catch (e) { say(e.message, true); if (/same moment/.test(e.message)) { await load().catch(() => {}); render(); } }
    finally { S.busy = false; document.body.removeAttribute('aria-busy'); }
  }

  // ---------- derived data ----------
  function derive() {
    const d = S.data, me = S.me;
    const resReqs = d.requests.filter(r => r.kind === 'reservation');
    const issReqs = d.requests.filter(r => r.kind === 'issue');
    const holding = {};
    resReqs.filter(r => OPEN.includes(r.status)).forEach(r => (r.summary.itemUids || []).forEach(u => { (holding[u] = holding[u] || []).push(r); }));
    const items = d.items.map(it => {
      const held = Math.max(0, it.volume - it.available), trade = isTrade(it.accountType), e = (d.earmarks[it.uid] || [])[0];
      const status = !trade ? 'redeemed' : (it.available <= 0 && held > 0 ? 'held' : e ? 'earmarked' : held > 0 ? 'held' : 'available');
      return Object.assign({}, it, { held, trade, earmark: e, status, holds: holding[it.uid] || [] });
    });
    return {
      me, canEdit: ORDER[me.role] >= 2, isApprover: me.role === 'approver',
      items, byUid: Object.fromEntries(items.map(i => [i.uid, i])), resReqs, issReqs,
      beneficiaries: d.beneficiaries.filter(b => b.active).sort((a, b) => a.name.localeCompare(b.name)),
      tradeAccounts: d.accounts.filter(a => isTrade(a.type) && a.active)
    };
  }
  const mine = (D, r) => r.createdBy && r.createdBy.id === D.me.id;

  // ---------- views ----------
  function viewNav(D) {
    const pendRes = D.isApprover ? D.resReqs.filter(r => r.status === 'awaiting' && !mine(D, r)).length : 0;
    const pendIss = D.isApprover ? D.issReqs.filter(r => r.status === 'draft' && !mine(D, r)).length : 0;
    const navItems = [['dashboard', 'Overview', 0], ['inventory', 'Inventory', 0], ['reservations', 'Redemptions and transfers', pendRes], ['issuance', 'Issuance', pendIss], ['audit', 'Activity log', 0]];
    return `<nav class="nav" aria-label="Main">${navItems.map(([k, l, b]) => `<button ${attr({ a: 'go', s: k })} ${S.screen === k ? 'aria-current="page"' : ''}><span>${l}</span>${b ? `<span class="badge">${b}<span class="sr"> waiting</span></span>` : ''}</button>`).join('')}</nav>`;
  }

  function viewSync(D) {
    const st = S.data.status;
    return `<div class="syncbox"><div class="t">Registry data</div>
      <div class="v">${S.data.syncedAt ? 'Synced ' + esc(dTime(S.data.syncedAt)) : 'Not synced yet'}</div>
      ${st.lastRunAt ? `<div class="t">Worker last ran ${esc(dTime(st.lastRunAt))}</div>` : ''}
      ${st.lastRunOk === false ? `<div class="warn">Last run failed: ${esc(st.lastRunMessage)}</div>` : ''}
      ${D.canEdit ? (S.data.refreshQueued ? '<div class="t">Refresh requested. It runs within the hour.</div>' : `<button class="btn-side" data-a="sync">Request a refresh</button>`) : ''}</div>`;
  }

  function viewDashboard(D) {
    const trade = D.items.filter(i => i.trade);
    const free = sum(trade.filter(i => !i.earmark), i => i.available);
    const emv = sum(trade.filter(i => i.earmark), i => i.available);
    const held = sum(trade, i => i.held);
    const clients = new Set(trade.filter(i => i.earmark && i.available > 0).map(i => i.earmark.client)).size;
    const open = D.resReqs.filter(r => OPEN.includes(r.status)).length;
    const awaiting = D.resReqs.filter(r => r.status === 'awaiting');
    const kpis = [
      ['Free to allocate', fmt(free) + ' MWh', 'Not earmarked or in a draft'],
      ['Earmarked for clients', fmt(emv) + ' MWh', `${clients} client${clients === 1 ? '' : 's'}`],
      ['Held in drafts', fmt(held) + ' MWh', `${open} open redemption${open === 1 ? '' : 's'} or transfers`],
      ['Awaiting approval', String(awaiting.length), 'Redemptions or transfers']
    ];
    let title;
    const list = [];
    const row = r => ({ r, desc: r.kind === 'issue' ? `${r.summary.deviceName}, ${period(r.summary.startDate, r.summary.endDate)}, ${fmt(r.summary.volume)} MWh`
      : `${r.summary.beneficiaryName || r.summary.destinationAccount || ''}, ${fmt(r.summary.volume)} MWh, by ${r.createdBy.name}` });
    if (D.isApprover) {
      title = 'Waiting for you';
      awaiting.filter(r => !mine(D, r)).forEach(r => list.push(row(r)));
      D.issReqs.filter(r => r.status === 'draft' && !mine(D, r)).forEach(r => list.push(row(r)));
      [...D.resReqs, ...D.issReqs].filter(r => r.status === 'approved' && r.verified && r.approvedByName === D.me.name).forEach(r => list.push(row(r)));
    } else if (D.canEdit) {
      title = 'Your open requests';
      [...D.resReqs, ...D.issReqs].filter(r => mine(D, r) && (OPEN.includes(r.status) || r.status === 'failed')).forEach(r => list.push(row(r)));
    } else {
      title = 'Awaiting approval';
      awaiting.forEach(r => list.push(row(r)));
    }
    const byDev = {};
    trade.forEach(i => { const n = i.deviceName || 'Unknown device'; byDev[n] = (byDev[n] || 0) + i.available; });
    const devs = Object.entries(byDev).sort((a, b) => b[1] - a[1]);
    const max = Math.max(1, ...devs.map(x => x[1]));
    return `<div class="kpis">${kpis.map(k => `<div class="card kpi"><span class="muted small">${k[0]}</span><span class="v">${k[1]}</span><span class="muted small">${k[2]}</span></div>`).join('')}</div>
    <div class="row">
      <section class="card col-main"><div class="card-h"><h2>${title}</h2><span class="muted small">${list.length} open</span></div>
        ${list.length ? list.map(({ r, desc }) => `<div class="listrow"><div><b>${esc(r.ref)}</b><span class="muted">, ${r.kind === 'issue' ? 'issue request' : esc(r.summary.type.toLowerCase())}</span><div class="muted small">${esc(desc)}</div></div>
          <div class="actions">${pill(r)}<button class="btn sm" ${attr({ a: 'open', k: r.kind === 'issue' ? 'issue' : 'res', id: r.id, s: r.kind === 'issue' ? 'issuance' : 'reservations' })}>Open</button></div></div>`).join('') : `<p class="empty">Nothing waiting on you right now.</p>`}
      </section>
      <section class="card card-b col-side"><h2 class="h2">Available by device</h2>
        ${devs.length ? devs.map(([n, v]) => `<div><div class="kv"><span class="kv-strong">${esc(n)}</span><span>${fmt(v)} MWh</span></div><div class="bar"><span data-pct="${Math.round(v / max * 100)}"></span></div></div>`).join('') : '<p class="muted mt0">No registry data yet. The worker loads it on its first run.</p>'}
      </section>
    </div>
    <section class="card"><div class="card-h"><h2>Recent activity</h2><button class="linkbtn" ${attr({ a: 'go', s: 'audit' })}>See all activity</button></div>
      ${S.audit.slice(0, 6).map(e => `<div class="listrow"><span class="muted when">${esc(dTime(e.at))}</span><b class="whom">${esc(e.name)}</b><span class="what">${esc(e.action)}${e.ref ? ` <span class="muted">(${esc(e.ref)})</span>` : ''}</span></div>`).join('') || '<p class="empty">No activity yet.</p>'}
    </section>`;
  }

  function viewInventory(D) {
    const devices = [...new Set(D.items.map(i => i.deviceName).filter(Boolean))].sort();
    const periods = [...new Map(D.items.filter(i => i.periodStart).map(i => [`${i.periodStart}|${i.periodEnd}`, period(i.periodStart, i.periodEnd)])).entries()].sort((a, b) => b[0].localeCompare(a[0]));
    const qq = S.f.q.trim().toLowerCase();
    const rows = D.items.filter(i => (!S.f.device || i.deviceName === S.f.device) && (!S.f.period || `${i.periodStart}|${i.periodEnd}` === S.f.period) && (!S.f.status || i.status === S.f.status)
      && (!qq || `${i.uid} ${i.deviceName} ${i.accountName} ${i.earmark ? i.earmark.client : ''}`.toLowerCase().includes(qq)));
    const selected = [...S.sel].map(u => D.byUid[u]).filter(Boolean);
    const selAvail = sum(selected, i => i.available);
    const det = S.detail && S.detail.kind === 'item' ? D.byUid[S.detail.id] : null;
    return `<div class="card filters">
        <div class="wide"><label class="lbl" for="f-q">Search</label><input id="f-q" class="fld" type="search" placeholder="Certificate ID, device, account or client" value="${esc(S.f.q)}" data-filter="q"></div>
        <div><label class="lbl" for="f-device">Device</label><select id="f-device" class="fld" data-filter="device">${opt('', 'All devices', S.f.device)}${devices.map(x => opt(x, x, S.f.device)).join('')}</select></div>
        <div><label class="lbl" for="f-period">Production period</label><select id="f-period" class="fld" data-filter="period">${opt('', 'All periods', S.f.period)}${periods.map(([k, l]) => opt(k, l, S.f.period)).join('')}</select></div>
        <div><label class="lbl" for="f-status">Status</label><select id="f-status" class="fld" data-filter="status">${opt('', 'All statuses', S.f.status)}${['available', 'earmarked', 'held', 'redeemed'].map(k => opt(k, LABEL[k], S.f.status)).join('')}</select></div>
      </div>
      ${D.canEdit && selected.length && !S.form ? `<div class="selbar"><b>${selected.length} selected, ${fmt(selAvail)} MWh available</b><div class="actions">
        <button class="btn" ${attr({ a: 'form', f: 'earmark' })}>Earmark for a client</button>
        <button class="btn primary" ${attr({ a: 'form', f: 'res' })}>Request a redemption or transfer</button>
        <button class="btn" data-a="clearsel">Clear selection</button></div></div>` : ''}
      ${S.form === 'earmark' ? viewEarmarkForm(D, selAvail) : ''}
      ${S.form === 'res' ? viewResForm(D, selected, selAvail) : ''}
      <div class="row">
        <section class="card grow"><div class="card-h"><h2>Certificate blocks</h2><span class="muted small">${rows.length} shown, ${fmt(sum(rows, i => i.volume))} MWh</span></div>
          <div class="tablewrap"><table><thead><tr><th><span class="sr">Select</span></th><th>Certificate</th><th>Device</th><th>Period</th><th>Account</th><th class="num">Volume</th><th class="num">Available</th><th>Status</th><th>Earmarked for</th></tr></thead><tbody>
          ${rows.map(i => `<tr class="${det && det.uid === i.uid ? 'current' : S.sel.has(i.uid) ? 'sel' : ''}">
            <td><input type="checkbox" ${attr({ a: 'toggle', id: i.uid })} ${S.sel.has(i.uid) ? 'checked' : ''} ${!D.canEdit || !i.trade || i.available <= 0 ? 'disabled' : ''} aria-label="Select ${esc(short(i.uid))}"></td>
            <td><button class="linkbtn mono" ${attr({ a: 'open', k: 'item', id: i.uid, s: 'inventory' })}>${esc(short(i.uid))}</button></td>
            <td>${esc(i.deviceName || '–')}</td><td>${period(i.periodStart, i.periodEnd)}</td><td>${esc(i.accountName)}</td>
            <td class="num">${fmt(i.volume)}</td><td class="num"><b>${i.trade ? fmt(i.available) : '–'}</b></td><td>${pill(i.status)}</td><td>${esc(i.earmark ? i.earmark.client : '–')}</td></tr>`).join('')}
          </tbody></table>${rows.length ? '' : '<p class="empty">No certificate blocks match these filters. Clear a filter to see more.</p>'}</div>
        </section>
        ${det ? viewItemDetail(D, det) : ''}
      </div>`;
  }

  function viewEarmarkForm(D, selAvail) {
    const names = [...new Set(D.beneficiaries.map(b => b.name))];
    return `<section class="card accent form"><div><h2>Earmark ${fmt(selAvail)} MWh for a client</h2>
      <p class="muted small mt4">An earmark is a note in SharePoint only. The registry is not changed, and the certificates stay available until a draft is created.</p></div>
      <div class="fields"><div class="wide"><label class="lbl" for="em-client">Client (beneficiary)</label><select id="em-client" class="fld" data-f="client">${opt('', 'Choose a client', S.fv.client)}${names.map(n => opt(n, n, S.fv.client)).join('')}</select></div>
      <div class="wide"><label class="lbl" for="em-note">Remark (optional)</label><input id="em-note" class="fld" data-f="note" value="${esc(S.fv.note)}" placeholder="For example: FY2026 renewal, confirm volume first"></div></div>
      ${S.err ? `<p class="error" role="alert">${esc(S.err)}</p>` : ''}
      <div class="actions"><button class="btn primary" data-a="earmark">Save earmark</button><button class="btn" data-a="cancel">Cancel</button></div></section>`;
  }

  function viewResForm(D, selected, selAvail) {
    const fv = S.fv, red = fv.type !== 'Transfer';
    return `<section class="card accent form"><div><h2>Request a draft from ${selected.length} certificate block${selected.length === 1 ? '' : 's'}</h2>
      <p class="muted small mt4">Your request is queued. Within the hour, the worker checks it and creates a Draft in the registry. Nothing moves or is redeemed until it is approved and finalised.</p></div>
      <div class="fields">
        <div><label class="lbl" for="rs-type">Type</label><select id="rs-type" class="fld" data-f="type" data-rerender="1">${opt('Redemption', 'Redemption', fv.type)}${opt('Transfer', 'Transfer', fv.type)}</select></div>
        <div><label class="lbl" for="rs-vol">Volume (MWh)</label><input id="rs-vol" class="fld" inputmode="decimal" data-f="volume" value="${esc(fv.volume)}"><span class="muted small">Up to ${fmt(selAvail)} MWh</span></div>
        ${red ? '' : `<div class="wide"><label class="lbl" for="rs-dest">Receiving account code</label><input id="rs-dest" class="fld mono" data-f="destinationAccount" value="${esc(fv.destinationAccount)}" placeholder="Account code from the buyer"></div>`}
      </div>
      ${red ? `<div class="fields">
        <div class="wide"><label class="lbl" for="rs-ben">Beneficiary</label><select id="rs-ben" class="fld" data-f="beneficiaryUid">${opt('', 'Choose a beneficiary', fv.beneficiaryUid)}${D.beneficiaries.map(b => opt(b.uid, b.location ? `${b.name} (${b.location.slice(0, 40)})` : b.name, fv.beneficiaryUid)).join('')}</select></div>
        <div class="wide"><label class="lbl" for="rs-pur">Reporting purpose</label><input id="rs-pur" class="fld" data-f="purpose" value="${esc(fv.purpose)}"></div>
        <div><label class="lbl" for="rs-start">Consumption from</label><input id="rs-start" class="fld" type="date" data-f="periodStart" value="${esc(fv.periodStart)}"></div>
        <div><label class="lbl" for="rs-end">Consumption to</label><input id="rs-end" class="fld" type="date" data-f="periodEnd" value="${esc(fv.periodEnd)}"></div></div>` : ''}
      ${S.err ? `<p class="error" role="alert">${esc(S.err)}</p>` : ''}
      <div class="actions"><button class="btn primary" data-a="createres">Queue the request</button><button class="btn" data-a="cancel">Cancel</button></div></section>`;
  }

  function viewRemarks(D, type, id) {
    return `<div class="panel-sec"><h3>Remarks</h3>
      ${S.remarksKey !== `${type}:${id}` ? '<p class="muted small mt0">Loading…</p>' : S.remarks.length ? S.remarks.map(m => `<div class="remark"><span class="m"><b>${esc(m.byName)}</b>, ${esc(dTime(m.at))}</span><span>${esc(m.text)}</span></div>`).join('') : '<p class="muted small mt0">No remarks yet.</p>'}
      ${D.canEdit ? `<label class="lbl flush" for="rm-text">Add a remark</label><textarea id="rm-text" class="fld" rows="3" data-f="remark">${esc(S.fv.remark)}</textarea>
      <div><button class="btn primary sm" ${attr({ a: 'remark', t: type, id })}>Add remark</button></div>` : ''}</div>`;
  }

  function panelHead(kicker, title, pillHtml) {
    return `<div class="panel-sec panel-head"><div class="panel-title"><span class="muted small">${esc(kicker)}</span><span class="t">${title}</span><span>${pillHtml}</span></div>
      <button class="btn sm" data-a="close">Close</button></div>`;
  }

  function viewItemDetail(D, i) {
    return `<aside class="card aside">${panelHead('Certificate block', `<span class="mono">${esc(i.uid)}</span>`, pill(i.status))}
      <div class="panel-sec"><h3>Where it came from</h3>
        <div class="kv"><span>Device</span><span>${esc(i.deviceName || '–')}${i.fuel ? ', ' + esc(i.fuel) : ''}</span></div>
        <div class="kv"><span>Device code</span><span class="mono">${esc(i.deviceCode || '–')}</span></div>
        <div class="kv"><span>Issue request</span><span class="mono">${esc(short(i.issueUid))}</span></div>
        <div class="kv"><span>Asset</span><span class="mono">${esc(short(i.assetUid))}</span></div>
        <div class="kv"><span>Production period</span><span>${period(i.periodStart, i.periodEnd)}</span></div>
        <div class="kv"><span>Account</span><span>${esc(i.accountName)}</span></div>
        <div class="kv"><span>Volume</span><span>${fmt(i.volume)} MWh</span></div>
        <div class="kv"><span>Available</span><span><b>${fmt(i.available)} MWh</b></span></div></div>
      ${i.earmark ? `<div class="panel-sec"><h3>Earmark</h3><p class="mt0">Held for <b>${esc(i.earmark.client)}</b> by ${esc(i.earmark.byName)}.</p>
        ${D.canEdit ? `<div><button class="btn danger sm" ${attr({ a: 'unearmark', id: i.uid })}>Remove earmark</button></div>` : ''}</div>` : ''}
      ${i.holds.length ? `<div class="panel-sec"><h3>In requests</h3>${i.holds.map(h => `<div class="kv"><button class="linkbtn" ${attr({ a: 'open', k: 'res', id: h.id, s: 'reservations' })}>${esc(h.ref)}</button><span>${esc(label(h))}</span></div>`).join('')}</div>` : ''}
      ${viewRemarks(D, 'item', i.uid)}</aside>`;
  }

  function viewReservations(D) {
    const F = { open: ['Open', r => OPEN.includes(r.status)], failed: ['Failed', r => r.status === 'failed'], awaiting: ['Awaiting approval', r => r.status === 'awaiting'], approved: ['Approved', r => r.status === 'approved'], completed: ['Completed', r => r.status === 'completed'], all: ['All', () => true] };
    const rows = D.resReqs.filter(F[S.resFilter][1]);
    const det = S.detail && S.detail.kind === 'res' ? D.resReqs.find(r => r.id === Number(S.detail.id)) : null;
    const tx = (S.data.transactions || []).slice().sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp))).slice(0, 50);
    return `<div class="row between">
        <div class="actions" role="group" aria-label="Filter by status">${Object.entries(F).map(([k, [l, f]]) => `<button class="btn sm ${S.resFilter === k ? 'on' : ''}" aria-pressed="${S.resFilter === k}" ${attr({ a: 'resfilter', v: k })}>${l} (${D.resReqs.filter(f).length})</button>`).join('')}</div>
        ${D.canEdit ? `<button class="btn primary" ${attr({ a: 'go', s: 'inventory' })}>Pick certificates for a new request</button>` : ''}</div>
      <div class="row">
        <section class="card grow"><div class="card-h"><h2>Requests from the portal</h2></div><div class="tablewrap"><table>
          <thead><tr><th>Reference</th><th>Type</th><th>Beneficiary or receiving account</th><th class="num">Volume</th><th>Status</th><th>Requested by</th><th>Requested</th></tr></thead><tbody>
          ${rows.map(r => `<tr class="${det && det.id === r.id ? 'current' : ''}"><td><button class="linkbtn" ${attr({ a: 'open', k: 'res', id: r.id, s: 'reservations' })}>${esc(r.ref)}</button></td><td>${esc(r.summary.type)}</td>
            <td class="wrap">${esc(r.summary.beneficiaryName || r.summary.destinationAccount || '–')}</td><td class="num">${fmt(r.summary.volume)}</td><td>${pill(r)}</td><td>${esc(r.createdBy.name)}</td><td>${dDate(r.createdAt)}</td></tr>`).join('')}
          </tbody></table>${rows.length ? '' : '<p class="empty">No requests with this status.</p>'}</div></section>
        ${det ? viewResDetail(D, det) : ''}
      </div>
      <section class="card"><div class="card-h"><h2>Recent transactions in the registry</h2><span class="muted small">Latest ${tx.length}, from the last sync</span></div><div class="tablewrap"><table>
        <thead><tr><th>Date</th><th>Type</th><th>From</th><th>To</th><th>Beneficiary</th><th class="num">Volume</th></tr></thead><tbody>
        ${tx.map(t => `<tr><td>${dDate(t.timestamp)}</td><td>${esc(t.type)}</td><td class="mono">${esc(t.source)}</td><td class="mono">${esc(t.destination)}</td><td class="wrap">${esc(t.beneficiary || '–')}</td><td class="num">${fmt(t.volume)}</td></tr>`).join('')}
        </tbody></table>${tx.length ? '' : '<p class="empty">No transactions synced yet.</p>'}</div></section>`;
  }

  function requestActions(D, r) {
    const own = mine(D, r);
    const btn = (act, text, cls) => `<button class="btn ${cls || ''}" ${attr({ a: 'reqact', act, id: r.id })}>${text}</button>`;
    let html = r.error ? `<p class="note ${r.status === 'failed' ? 'err' : 'warn'}">${esc(r.error)}</p>` : '';
    if (r.status === 'requested') html += `<p class="note">Queued. The worker checks it and creates the Draft in the registry within the hour.</p>${D.canEdit ? `<div>${btn('cancel', 'Cancel request', 'danger sm')}</div>` : ''}`;
    else if (r.status === 'processing') html += `<p class="note">The worker is creating this in the registry right now.</p>`;
    else if (r.status === 'failed') html += D.canEdit ? `<div>${btn('cancel', 'Remove from the list', 'sm')}</div>` : '';
    else if (r.status === 'draft' && r.kind === 'issue') {
      if (D.isApprover && !own) html += `<p class="note">Approving records your approval. You then submit it to the Issuer in the registry.</p><div>${btn('approve', 'Approve for submission', 'primary')}</div>`;
      else if (D.isApprover) html += `<p class="note warn">You requested this, so another approver has to approve it.</p>`;
      else html += `<p class="note">An approver approves this before it is submitted to the Issuer.</p>`;
      if (D.canEdit) html += `<p class="muted small mt4">Issue requests cannot be deleted in the registry, only withdrawn there.</p><div>${btn('withdraw', 'Mark as withdrawn', 'danger sm')}</div>`;
    }
    else if (r.status === 'draft' && D.canEdit) html += `<div class="actions">${btn('request-approval', 'Send for approval', 'primary')}${btn('delete', 'Delete draft', 'danger')}</div>`;
    else if (r.status === 'draft') html += `<p class="note">Draft in the registry. Operators can send it for approval or delete it.</p>`;
    else if (r.status === 'awaiting' && D.isApprover && !own) html += `<p class="note">Approving records your approval. You then submit and approve it in the registry, where it becomes final.</p><div class="actions">${btn('approve', 'Approve', 'primary')}${btn('send-back', 'Send back to draft')}</div>`;
    else if (r.status === 'awaiting' && D.isApprover) html += `<p class="note warn">You requested this, so another approver has to approve it.</p>`;
    else if (r.status === 'awaiting') html += `<p class="note">Waiting for an approver. The certificates stay held until then.</p>${D.canEdit ? `<div>${btn('delete', 'Delete draft', 'danger sm')}</div>` : ''}`;
    else if (r.status === 'approved' && !r.verified) html += `<p class="note">Approved. The worker confirms the approver within the hour.</p>`;
    else if (r.status === 'approved') html += `<p class="note warn">Approved by ${esc(r.approvedByName)}. ${r.kind === 'issue' ? 'Submit' : 'Submit and approve'} <span class="mono">${esc(short(r.registryUid))}</span> in the registry. The next sync updates its status.</p>`;
    else if (r.status === 'delete-requested') html += `<p class="note">Deletion queued. The worker removes it from the registry within the hour.</p>`;
    else if (r.status === 'completed') html += `<p class="note">Final in the registry. The redemption statement is available from the registry.</p>`;
    else if (r.status === 'submitted') html += `<p class="note">With the Issuer for review. Certificates appear in the inventory once it is approved.</p>`;
    else if (r.status === 'issued') html += `<p class="note">Issued. The certificates are in the inventory.</p>`;
    else if (r.status === 'rejected') html += `<p class="note warn">Rejected by the Issuer. Check the Issuer's notes in the registry.</p>`;
    return html || '<p class="muted small mt0">No actions for this status.</p>';
  }

  function viewResDetail(D, r) {
    const s = r.summary;
    return `<aside class="card aside">${panelHead(s.type, esc(r.ref), pill(r))}
      <div class="panel-sec">
        <div class="kv"><span>${s.type === 'Redemption' ? 'Beneficiary' : 'Receiving account'}</span><span>${esc(s.beneficiaryName || s.destinationAccount || '–')}</span></div>
        <div class="kv"><span>Volume</span><span><b>${fmt(s.volume)} MWh</b></span></div>
        ${s.sourceAccount ? `<div class="kv"><span>From account</span><span class="mono">${esc(s.sourceAccount)}</span></div>` : ''}
        ${s.type === 'Redemption' ? `<div class="kv"><span>Purpose</span><span>${esc(s.purpose)}</span></div><div class="kv"><span>Consumption period</span><span>${period(s.periodStart, s.periodEnd)}</span></div>` : ''}
        <div class="kv"><span>Requested by</span><span>${esc(r.createdBy.name)}, ${dDate(r.createdAt)}</span></div>
        ${r.approvedByName ? `<div class="kv"><span>Approved by</span><span>${esc(r.approvedByName)}</span></div>` : ''}
        <div class="kv"><span>Registry reservation</span><span class="mono">${esc(r.registryUid || 'Not created yet')}</span></div></div>
      <div class="panel-sec"><h3>Certificate blocks selected</h3>${(s.itemUids || []).map(u => { const it = D.byUid[u]; return `<div class="kv"><span><span class="mono">${esc(short(u))}</span>${it ? ', ' + esc(it.deviceName) : ''}</span><span>${it ? fmt(it.available) + ' MWh available' : ''}</span></div>`; }).join('') || '<p class="muted small mt0">None</p>'}</div>
      <div class="panel-sec">${requestActions(D, r)}</div>
      ${viewRemarks(D, 'request', r.id)}</aside>`;
  }

  function viewIssuance(D) {
    const devices = S.data.devices.filter(d => /approved/i.test(d.status));
    const det = S.detail && S.detail.kind === 'issue' ? D.issReqs.find(r => r.id === Number(S.detail.id)) : null;
    const reg = S.data.issues.slice().sort((a, b) => String(b.periodEnd).localeCompare(String(a.periodEnd)));
    const lastIssued = {};
    reg.forEach(i => { if (/approved|issued/i.test(i.status) && !lastIssued[i.deviceCode]) lastIssued[i.deviceCode] = period(i.periodStart, i.periodEnd); });
    return `<div class="devices">${devices.map(d => `<div class="card card-b device">
        <div class="kv"><span class="device-name"><b>${esc(d.name)}</b><span class="muted small">${esc(d.fuel)}${d.capacity ? ', ' + fmt(d.capacity) + ' MW' : ''}, <span class="mono">${esc(d.code)}</span></span></span><span><span class="pill p-available">Approved</span></span></div>
        <span class="small">Last issued: ${lastIssued[d.code] || 'None yet'}</span>
        ${D.canEdit ? `<div><button class="btn sm" ${attr({ a: 'issueform', code: d.code })}>Request an issue</button></div>` : ''}</div>`).join('') || '<p class="muted">No approved devices in the last sync.</p>'}</div>
      ${S.form === 'issue' ? viewIssueForm(D) : ''}
      <div class="row">
        <section class="card grow"><div class="card-h"><h2>Issue requests from the portal</h2></div><div class="tablewrap"><table>
          <thead><tr><th>Reference</th><th>Device</th><th>Period</th><th class="num">Volume</th><th>Status</th><th>Requested by</th><th>Requested</th></tr></thead><tbody>
          ${D.issReqs.map(r => `<tr class="${det && det.id === r.id ? 'current' : ''}"><td><button class="linkbtn" ${attr({ a: 'open', k: 'issue', id: r.id, s: 'issuance' })}>${esc(r.ref)}</button></td><td>${esc(r.summary.deviceName)}</td><td>${period(r.summary.startDate, r.summary.endDate)}</td>
            <td class="num">${fmt(r.summary.volume)}</td><td>${pill(r)}</td><td>${esc(r.createdBy.name)}</td><td>${dDate(r.createdAt)}</td></tr>`).join('')}
          </tbody></table>${D.issReqs.length ? '' : '<p class="empty">No issue requests made in the portal yet.</p>'}</div></section>
        ${det ? viewIssueDetail(D, det) : ''}
      </div>
      <section class="card"><div class="card-h"><h2>All issue requests in the registry</h2><span class="muted small">From the last sync</span></div><div class="tablewrap"><table>
        <thead><tr><th>Device</th><th>Period</th><th class="num">Production</th><th class="num">Issued</th><th>Registry status</th></tr></thead><tbody>
        ${reg.slice(0, 100).map(i => `<tr><td>${esc(i.deviceName)}</td><td>${period(i.periodStart, i.periodEnd)}</td><td class="num">${fmt(i.productionVolume)}</td><td class="num">${fmt(i.issuedVolume)}</td><td>${esc(i.status)}</td></tr>`).join('')}
        </tbody></table>${reg.length ? '' : '<p class="empty">No issue requests synced yet.</p>'}</div></section>`;
  }

  function viewIssueForm(D) {
    const dev = S.data.devices.find(d => d.code === S.fv.deviceCode);
    return `<section class="card accent form"><div><h2>Request an issue for ${esc(dev ? dev.name : '')}</h2>
      <p class="muted small mt4">Queued for the worker, which creates a Draft issue request in the registry within the hour. The Issuer only sees it after it is approved and submitted.</p></div>
      <div class="fields">
        <div><label class="lbl" for="is-start">Production from</label><input id="is-start" class="fld" type="date" data-f="startDate" value="${esc(S.fv.startDate)}"></div>
        <div><label class="lbl" for="is-end">Production to</label><input id="is-end" class="fld" type="date" data-f="endDate" value="${esc(S.fv.endDate)}"></div>
        <div><label class="lbl" for="is-vol">Production (MWh)</label><input id="is-vol" class="fld" inputmode="decimal" data-f="volume" value="${esc(S.fv.volume)}"></div>
        <div class="wide"><label class="lbl" for="is-acc">Deposit into</label><select id="is-acc" class="fld" data-f="recipientAccount">${D.tradeAccounts.map(a => opt(a.code, `${a.name} (${a.code})`, S.fv.recipientAccount)).join('')}</select></div></div>
      <p class="muted small mt0">Attach meter readings and invoices to the request in the registry. They are required by the Issuer.</p>
      ${S.err ? `<p class="error" role="alert">${esc(S.err)}</p>` : ''}
      <div class="actions"><button class="btn primary" data-a="createissue">Queue the request</button><button class="btn" data-a="cancel">Cancel</button></div></section>`;
  }

  function viewIssueDetail(D, r) {
    const s = r.summary;
    return `<aside class="card aside">${panelHead('Issue request', esc(r.ref), pill(r))}
      <div class="panel-sec">
        <div class="kv"><span>Device</span><span>${esc(s.deviceName)}</span></div>
        <div class="kv"><span>Production period</span><span>${period(s.startDate, s.endDate)}</span></div>
        <div class="kv"><span>Volume</span><span><b>${fmt(s.volume)} MWh</b></span></div>
        <div class="kv"><span>Deposit into</span><span class="mono">${esc(s.recipientAccount)}</span></div>
        <div class="kv"><span>Requested by</span><span>${esc(r.createdBy.name)}, ${dDate(r.createdAt)}</span></div>
        ${r.approvedByName ? `<div class="kv"><span>Approved by</span><span>${esc(r.approvedByName)}</span></div>` : ''}
        <div class="kv"><span>Registry issue request</span><span class="mono">${esc(r.registryUid || 'Not created yet')}</span></div></div>
      <div class="panel-sec">${requestActions(D, r)}</div>
      ${viewRemarks(D, 'request', r.id)}</aside>`;
  }

  function viewAudit() {
    return `<section class="card"><div class="card-h stack"><h2>What the worker did in the registry</h2>
      <span class="muted small">Every registry action is listed with the person it was done for. SharePoint's own version history shows who made each request, earmark and remark.</span></div>
      <div class="tablewrap"><table><thead><tr><th>When</th><th>Who</th><th>What</th><th>Reference</th><th>Result</th></tr></thead><tbody>
      ${S.audit.map(e => `<tr><td>${esc(dTime(e.at))}</td><td><b>${esc(e.name)}</b></td><td class="wrap">${esc(e.action)}</td><td class="mono">${esc(e.ref || '–')}</td><td>${e.ok ? 'Done' : '<span class="pill p-failed">Failed</span>'}</td></tr>`).join('')}
      </tbody></table>${S.audit.length ? '' : '<p class="empty">No activity yet.</p>'}</div></section>`;
  }

  // ---------- render ----------
  function render() {
    if (!S.data) return;
    const D = derive();
    const active = document.activeElement;
    const focusId = active && active.id;
    const caret = active && typeof active.selectionStart === 'number' ? [active.selectionStart, active.selectionEnd] : null;
    const screens = { dashboard: viewDashboard, inventory: viewInventory, reservations: viewReservations, issuance: viewIssuance, audit: viewAudit };
    $app.innerHTML = `<div class="layout">
      <aside class="side"><div class="brand"><div class="brand-name">Saxon Renewables</div><div class="brand-sub">I-REC portal</div></div>${viewNav(D)}${viewSync(D)}</aside>
      <div class="main">
        <header class="top"><h1>${TITLES[S.screen]}</h1><div class="who"><div class="avatar" aria-hidden="true">${esc(initials(D.me.name))}</div>
          <div class="who-text"><span class="n">${esc(D.me.name)}</span><span class="r">${ROLE_LABEL[D.me.role]}</span></div>
          <button class="btn sm" data-a="reload">Reload</button><button class="btn sm" data-a="signout">Sign out</button></div></header>
        ${S.toast ? `<div class="toast ${S.toast.err ? 'err' : ''}" role="${S.toast.err ? 'alert' : 'status'}"><span>${esc(S.toast.msg)}</span><button class="linkbtn light" data-a="toast">Dismiss</button></div>` : ''}
        <main class="content">${screens[S.screen](D)}</main>
      </div></div>`;
    document.querySelectorAll('[data-pct]').forEach(el => { el.style.width = el.dataset.pct + '%'; });
    if (focusId) {
      const el = document.getElementById(focusId);
      if (el) { el.focus(); if (caret && typeof el.setSelectionRange === 'function') { try { el.setSelectionRange(caret[0], caret[1]); } catch (e) { /* not a text field */ } } }
    }
  }

  async function loadRemarks(type, id) {
    const key = `${type}:${id}`;
    S.remarksKey = '';
    try {
      const list = await items('remarks', { filter: `fields/TargetId eq '${q(id)}'` });
      const mapped = list.filter(m => m.f.Title === type).sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
        .map(m => ({ text: m.f.RemarkText, byName: m.createdBy.name, at: m.createdAt }));
      const cur = S.detail && `${S.detail.kind === 'item' ? 'item' : 'request'}:${S.detail.id}`;
      if (cur === key) { S.remarks = mapped; S.remarksKey = key; render(); }
    } catch (e) { say(e.message, true); }
  }

  function openDetail(kind, id, screen) {
    S.screen = screen; S.detail = { kind, id }; S.form = null; S.err = ''; S.fv.remark = '';
    render();
    loadRemarks(kind === 'item' ? 'item' : 'request', id);
  }
  async function refreshAll(msg) { await load(); render(); if (msg) say(msg); }

  // ---------- events ----------
  document.addEventListener('click', e => {
    const el = e.target.closest('[data-a]');
    if (!el || el.disabled) return;
    const a = el.dataset.a;
    if (a === 'signin') { msal.loginRedirect({ scopes: SCOPES }); return; }
    if (a === 'signout') { msal.logoutRedirect({ account: msal.getActiveAccount(), postLogoutRedirectUri: HOME }); return; }
    if (a === 'reload') { if (S.data) run(() => refreshAll('Reloaded.')); else location.reload(); return; }
    if (a === 'toggle') return;
    if (!S.data) return;
    const D = derive();
    if (a === 'go') { S.screen = el.dataset.s; S.detail = null; S.form = null; S.err = ''; render(); window.scrollTo(0, 0); return; }
    if (a === 'open') { openDetail(el.dataset.k, el.dataset.id, el.dataset.s); return; }
    if (a === 'close') { S.detail = null; render(); return; }
    if (a === 'toast') { S.toast = null; render(); return; }
    if (a === 'clearsel') { S.sel.clear(); render(); return; }
    if (a === 'cancel') { S.form = null; S.err = ''; render(); return; }
    if (a === 'resfilter') { S.resFilter = el.dataset.v; render(); return; }
    if (a === 'form') {
      const selAvail = sum([...S.sel].map(u => D.byUid[u]).filter(Boolean), i => i.available), y = new Date().getFullYear();
      S.form = el.dataset.f; S.err = ''; S.detail = null;
      S.fv = el.dataset.f === 'earmark' ? { client: '', note: '' } : { type: 'Redemption', volume: String(selAvail), beneficiaryUid: '', purpose: `Scope 2 reporting ${y}`, periodStart: `${y}-01-01`, periodEnd: '', destinationAccount: '' };
      render(); return;
    }
    if (a === 'issueform') {
      S.form = 'issue'; S.err = ''; S.detail = null;
      S.fv = { deviceCode: el.dataset.code, startDate: '', endDate: '', volume: '', recipientAccount: (D.tradeAccounts[0] || {}).code || '' };
      render(); return;
    }
    if (a === 'sync') run(async () => {
      await addItem('requests', { Title: 'Refresh', Kind: 'refresh', RequestStatus: 'requested' });
      await refreshAll('Refresh requested. The worker reloads registry data within the hour.');
    });
    if (a === 'earmark') run(async () => {
      if (!S.fv.client) { S.err = 'Choose a client first.'; render(); return; }
      const client = S.fv.client, note = String(S.fv.note || '').trim(), uids = [...S.sel];
      for (const uid of uids) {
        for (const old of S.data.earmarks[uid] || []) await removeItem('earmarks', old.id);
        await addItem('earmarks', { Title: uid, Client: client });
        if (note) await addItem('remarks', { Title: 'item', TargetId: uid, RemarkText: note });
      }
      S.sel.clear(); S.form = null; await refreshAll(`Earmarked for ${client}. Nothing changed in the registry.`);
    });
    if (a === 'unearmark') run(async () => {
      for (const old of S.data.earmarks[el.dataset.id] || []) await removeItem('earmarks', old.id);
      await refreshAll('Earmark removed.');
    });
    if (a === 'createres') run(async () => {
      const fv = S.fv, vol = Number(String(fv.volume).replace(/,/g, ''));
      const selAvail = sum([...S.sel].map(u => D.byUid[u]).filter(Boolean), i => i.available);
      if (!(vol > 0)) { S.err = 'Enter a volume above zero.'; render(); return; }
      if (vol > selAvail) { S.err = `Only ${fmt(selAvail)} MWh is available in the selected blocks.`; render(); return; }
      if (fv.type === 'Redemption' && !fv.beneficiaryUid) { S.err = 'Choose a beneficiary.'; render(); return; }
      if (fv.type === 'Redemption' && (!fv.periodStart || !fv.periodEnd || fv.periodEnd < fv.periodStart)) { S.err = 'Enter a consumption period that ends after it starts.'; render(); return; }
      if (fv.type === 'Transfer' && !String(fv.destinationAccount).trim()) { S.err = 'Enter the receiving account code.'; render(); return; }
      const payload = fv.type === 'Transfer' ? { itemUids: [...S.sel], volume: vol, destinationAccount: String(fv.destinationAccount).trim() }
        : { itemUids: [...S.sel], volume: vol, beneficiaryUid: fv.beneficiaryUid, purpose: fv.purpose, periodStart: fv.periodStart, periodEnd: fv.periodEnd };
      const created = await addItem('requests', { Title: 'New request', Kind: fv.type.toLowerCase(), RequestStatus: 'requested', Payload: JSON.stringify(payload) });
      S.sel.clear(); S.form = null; await load(); S.resFilter = 'open';
      openDetail('res', Number(created.id), 'reservations'); say(`${refOf('reservation', created.id)} queued. The worker creates the Draft in the registry within the hour.`);
    });
    if (a === 'createissue') run(async () => {
      const fv = S.fv, vol = Number(String(fv.volume).replace(/,/g, ''));
      if (!fv.startDate || !fv.endDate || fv.endDate < fv.startDate) { S.err = 'Enter a production period that ends after it starts.'; render(); return; }
      if (!(vol > 0)) { S.err = 'Enter the production volume in MWh.'; render(); return; }
      if (!fv.recipientAccount) { S.err = 'Choose the account to deposit into.'; render(); return; }
      const created = await addItem('requests', { Title: 'New request', Kind: 'issue', RequestStatus: 'requested', Payload: JSON.stringify({ deviceCode: fv.deviceCode, startDate: fv.startDate, endDate: fv.endDate, volume: vol, recipientAccount: fv.recipientAccount }) });
      S.form = null; await load(); openDetail('issue', Number(created.id), 'issuance'); say(`${refOf('issue', created.id)} queued. The worker creates the Draft within the hour.`);
    });
    if (a === 'reqact') run(async () => {
      const r = S.data.requests.find(x => x.id === Number(el.dataset.id)), act = el.dataset.act;
      if (!r) return;
      const next = { 'request-approval': 'awaiting', approve: 'approved', 'send-back': 'draft', delete: 'delete-requested', withdraw: 'withdrawn' }[act];
      const asks = { approve: `Approve ${r.ref}? You then finalise it in the registry.`, delete: `Delete ${r.ref}? The worker removes the draft from the registry within the hour.`, withdraw: `Mark ${r.ref} as withdrawn? You also need to withdraw it in the registry.`, cancel: `Remove ${r.ref} from the list?` };
      if (asks[act] && !confirm(asks[act])) return;
      if (act === 'cancel') { await removeItem('requests', r.id); S.detail = null; await refreshAll(`${r.ref} removed.`); return; }
      await setFields('requests', r.id, { RequestStatus: next, ErrorMessage: '' }, r.etag);
      const msg = { approved: `${r.ref} approved. The worker confirms it within the hour; then finalise it in the registry.`, awaiting: `${r.ref} sent for approval.`, draft: `${r.ref} sent back to draft.`, 'delete-requested': `${r.ref} will be deleted from the registry within the hour.`, withdrawn: `${r.ref} marked as withdrawn.` }[next];
      await refreshAll(msg);
    });
    if (a === 'remark') run(async () => {
      const text = String(S.fv.remark || '').trim();
      if (!text) { say('Write a remark first.', true); return; }
      if (text.length > 2000) { say('Remarks are limited to 2,000 characters.', true); return; }
      await addItem('remarks', { Title: el.dataset.t, TargetId: String(el.dataset.id), RemarkText: text });
      S.fv.remark = ''; await loadRemarks(el.dataset.t, el.dataset.id);
    });
  });

  let filterTimer;
  document.addEventListener('input', e => {
    const t = e.target;
    if (t.dataset.f) { S.fv[t.dataset.f] = t.value; if (S.err) S.err = ''; return; }
    if (t.dataset.filter === 'q') { S.f.q = t.value; clearTimeout(filterTimer); filterTimer = setTimeout(render, 200); }
  });
  document.addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.a === 'toggle') { if (t.checked) S.sel.add(t.dataset.id); else S.sel.delete(t.dataset.id); render(); return; }
    if (t.dataset.f) { S.fv[t.dataset.f] = t.value; if (t.dataset.rerender) render(); return; }
    if (t.dataset.filter && t.dataset.filter !== 'q') { S.f[t.dataset.filter] = t.value; render(); }
  });

  // ---------- start ----------
  (async function boot() {
    try {
      if (!C.tenantId || /^</.test(C.tenantId) || !C.clientId || /^</.test(C.clientId) || !C.siteUrl) {
        gate('The portal is not set up yet', 'Fill in tenantId and clientId in config.js. See the setup guide.');
        return;
      }
      if (!window.msal) throw new Error('The Microsoft sign-in library did not load.');
      msal = new window.msal.PublicClientApplication({
        auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenantId}`, redirectUri: HOME, postLogoutRedirectUri: HOME },
        cache: { cacheLocation: 'sessionStorage' }
      });
      await msal.initialize();
      const result = await msal.handleRedirectPromise();
      const account = (result && result.account) || msal.getAllAccounts()[0];
      if (!account) {
        gate('Sign in to manage your certificates', 'Use your Saxon Renewables work account. You will be asked to approve the sign-in in Microsoft Authenticator.',
          '<button class="btn primary lg" data-a="signin">Sign in with Microsoft</button><p class="small">Only staff assigned to this app can sign in.</p>');
        return;
      }
      msal.setActiveAccount(account);
      const claims = account.idTokenClaims || {};
      const roles = (claims.roles || []).map(r => String(r).toLowerCase()).filter(r => PORTAL_ROLES.includes(r)).sort((x, y) => ORDER[y] - ORDER[x]);
      if (!roles.length) {
        gate("You don't have access yet", "You're signed in, but no portal role has been assigned to your account. Ask Freda Tan to assign you Viewer, Operator or Approver.",
          '<button class="btn" data-a="signout">Sign out</button>');
        return;
      }
      S.me = { id: claims.oid, name: claims.name || account.username, email: String(claims.preferred_username || account.username || '').toLowerCase(), role: roles[0] };
      await load();
      render();
    } catch (e) {
      gate('The portal could not load', e.message, '<button class="btn primary" data-a="reload">Try again</button>');
    }
  })();
})();
