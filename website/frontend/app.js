'use strict';
// Saxon I-REC portal: browser app. Plain JavaScript, no build step, no backend.
// Staff sign in with Microsoft (Authenticator). Everything is stored in SharePoint as
// the signed-in person, and every action is written to the IREC Activity list, where
// SharePoint itself records who added each entry.
// Registry data comes from a file made by the private exporter on GitHub and imported
// here. The portal never contacts the I-REC registry.

(function () {
  const C = window.IREC_PORTAL_CONFIG || {};
  const GRAPH = (C.graphBase || 'https://graph.microsoft.com/v1.0').replace(/\/+$/, '');
  const SCOPES = ['User.Read', 'Sites.ReadWrite.All'];
  const PORTAL_ROLES = ['viewer', 'operator', 'approver'];
  const ORDER = { viewer: 1, operator: 2, approver: 2 };
  const ROLE_LABEL = { viewer: 'Viewer', operator: 'Operator', approver: 'Operator' };
  const LIST_NAMES = { earmarks: 'IREC Earmarks', remarks: 'IREC Remarks', activity: 'IREC Activity' };
  const OPTIONAL_LISTS = { requests: 'IREC Requests' }; // draft requests work once this list exists
  const FORMAT = 'saxon-irec-snapshot';
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const LABEL = { available: 'Available', earmarked: 'Earmarked', queued: 'Request queued', held: 'In a registry draft', redeemed: 'Redeemed' };
  const REQ_LABEL = { requested: 'Queued for the registry', processing: 'Being created', draft: 'Draft in the registry', submitted: 'Submitted in the registry', completed: 'Completed',
    failed: 'Failed', 'delete-requested': 'Deletion queued', deleted: 'Deleted', withdrawn: 'Withdrawn', issued: 'Issued', rejected: 'Rejected' };
  const OPEN_REQ = ['requested', 'processing', 'draft', 'submitted', 'delete-requested'];
  const refOf = (kind, id) => `${kind === 'issue' ? 'IR' : 'RSV'}-${1000 + Number(id)}`;
  const TITLES = { dashboard: 'Overview', inventory: 'Inventory', reservations: 'Redemptions and transfers', issuance: 'Issuance', audit: 'Activity log', import: 'Import registry data' };
  const HOME = location.origin + location.pathname.replace(/index\.html$/, '');

  const S = { me: null, data: null, audit: [], screen: 'dashboard', sel: new Set(), detail: null, form: null, err: '', toast: null,
    f: { q: '', device: '', period: '', status: '' }, remarks: [], remarksKey: '', fv: {}, busy: false, pending: null };
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
  const pill = k => `<span class="pill p-${LABEL[k] ? esc(k) : 'other'}">${esc(LABEL[k] || k)}</span>`;
  const reqPill = r => `<span class="pill p-${REQ_LABEL[r.status] ? esc(r.status) : 'other'}">${esc(r.kind === 'issue' && r.status === 'submitted' ? 'With the Issuer' : REQ_LABEL[r.status] || r.status || 'Unknown')}</span>`;
  const regPill = s => { const k = String(s || '').toLowerCase(); const cls = k === 'draft' ? 'draft' : k === 'submitted' ? 'submitted' : k === 'approved' ? 'completed' : k === 'rejected' ? 'rejected' : 'other'; return `<span class="pill p-${cls}">${esc(s || 'Unknown')}</span>`; };
  const sum = (arr, f) => arr.reduce((t, x) => t + f(x), 0);
  const isTrade = t => /trade/i.test(t || '');
  const attr = o => Object.entries(o).map(([k, v]) => `data-${k}="${esc(v)}"`).join(' ');
  const opt = (v, l, cur) => `<option value="${esc(v)}"${String(v) === String(cur) ? ' selected' : ''}>${esc(l)}</option>`;
  const q = v => String(v).replace(/'/g, "''");
  const enc = s => s.split('/').map(encodeURIComponent).join('/');

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

  async function graph(method, path, body, headers, raw) {
    const h = Object.assign({ Authorization: `Bearer ${await token()}` }, headers || {});
    if (body !== undefined && !raw) h['Content-Type'] = 'application/json';
    const res = await fetch(path.startsWith('http') ? path : GRAPH + path, { method, headers: h, body: body === undefined ? undefined : raw ? body : JSON.stringify(body) });
    if (res.status === 204) return null;
    const json = await res.json().catch(() => null);
    if (!res.ok) {
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
    for (const [k, n] of Object.entries(OPTIONAL_LISTS)) {
      const l = (lists.value || []).find(x => x.displayName === n);
      if (l) listIds[k] = l.id;
    }
    ids = { siteId: site.id, driveId: drive.id, lists: listIds };
    return ids;
  }

  const person = x => (x && x.user ? { id: x.user.id, name: x.user.displayName, email: (x.user.email || '').toLowerCase() } : { id: null, name: 'Unknown', email: '' });
  const flat = i => ({ id: Number(i.id), f: i.fields, createdAt: i.createdDateTime, createdBy: person(i.createdBy) });

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
  async function removeItem(key, id) { const { siteId, lists } = await resolve(); return graph('DELETE', `/sites/${siteId}/lists/${lists[key]}/items/${id}`); }

  // Every action goes in the activity log. SharePoint records who added each entry.
  // The action itself has already been saved by the time this runs, so a problem with
  // the log never turns a completed action into an error. If SharePoint rejects a
  // column, the entry is saved without it and the person is told what needs fixing.
  async function log(action, ref, detail) {
    const fields = { Title: String(action).slice(0, 255), Ref: ref ? String(ref).slice(0, 100) : '', PersonName: S.me.name, PersonEmail: S.me.email, Detail: detail ? JSON.stringify(detail).slice(0, 60000) : '', Ok: true };
    try { await addItem('activity', fields); return; } catch (first) {
      const bad = [];
      for (const k of Object.keys(fields).filter(k => k !== 'Title')) {
        const trial = Object.fromEntries(Object.entries(fields).filter(([key]) => key !== k));
        try { await addItem('activity', trial); bad.push(k); S.logWarning = `Done, but the activity log rejected its "${k}" column. Fix it in IREC Activity > List settings (see the setup guide).`; return; } catch (e) { /* try the next column */ }
      }
      try { await addItem('activity', { Title: fields.Title }); S.logWarning = 'Done, but the activity log only accepted the description. Check the IREC Activity columns (see the setup guide).'; }
      catch (e) { S.logWarning = `Done, but it could not be written to the activity log: ${first.message}`; }
    }
  }

  // Reads a file from the Portal data folder. SharePoint doesn't always include the
  // temporary download link, so ask for the full item, then fall back to downloading
  // through Microsoft Graph directly. Never fetch a missing link.
  async function readFile(name) {
    const { driveId } = await resolve();
    let meta;
    try { meta = await graph('GET', `/drives/${driveId}/root:/${enc(`${C.folder}/${name}`)}`); }
    catch (e) { if (e.notFound) return null; throw e; }
    let res = null;
    const link = meta['@microsoft.graph.downloadUrl'];
    if (link) {
      try { res = await fetch(link); } catch (e) { res = null; }
    }
    if (!res || !res.ok) {
      res = await fetch(`${GRAPH}/drives/${driveId}/items/${meta.id}/content`, { headers: { Authorization: `Bearer ${await token()}` } });
    }
    if (!res.ok) throw new Error(`Could not download ${name} from SharePoint (HTTP ${res.status}). Try Reload; if it persists, check you can open the "${C.folder}" folder in SharePoint.`);
    try { return await res.json(); }
    catch (e) { throw new Error(`${name} in SharePoint is not readable data. Run the export again, or import the downloaded file.`); }
  }

  async function writeFile(name, text) {
    const { driveId } = await resolve();
    const bytes = new TextEncoder().encode(text);
    const path = enc(`${C.folder}/${name}`);
    if (bytes.length <= 4 * 1024 * 1024) return graph('PUT', `/drives/${driveId}/root:/${path}:/content`, bytes, { 'Content-Type': 'application/json' }, true);
    const session = await graph('POST', `/drives/${driveId}/root:/${path}:/createUploadSession`, { item: { '@microsoft.graph.conflictBehavior': 'replace' } });
    const CHUNK = 320 * 1024 * 16;
    for (let start = 0; start < bytes.length; start += CHUNK) {
      const end = Math.min(start + CHUNK, bytes.length);
      const res = await fetch(session.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${bytes.length}` }, body: bytes.slice(start, end) });
      if (!res.ok && res.status !== 202) throw new Error(`Uploading ${name} to SharePoint failed (HTTP ${res.status}).`);
    }
    return null;
  }

  // ---------- reading the exporter's file (.json, or the .zip GitHub provides) ----------
  async function readJsonFromZip(buf) {
    const dv = new DataView(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) throw new Error('That file is not a valid zip file.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('That zip file is damaged.');
      const method = dv.getUint16(p + 10, true), size = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true), local = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(new Uint8Array(buf, p + 46, nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      if (!/\.json$/i.test(name)) continue;
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      const data = new Uint8Array(buf, start, size);
      if (method === 0) return new TextDecoder().decode(data);
      if (method === 8) {
        if (typeof DecompressionStream === 'undefined') throw new Error('This browser cannot open zip files. Unzip it first and choose the .json file inside.');
        const out = await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer();
        return new TextDecoder().decode(out);
      }
      throw new Error('That zip file uses an unsupported format. Unzip it and choose the .json file inside.');
    }
    throw new Error('No registry data file was found inside that zip.');
  }

  function checkSnapshot(s) {
    const lists = ['accounts', 'devices', 'issues', 'items', 'transactions', 'beneficiaries', 'reservations'];
    if (!s || s.format !== FORMAT) throw new Error("That isn't a registry data file from the exporter.");
    if (s.version !== 1) throw new Error(`This data file is version ${s.version}; this portal reads version 1. Update the portal or the exporter so they match.`);
    if (!lists.every(k => Array.isArray(s[k])) || isNaN(new Date(s.syncedAt))) throw new Error('The data file is incomplete or damaged. Run the export again.');
    return s;
  }

  async function prepareImport(file) {
    const buf = await file.arrayBuffer();
    const isZip = new DataView(buf).byteLength > 4 && new DataView(buf).getUint32(0, true) === 0x04034b50;
    let text;
    try { text = isZip ? await readJsonFromZip(buf) : new TextDecoder().decode(buf); } catch (e) { throw e; }
    let snap;
    try { snap = JSON.parse(text); } catch (e) { throw new Error("That file isn't valid registry data."); }
    return checkSnapshot(snap);
  }

  // ---------- loading ----------
  const EMPTY = { syncedAt: null, accounts: [], devices: [], issues: [], items: [], transactions: [], beneficiaries: [], reservations: [] };

  async function load() {
    const hasRequests = !!(await resolve()).lists.requests;
    const [snap, info, ems, acts, reqs] = await Promise.all([
      readFile('registry-snapshot.json'), readFile('import-status.json'),
      items('earmarks'), items('activity', { orderby: 'fields/Created desc', top: 300 }),
      hasRequests ? items('requests') : Promise.resolve([])
    ]);
    const earmarks = {};
    ems.forEach(e => { (earmarks[e.f.Title] = earmarks[e.f.Title] || []).push({ id: e.id, client: e.f.Client, byName: e.createdBy.name }); });
    const base = Object.assign({}, EMPTY, snap || {});
    S.data = Object.assign(base, { info: info || null, earmarks, hasRequests, requests: shapeRequests(reqs, base) });
    S.audit = acts.map(a => ({ at: a.createdAt, name: a.createdBy.name !== 'Unknown' ? a.createdBy.name : a.f.PersonName, action: a.f.Title, ref: a.f.Ref, ok: a.f.Ok !== false }));
  }

  // Requests made in the portal (the IREC Requests list). The worker on GitHub turns them into drafts.
  function shapeRequests(rows, snap) {
    const bens = Object.fromEntries((snap.beneficiaries || []).map(b => [b.uid, b]));
    const devs = Object.fromEntries((snap.devices || []).map(d => [d.code, d]));
    return rows.map(r => {
      const raw = String(r.f.Kind || '').toLowerCase(), kind = raw === 'issue' ? 'issue' : 'reservation';
      let p = {}, res = {};
      try { p = JSON.parse(r.f.Payload || '{}'); } catch (e) { p = {}; }
      try { res = JSON.parse(r.f.Result || '{}'); } catch (e) { res = {}; }
      const summary = kind === 'issue'
        ? { deviceName: res.deviceName || (devs[p.deviceCode] || {}).name || p.deviceCode, startDate: p.startDate, endDate: p.endDate, volume: Number(p.volume), recipientAccount: p.recipientAccount }
        : { type: raw === 'transfer' ? 'Transfer' : 'Redemption', party: res.beneficiaryName || (bens[p.beneficiaryUid] || {}).name || res.destinationAccount || p.destinationAccount || '', volume: Number(p.volume), itemUids: p.itemUids || [] };
      return { id: r.id, kind, ref: refOf(kind, r.id), status: r.f.RequestStatus || 'requested', summary, error: r.f.ErrorMessage || '', registryUid: r.f.RegistryUid || null, createdBy: r.createdBy, createdAt: r.createdAt };
    }).filter(r => r.status !== 'deleted' && r.status !== 'done').sort((a, b) => b.id - a.id);
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
    try { await fn(); } catch (e) { say(e.message, true); }
    finally { S.busy = false; document.body.removeAttribute('aria-busy'); }
  }

  // ---------- derived data ----------
  function derive() {
    const d = S.data;
    const queued = {};
    d.requests.filter(r => r.kind === 'reservation' && ['requested', 'processing'].includes(r.status)).forEach(r => (r.summary.itemUids || []).forEach(u => { queued[u] = r; }));
    const items = d.items.map(it => {
      const held = Math.max(0, it.volume - it.available), trade = isTrade(it.accountType), e = (d.earmarks[it.uid] || [])[0];
      const status = !trade ? 'redeemed' : (queued[it.uid] ? 'queued' : it.available <= 0 && held > 0 ? 'held' : e ? 'earmarked' : held > 0 ? 'held' : 'available');
      return Object.assign({}, it, { held, trade, earmark: e, status, queuedIn: queued[it.uid] || null });
    });
    return {
      me: S.me, canEdit: ORDER[S.me.role] >= 2,
      items, byUid: Object.fromEntries(items.map(i => [i.uid, i])),
      beneficiaries: d.beneficiaries.filter(b => b.active).sort((a, b) => a.name.localeCompare(b.name)),
      pendingRes: d.reservations.filter(r => /draft|submitted/i.test(r.status)),
      myRequests: d.requests.filter(r => (OPEN_REQ.includes(r.status) || r.status === 'failed') && r.createdBy && r.createdBy.id === S.me.id)
    };
  }

  // ---------- views ----------
  function viewNav(D) {
    const navItems = [['dashboard', 'Overview'], ['inventory', 'Inventory'], ['reservations', 'Redemptions and transfers'], ['issuance', 'Issuance'], ['audit', 'Activity log']];
    if (D.canEdit) navItems.push(['import', 'Import registry data']);
    return `<nav class="nav" aria-label="Main">${navItems.map(([k, l]) => `<button ${attr({ a: 'go', s: k })} ${S.screen === k ? 'aria-current="page"' : ''}><span>${l}</span></button>`).join('')}</nav>`;
  }

  function viewSync() {
    const d = S.data, info = d.info;
    const ageDays = d.syncedAt ? Math.floor((Date.now() - new Date(d.syncedAt)) / 86400000) : null;
    return `<div class="syncbox"><div class="t">Registry data</div>
      <div class="v">${d.syncedAt ? 'As of ' + esc(dTime(d.syncedAt)) : 'Not imported yet'}</div>
      ${info ? `<div class="t">${info.automatic ? 'Uploaded automatically' : `Imported by ${esc(info.importedBy)}`}, ${esc(dTime(info.importedAt))}</div>` : ''}
      ${info && info.automatic ? renewNotice(info.signInRenewBy) : ''}
      ${d.environment === 'sandbox' ? '<div class="warn">This is SANDBOX data, not production.</div>' : ''}
      ${ageDays !== null && ageDays >= 7 ? `<div class="warn">This data is ${ageDays} days old.</div>` : ''}</div>`;
  }

  // The exporter's SharePoint sign-in lasts 90 days unless it renews itself.
  function renewNotice(by) {
    if (!by) return '';
    const days = Math.ceil((new Date(`${by}T00:00:00Z`) - Date.now()) / 86400000);
    if (days > 14) return '';
    return days < 0
      ? `<div class="warn">Automatic upload stopped: its SharePoint sign-in expired on ${esc(dDate(by))}. Renew it with sign-in-worker.ps1, or import by hand.</div>`
      : `<div class="warn">Automatic upload's SharePoint sign-in must be renewed by ${esc(dDate(by))} (${days} day${days === 1 ? '' : 's'}).</div>`;
  }

  function viewDashboard(D) {
    const trade = D.items.filter(i => i.trade);
    const free = sum(trade.filter(i => !i.earmark), i => i.available);
    const emv = sum(trade.filter(i => i.earmark), i => i.available);
    const held = sum(trade, i => i.held);
    const clients = new Set(trade.filter(i => i.earmark && i.available > 0).map(i => i.earmark.client)).size;
    const kpis = [
      ['Free to allocate', fmt(free) + ' MWh', 'Not earmarked or in a registry draft'],
      ['Earmarked for clients', fmt(emv) + ' MWh', `${clients} client${clients === 1 ? '' : 's'}`],
      ['Held in registry drafts', fmt(held) + ' MWh', 'Reserved in the registry'],
      ['Waiting in the registry', String(D.pendingRes.length), 'Drafts or submitted, not yet approved']
    ];
    const byDev = {};
    trade.forEach(i => { const n = i.deviceName || 'Unknown device'; byDev[n] = (byDev[n] || 0) + i.available; });
    const devs = Object.entries(byDev).sort((a, b) => b[1] - a[1]);
    const max = Math.max(1, ...devs.map(x => x[1]));
    const empty = !S.data.syncedAt;
    return `${empty ? `<div class="card card-b"><h2 class="h2">No registry data yet</h2><p class="muted mt4">${D.canEdit ? 'Run the export on GitHub. With automatic upload it appears here by itself; otherwise use <b>Import registry data</b>.' : 'Ask an operator to run the registry export.'}</p></div>` : ''}
    <div class="kpis">${kpis.map(k => `<div class="card kpi"><span class="muted small">${k[0]}</span><span class="v">${k[1]}</span><span class="muted small">${k[2]}</span></div>`).join('')}</div>
    <div class="row">
      <section class="card col-main"><div class="card-h"><h2>Waiting in the registry</h2><span class="muted small">${D.pendingRes.length}</span></div>
        ${D.pendingRes.length ? D.pendingRes.slice(0, 8).map(r => `<div class="listrow"><div><b>${esc(r.type || 'Reservation')}</b>, ${fmt(r.volume)} MWh<div class="muted small">${esc(r.beneficiary || r.destination || '')}</div></div><div class="actions">${regPill(r.status)}</div></div>`).join('') : '<p class="empty">Nothing waiting in the registry.</p>'}
      </section>
      <section class="card card-b col-side"><h2 class="h2">Available by device</h2>
        ${devs.length ? devs.map(([n, v]) => `<div><div class="kv"><span class="kv-strong">${esc(n)}</span><span>${fmt(v)} MWh</span></div><div class="bar"><span data-pct="${Math.round(v / max * 100)}"></span></div></div>`).join('') : '<p class="muted mt0">No certificates in the imported data.</p>'}
      </section>
    </div>
    ${D.myRequests.length ? `<section class="card"><div class="card-h"><h2>Your requests</h2><span class="muted small">${D.myRequests.length}</span></div>
      ${D.myRequests.slice(0, 6).map(r => `<div class="listrow"><div><b>${esc(r.ref)}</b><span class="muted">, ${r.kind === 'issue' ? 'issue request' : esc(String(r.summary.type).toLowerCase())}, ${fmt(r.summary.volume)} MWh</span>${r.error ? `<div class="small error">${esc(r.error)}</div>` : ''}</div><div class="actions">${reqPill(r)}</div></div>`).join('')}</section>` : ''}
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
        <div><label class="lbl" for="f-status">Status</label><select id="f-status" class="fld" data-filter="status">${opt('', 'All statuses', S.f.status)}${Object.keys(LABEL).map(k => opt(k, LABEL[k], S.f.status)).join('')}</select></div>
      </div>
      ${D.canEdit && selected.length && !S.form ? `<div class="selbar"><b>${selected.length} selected, ${fmt(selAvail)} MWh available</b><div class="actions">
        <button class="btn primary" ${attr({ a: 'form', f: 'earmark' })}>Earmark for a client</button>
        ${S.data.hasRequests ? `<button class="btn primary" ${attr({ a: 'form', f: 'res' })}>Request a redemption or transfer</button>` : ''}
        <button class="btn" data-a="clearsel">Clear selection</button></div></div>` : ''}
      ${S.form === 'earmark' ? viewEarmarkForm(D, selAvail) : ''}
      ${S.form === 'res' ? viewResForm(D, selected, selAvail) : ''}
      <div class="row">
        <section class="card grow"><div class="card-h"><h2>Certificate blocks</h2><span class="muted small">${rows.length} shown, ${fmt(sum(rows, i => i.volume))} MWh</span></div>
          <div class="tablewrap"><table><thead><tr><th><span class="sr">Select</span></th><th>Certificate</th><th>Device</th><th>Period</th><th>Account</th><th class="num">Volume</th><th class="num">Available</th><th>Status</th><th>Earmarked for</th></tr></thead><tbody>
          ${rows.map(i => `<tr class="${det && det.uid === i.uid ? 'current' : S.sel.has(i.uid) ? 'sel' : ''}">
            <td><input type="checkbox" ${attr({ a: 'toggle', id: i.uid })} ${S.sel.has(i.uid) ? 'checked' : ''} ${!D.canEdit || !i.trade || i.available <= 0 ? 'disabled' : ''} aria-label="Select ${esc(short(i.uid))}"></td>
            <td><button class="linkbtn mono" ${attr({ a: 'open', id: i.uid })}>${esc(short(i.uid))}</button></td>
            <td>${esc(i.deviceName || '–')}</td><td>${period(i.periodStart, i.periodEnd)}</td><td>${esc(i.accountName)}</td>
            <td class="num">${fmt(i.volume)}</td><td class="num"><b>${i.trade ? fmt(i.available) : '–'}</b></td><td>${pill(i.status)}</td><td>${esc(i.earmark ? i.earmark.client : '–')}</td></tr>`).join('')}
          </tbody></table>${rows.length ? '' : '<p class="empty">No certificate blocks match these filters.</p>'}</div>
        </section>
        ${det ? viewItemDetail(D, det) : ''}
      </div>`;
  }

  function viewResForm(D, selected, selAvail) {
    const fv = S.fv, red = fv.type !== 'Transfer';
    return `<section class="card accent form"><div><h2>Request a draft from ${selected.length} certificate block${selected.length === 1 ? '' : 's'}</h2>
      <p class="muted small mt4">Your request is queued. The next time the export runs on GitHub, it checks the request and creates a <b>Draft</b> in the registry. Nothing moves or is redeemed until your approver submits and approves it in the registry.</p></div>
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

  function viewEarmarkForm(D, selAvail) {
    const names = [...new Set(D.beneficiaries.map(b => b.name))];
    return `<section class="card accent form"><div><h2>Earmark ${fmt(selAvail)} MWh for a client</h2>
      <p class="muted small mt4">An earmark is a note in SharePoint only. Nothing changes in the registry. Create the draft in the registry when you're ready.</p></div>
      <div class="fields"><div class="wide"><label class="lbl" for="em-client">Client (beneficiary)</label><select id="em-client" class="fld" data-f="client">${opt('', 'Choose a client', S.fv.client)}${names.map(n => opt(n, n, S.fv.client)).join('')}</select></div>
      <div class="wide"><label class="lbl" for="em-note">Remark (optional)</label><input id="em-note" class="fld" data-f="note" value="${esc(S.fv.note)}" placeholder="For example: FY2026 renewal, confirm volume first"></div></div>
      ${S.err ? `<p class="error" role="alert">${esc(S.err)}</p>` : ''}
      <div class="actions"><button class="btn primary" data-a="earmark">Save earmark</button><button class="btn" data-a="cancel">Cancel</button></div></section>`;
  }

  function viewItemDetail(D, i) {
    return `<aside class="card aside"><div class="panel-sec panel-head"><div class="panel-title"><span class="muted small">Certificate block</span><span class="t"><span class="mono">${esc(i.uid)}</span></span><span>${pill(i.status)}</span></div>
      <button class="btn sm" data-a="close">Close</button></div>
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
      <div class="panel-sec"><h3>Remarks</h3>
        ${S.remarksKey !== i.uid ? '<p class="muted small mt0">Loading…</p>' : S.remarks.length ? S.remarks.map(m => `<div class="remark"><span class="m"><b>${esc(m.byName)}</b>, ${esc(dTime(m.at))}</span><span>${esc(m.text)}</span></div>`).join('') : '<p class="muted small mt0">No remarks yet.</p>'}
        ${D.canEdit ? `<label class="lbl flush" for="rm-text">Add a remark</label><textarea id="rm-text" class="fld" rows="3" data-f="remark">${esc(S.fv.remark)}</textarea>
        <div><button class="btn primary sm" ${attr({ a: 'remark', id: i.uid })}>Add remark</button></div>` : ''}</div></aside>`;
  }

  function requestActions(D, r) {
    if (!D.canEdit) return '';
    const b = (act, text, cls) => `<button class="btn sm ${cls || ''}" ${attr({ a: 'reqact', act, id: r.id })}>${text}</button>`;
    if (r.status === 'requested') return b('cancel', 'Cancel', 'danger');
    if (r.status === 'failed') return b('cancel', 'Remove');
    if (r.status === 'draft') return r.kind === 'issue' ? b('withdraw', 'Mark as withdrawn', 'danger') : b('delete', 'Delete draft', 'danger');
    return '';
  }

  function viewRequestsTable(D, list, title) {
    if (!S.data.hasRequests) return `<p class="note">Drafts can't be requested here yet: the SharePoint list <b>IREC Requests</b> hasn't been created. See the setup guide.</p>`;
    return `<section class="card"><div class="card-h"><h2>${title}</h2><span class="muted small">${list.length}</span></div>
      <p class="note pad">Requests become Drafts in the registry when the export next runs on GitHub. Approval stays in the registry.</p>
      <div class="tablewrap"><table><thead><tr><th>Reference</th><th>${list[0] && list[0].kind === 'issue' ? 'Device' : 'Type'}</th><th>${list[0] && list[0].kind === 'issue' ? 'Period' : 'Beneficiary or receiving account'}</th><th class="num">Volume</th><th>Status</th><th>Requested by</th><th>Requested</th><th></th></tr></thead><tbody>
      ${list.map(r => `<tr><td class="mono">${esc(r.ref)}</td>
        <td>${r.kind === 'issue' ? esc(r.summary.deviceName) : esc(r.summary.type)}</td>
        <td class="wrap">${r.kind === 'issue' ? period(r.summary.startDate, r.summary.endDate) : esc(r.summary.party || '–')}</td>
        <td class="num">${fmt(r.summary.volume)}</td>
        <td class="wrap">${reqPill(r)}${r.error ? `<div class="small ${r.status === 'failed' ? 'error' : 'muted'}">${esc(r.error)}</div>` : ''}</td>
        <td>${esc(r.createdBy.name)}</td><td>${dDate(r.createdAt)}</td><td>${requestActions(D, r)}</td></tr>`).join('')}
      </tbody></table>${list.length ? '' : '<p class="empty">No requests yet.</p>'}</div></section>`;
  }

  function viewReservations(D) {
    const res = S.data.reservations.slice().sort((a, b) => String(a.status).localeCompare(String(b.status)));
    const tx = S.data.transactions.slice().sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp))).slice(0, 100);
    return `${viewRequestsTable(D, S.data.requests.filter(r => r.kind === 'reservation'), 'Requests from the portal')}
      <p class="note">Drafts are submitted and approved in the registry. The tables below show the registry as of the last refresh.</p>
      <section class="card"><div class="card-h"><h2>Drafts and submissions in the registry</h2><span class="muted small">${res.length}</span></div><div class="tablewrap"><table>
        <thead><tr><th>Registry ID</th><th>Type</th><th>Beneficiary or receiving account</th><th>From</th><th class="num">Volume</th><th>Status</th></tr></thead><tbody>
        ${res.map(r => `<tr><td class="mono">${esc(short(r.uid))}</td><td>${esc(r.type || '–')}</td><td class="wrap">${esc(r.beneficiary || r.destination || '–')}</td><td class="mono">${esc(r.source || '–')}</td><td class="num">${fmt(r.volume)}</td><td>${regPill(r.status)}</td></tr>`).join('')}
        </tbody></table>${res.length ? '' : '<p class="empty">No drafts or submissions in the imported data.</p>'}</div></section>
      <section class="card"><div class="card-h"><h2>Completed transactions</h2><span class="muted small">Latest ${tx.length}</span></div><div class="tablewrap"><table>
        <thead><tr><th>Date</th><th>Type</th><th>From</th><th>To</th><th>Beneficiary</th><th class="num">Volume</th></tr></thead><tbody>
        ${tx.map(t => `<tr><td>${dDate(t.timestamp)}</td><td>${esc(t.type)}</td><td class="mono">${esc(t.source)}</td><td class="mono">${esc(t.destination)}</td><td class="wrap">${esc(t.beneficiary || '–')}</td><td class="num">${fmt(t.volume)}</td></tr>`).join('')}
        </tbody></table>${tx.length ? '' : '<p class="empty">No transactions in the imported data.</p>'}</div></section>`;
  }

  function viewIssueForm(D) {
    const dev = S.data.devices.find(d => d.code === S.fv.deviceCode);
    const trade = S.data.accounts.filter(a => isTrade(a.type) && a.active);
    return `<section class="card accent form"><div><h2>Request an issue for ${esc(dev ? dev.name : '')}</h2>
      <p class="muted small mt4">Queued for the next export on GitHub, which creates a <b>Draft</b> issue request in the registry. The Issuer only sees it once it is submitted there.</p></div>
      <div class="fields">
        <div><label class="lbl" for="is-start">Production from</label><input id="is-start" class="fld" type="date" data-f="startDate" value="${esc(S.fv.startDate)}"></div>
        <div><label class="lbl" for="is-end">Production to</label><input id="is-end" class="fld" type="date" data-f="endDate" value="${esc(S.fv.endDate)}"></div>
        <div><label class="lbl" for="is-vol">Production (MWh)</label><input id="is-vol" class="fld" inputmode="decimal" data-f="volume" value="${esc(S.fv.volume)}"></div>
        <div class="wide"><label class="lbl" for="is-acc">Deposit into</label><select id="is-acc" class="fld" data-f="recipientAccount">${trade.map(a => opt(a.code, `${a.name} (${a.code})`, S.fv.recipientAccount)).join('')}</select></div></div>
      <p class="muted small mt0">Attach meter readings and invoices to the request in the registry. The Issuer requires them.</p>
      ${S.err ? `<p class="error" role="alert">${esc(S.err)}</p>` : ''}
      <div class="actions"><button class="btn primary" data-a="createissue">Queue the request</button><button class="btn" data-a="cancel">Cancel</button></div></section>`;
  }

  function viewIssuance(D) {
    const devices = S.data.devices.slice().sort((a, b) => String(a.name).localeCompare(String(b.name)));
    const reg = S.data.issues.slice().sort((a, b) => String(b.periodEnd).localeCompare(String(a.periodEnd)));
    const lastIssued = {};
    reg.forEach(i => { if (/approved|issued/i.test(i.status) && !lastIssued[i.deviceCode]) lastIssued[i.deviceCode] = period(i.periodStart, i.periodEnd); });
    return `${S.form === 'issue' ? viewIssueForm(D) : ''}<div class="devices">${devices.map(d => `<div class="card card-b device">
        <div class="kv"><span class="device-name"><b>${esc(d.name)}</b><span class="muted small">${esc(d.fuel)}${d.capacity ? ', ' + fmt(d.capacity) + ' MW' : ''}, <span class="mono">${esc(d.code)}</span></span></span><span>${regPill(d.status)}</span></div>
        <span class="small">Last issued: ${lastIssued[d.code] || 'None yet'}</span>
        ${D.canEdit && S.data.hasRequests && /approved/i.test(d.status) ? `<div><button class="btn sm" ${attr({ a: 'issueform', code: d.code })}>Request an issue</button></div>` : ''}</div>`).join('') || '<p class="muted">No devices in the imported data.</p>'}</div>
      <section class="card"><div class="card-h"><h2>Issue requests in the registry</h2></div><div class="tablewrap"><table>
        <thead><tr><th>Device</th><th>Period</th><th class="num">Production</th><th class="num">Issued</th><th>Status</th></tr></thead><tbody>
        ${reg.slice(0, 200).map(i => `<tr><td>${esc(i.deviceName)}</td><td>${period(i.periodStart, i.periodEnd)}</td><td class="num">${fmt(i.productionVolume)}</td><td class="num">${fmt(i.issuedVolume)}</td><td>${regPill(i.status)}</td></tr>`).join('')}
        </tbody></table>${reg.length ? '' : '<p class="empty">No issue requests in the imported data.</p>'}</div></section>
      ${viewRequestsTable(D, S.data.requests.filter(r => r.kind === 'issue'), 'Issue requests from the portal')}`;
  }

  function viewAudit() {
    return `<section class="card"><div class="card-h stack"><h2>Who did what in the portal</h2>
      <span class="muted small">Each entry's name comes from SharePoint, which records who added it. Entries cannot be added under someone else's name.</span></div>
      <div class="tablewrap"><table><thead><tr><th>When</th><th>Who</th><th>What</th><th>Reference</th></tr></thead><tbody>
      ${S.audit.map(e => `<tr><td>${esc(dTime(e.at))}</td><td><b>${esc(e.name)}</b></td><td class="wrap">${esc(e.action)}</td><td class="mono">${esc(e.ref || '–')}</td></tr>`).join('')}
      </tbody></table>${S.audit.length ? '' : '<p class="empty">No activity yet.</p>'}</div></section>`;
  }

  function viewImport() {
    const p = S.pending, cur = S.data;
    return `<section class="card card-b form">
      <h2>Import registry data</h2>
      <p class="note">When automatic upload is set up, running the export on GitHub updates the portal by itself, so this page is only a fallback. Click <b>Reload</b> to see the newest data.</p>
      <ol class="steps">
        <li>On GitHub, open the private <b>saxonirecworker</b> repository → <b>Actions</b> → <b>Export I-REC registry data</b> → <b>Run workflow</b>.</li>
        <li>When it finishes (about a minute), open the run and download <b>irec-registry-data</b> under Artifacts.</li>
        <li>Choose that file below. There is no need to unzip it.</li>
      </ol>
      <div><label class="lbl" for="imp-file">Registry data file (.zip or .json)</label><input id="imp-file" class="fld" type="file" accept=".zip,.json,application/zip,application/json" data-a="pickfile"></div>
      ${S.err ? `<p class="error" role="alert">${esc(S.err)}</p>` : ''}
      ${p ? `<div class="panel-sec"><h3>Check before importing</h3>
        <div class="kv"><span>Read from the registry</span><span><b>${esc(dTime(p.syncedAt))}</b></span></div>
        <div class="kv"><span>Environment</span><span><b>${esc(p.environment)}</b></span></div>
        <div class="kv"><span>Contents</span><span>${esc(Object.entries(p.counts || {}).map(([k, v]) => `${v} ${k}`).join(', '))}</span></div>
        ${cur.syncedAt && new Date(p.syncedAt) < new Date(cur.syncedAt) ? `<p class="note warn">This file is older than the data already in the portal (${esc(dTime(cur.syncedAt))}).</p>` : ''}
        ${cur.environment === 'production' && p.environment === 'sandbox' ? '<p class="note warn">This is sandbox data. It will replace production data in the portal.</p>' : ''}
        <div class="actions"><button class="btn primary" data-a="doimport">Import</button><button class="btn" data-a="cancelimport">Cancel</button></div></div>` : ''}
    </section>`;
  }

  // ---------- render ----------
  function render() {
    if (!S.data) return;
    const D = derive();
    const active = document.activeElement;
    const focusId = active && active.id !== 'imp-file' && active.id;
    const caret = active && typeof active.selectionStart === 'number' ? [active.selectionStart, active.selectionEnd] : null;
    const screens = { dashboard: viewDashboard, inventory: viewInventory, reservations: viewReservations, issuance: viewIssuance, audit: viewAudit, import: viewImport };
    if (S.screen === 'import' && !D.canEdit) S.screen = 'dashboard';
    $app.innerHTML = `<div class="layout">
      <aside class="side"><div class="brand"><div class="brand-name">Saxon Renewables</div><div class="brand-sub">I-REC portal</div></div>${viewNav(D)}${viewSync()}</aside>
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

  async function loadRemarks(uid) {
    S.remarksKey = '';
    try {
      const list = await items('remarks', { filter: `fields/TargetId eq '${q(uid)}'` });
      const mapped = list.filter(m => m.f.Title === 'item').sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
        .map(m => ({ text: m.f.RemarkText, byName: m.createdBy.name, at: m.createdAt }));
      if (S.detail && S.detail.id === uid) { S.remarks = mapped; S.remarksKey = uid; render(); }
    } catch (e) { say(e.message, true); }
  }
  async function refreshAll(msg) {
    await load(); render();
    if (S.logWarning) { const w = S.logWarning; S.logWarning = null; say(w, true); }
    else if (msg) say(msg);
  }

  // ---------- events ----------
  document.addEventListener('click', e => {
    const el = e.target.closest('[data-a]');
    if (!el || el.disabled) return;
    const a = el.dataset.a;
    if (a === 'signin') { msal.loginRedirect({ scopes: SCOPES }); return; }
    if (a === 'signout') { msal.logoutRedirect({ account: msal.getActiveAccount(), postLogoutRedirectUri: HOME }); return; }
    if (a === 'reload') { if (S.data) run(() => refreshAll('Reloaded.')); else location.reload(); return; }
    if (a === 'toggle' || a === 'pickfile') return;
    if (!S.data) return;
    if (a === 'go') { S.screen = el.dataset.s; S.detail = null; S.form = null; S.err = ''; S.pending = null; render(); window.scrollTo(0, 0); return; }
    if (a === 'open') { S.screen = 'inventory'; S.detail = { kind: 'item', id: el.dataset.id }; S.form = null; S.fv.remark = ''; render(); loadRemarks(el.dataset.id); return; }
    if (a === 'close') { S.detail = null; render(); return; }
    if (a === 'toast') { S.toast = null; render(); return; }
    if (a === 'clearsel') { S.sel.clear(); render(); return; }
    if (a === 'cancel') { S.form = null; S.err = ''; render(); return; }
    if (a === 'form') {
      S.form = el.dataset.f; S.err = ''; S.detail = null;
      const y = new Date().getFullYear(), avail = sum([...S.sel].map(u => derive().byUid[u]).filter(Boolean), i => i.available);
      S.fv = el.dataset.f === 'res' ? { type: 'Redemption', volume: String(avail), beneficiaryUid: '', purpose: `Scope 2 reporting ${y}`, periodStart: `${y}-01-01`, periodEnd: '', destinationAccount: '' } : { client: '', note: '' };
      render(); return;
    }
    if (a === 'issueform') {
      const trade = S.data.accounts.filter(x => isTrade(x.type) && x.active);
      S.form = 'issue'; S.err = ''; S.detail = null;
      S.fv = { deviceCode: el.dataset.code, startDate: '', endDate: '', volume: '', recipientAccount: (trade[0] || {}).code || '' };
      render(); return;
    }
    if (a === 'createres') run(async () => {
      const fv = S.fv, D = derive(), vol = Number(String(fv.volume).replace(/,/g, ''));
      const picked = [...S.sel].map(u => D.byUid[u]).filter(Boolean), avail = sum(picked, i => i.available);
      const busy = picked.find(i => i.queuedIn);
      if (!picked.length) { S.err = 'Select at least one certificate block first.'; render(); return; }
      if (busy) { S.err = `${short(busy.uid)} is already in the queued request ${busy.queuedIn.ref}. Wait for it, or cancel it first.`; render(); return; }
      if (!(vol > 0)) { S.err = 'Enter a volume above zero.'; render(); return; }
      if (vol > avail) { S.err = `Only ${fmt(avail)} MWh is available in the selected blocks.`; render(); return; }
      if (new Set(picked.map(i => i.accountCode)).size > 1) { S.err = 'All selected certificates must be in the same account.'; render(); return; }
      if (fv.type === 'Redemption' && !fv.beneficiaryUid) { S.err = 'Choose a beneficiary.'; render(); return; }
      if (fv.type === 'Redemption' && (!fv.periodStart || !fv.periodEnd || fv.periodEnd < fv.periodStart)) { S.err = 'Enter a consumption period that ends after it starts.'; render(); return; }
      if (fv.type === 'Redemption' && !String(fv.purpose || '').trim()) { S.err = 'Enter the reporting purpose.'; render(); return; }
      if (fv.type === 'Transfer' && !String(fv.destinationAccount).trim()) { S.err = 'Enter the receiving account code.'; render(); return; }
      const payload = fv.type === 'Transfer' ? { itemUids: [...S.sel], volume: vol, destinationAccount: String(fv.destinationAccount).trim() }
        : { itemUids: [...S.sel], volume: vol, beneficiaryUid: fv.beneficiaryUid, purpose: fv.purpose, periodStart: fv.periodStart, periodEnd: fv.periodEnd };
      const created = await addItem('requests', { Title: 'New request', Kind: fv.type.toLowerCase(), RequestStatus: 'requested', Payload: JSON.stringify(payload) });
      const ref = refOf('reservation', created.id);
      await log(`Requested a ${fv.type.toLowerCase()} draft, ${fmt(vol)} MWh`, ref, { itemUids: [...S.sel] });
      S.sel.clear(); S.form = null; S.screen = 'reservations';
      await refreshAll(`${ref} queued. The Draft is created in the registry the next time the export runs on GitHub.`);
    });
    if (a === 'createissue') run(async () => {
      const fv = S.fv, vol = Number(String(fv.volume).replace(/,/g, ''));
      if (!fv.startDate || !fv.endDate || fv.endDate < fv.startDate) { S.err = 'Enter a production period that ends after it starts.'; render(); return; }
      if (!(vol > 0)) { S.err = 'Enter the production volume in MWh.'; render(); return; }
      if (!fv.recipientAccount) { S.err = 'Choose the account to deposit into.'; render(); return; }
      const created = await addItem('requests', { Title: 'New request', Kind: 'issue', RequestStatus: 'requested', Payload: JSON.stringify({ deviceCode: fv.deviceCode, startDate: fv.startDate, endDate: fv.endDate, volume: vol, recipientAccount: fv.recipientAccount }) });
      const ref = refOf('issue', created.id);
      await log('Requested an issue request draft', ref, { deviceCode: fv.deviceCode });
      S.form = null;
      await refreshAll(`${ref} queued. The Draft is created in the registry the next time the export runs on GitHub.`);
    });
    if (a === 'reqact') run(async () => {
      const r = S.data.requests.find(x => x.id === Number(el.dataset.id)), act = el.dataset.act;
      if (!r) return;
      const asks = { cancel: `Remove ${r.ref} from the list?`, delete: `Delete the draft ${r.ref}? It is removed from the registry on the next export, and its certificates are freed.`, withdraw: `Mark ${r.ref} as withdrawn? You also need to withdraw it in the registry.` };
      if (!confirm(asks[act])) return;
      const { siteId, lists } = await resolve();
      if (act === 'cancel') await removeItem('requests', r.id);
      else await graph('PATCH', `/sites/${siteId}/lists/${lists.requests}/items/${r.id}/fields`, { RequestStatus: act === 'delete' ? 'delete-requested' : 'withdrawn', ErrorMessage: '' });
      await log(act === 'cancel' ? `Removed the request ${r.ref}` : act === 'delete' ? `Asked for the draft ${r.ref} to be deleted` : `Marked ${r.ref} as withdrawn`, r.ref);
      await refreshAll(act === 'delete' ? `${r.ref} will be deleted from the registry on the next export.` : act === 'withdraw' ? `${r.ref} marked as withdrawn.` : `${r.ref} removed.`);
    });
    if (a === 'cancelimport') { S.pending = null; S.pendingFile = null; S.err = ''; render(); return; }
    if (a === 'earmark') run(async () => {
      if (!S.fv.client) { S.err = 'Choose a client first.'; render(); return; }
      const client = S.fv.client, note = String(S.fv.note || '').trim(), uids = [...S.sel];
      for (const uid of uids) {
        for (const old of S.data.earmarks[uid] || []) await removeItem('earmarks', old.id);
        await addItem('earmarks', { Title: uid, Client: client });
        if (note) await addItem('remarks', { Title: 'item', TargetId: uid, RemarkText: note });
      }
      await log(`Earmarked ${uids.length} certificate block${uids.length === 1 ? '' : 's'} for ${client}`, null, { itemUids: uids, client });
      S.sel.clear(); S.form = null; await refreshAll(`Earmarked for ${client}. Nothing changed in the registry.`);
    });
    if (a === 'unearmark') run(async () => {
      const old = S.data.earmarks[el.dataset.id] || [];
      for (const o of old) await removeItem('earmarks', o.id);
      await log(`Removed the earmark for ${old[0] ? old[0].client : 'a client'}`, short(el.dataset.id), { itemUid: el.dataset.id });
      await refreshAll('Earmark removed.');
    });
    if (a === 'remark') run(async () => {
      const text = String(S.fv.remark || '').trim();
      if (!text) { say('Write a remark first.', true); return; }
      if (text.length > 2000) { say('Remarks are limited to 2,000 characters.', true); return; }
      await addItem('remarks', { Title: 'item', TargetId: el.dataset.id, RemarkText: text });
      await log('Added a remark', short(el.dataset.id), { itemUid: el.dataset.id });
      S.fv.remark = ''; S.audit = (await items('activity', { orderby: 'fields/Created desc', top: 300 })).map(x => ({ at: x.createdAt, name: x.createdBy.name, action: x.f.Title, ref: x.f.Ref, ok: x.f.Ok !== false }));
      await loadRemarks(el.dataset.id);
      if (S.logWarning) { const w = S.logWarning; S.logWarning = null; say(w, true); }
    });
    if (a === 'doimport') run(async () => {
      const p = S.pending;
      if (!p) return;
      const text = JSON.stringify(p);
      await writeFile('registry-snapshot.json', text);
      await writeFile(`history/${p.syncedAt.slice(0, 10)}-${p.environment}.json`, text);
      const info = { importedAt: new Date().toISOString(), importedBy: S.me.name, importedByEmail: S.me.email, syncedAt: p.syncedAt, environment: p.environment, counts: p.counts, exportRun: p.exportRun || null };
      await writeFile('import-status.json', JSON.stringify(info));
      await log(`Imported registry data (${p.environment}) read at ${dTime(p.syncedAt)}`, null, info);
      S.pending = null; S.screen = 'dashboard';
      await refreshAll('Registry data imported.');
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
    if (t.dataset.a === 'pickfile') {
      const file = t.files && t.files[0];
      S.pending = null; S.err = '';
      if (!file) { render(); return; }
      prepareImport(file).then(p => { S.pending = p; render(); }).catch(err => { S.err = err.message; render(); });
      return;
    }
    if (t.dataset.f) { S.fv[t.dataset.f] = t.value; return; }
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
        gate('Sign in to the I-REC portal', 'Use your Saxon Renewables work account. You will be asked to approve the sign-in in Microsoft Authenticator.',
          '<button class="btn primary lg" data-a="signin">Sign in with Microsoft</button>');
        return;
      }
      msal.setActiveAccount(account);
      const claims = account.idTokenClaims || {};
      const roles = (claims.roles || []).map(r => String(r).toLowerCase()).filter(r => PORTAL_ROLES.includes(r)).sort((x, y) => ORDER[y] - ORDER[x]);
      if (!roles.length && ['viewer', 'operator'].includes(C.defaultRole)) roles.push(C.defaultRole);
      if (!roles.length) {
        gate("You don't have access yet", "You're signed in, but no portal role has been assigned to your account. Ask the portal's owner to add you.",
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
