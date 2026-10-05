(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.StudioFSharedStorage = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const DOCUMENTS = new Set(['AppArbeitsstand.json', 'Disposition.json']);
  const MAX_PAYLOAD = 60000;
  const CHUNK_SIZE = 8000, MAX_CHUNKS = 200;
  const isChunk = name => /^Arbeitsstand_[a-f0-9-]{36}_\d{1,3}\.json$/.test(name);
  const digest = async text => Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)))).map(x => x.toString(16).padStart(2,'0')).join('');
  class StorageError extends Error {
    constructor(code, message) { super(message); this.name = 'StorageError'; this.code = code; }
  }
  function assertTestRoot(cfg) {
    const path = String(cfg.root || '').replace(/^\/+|\/+$/g, '');
    if (!/^Gutachten_Mehrbenutzer_TEST_[A-Za-z0-9_-]+$/.test(path)) {
      throw new StorageError('TEST_ROOT_REQUIRED', 'Diese Testversion darf nur einen eigenen Gutachten_Mehrbenutzer_TEST_-Ordner verwenden.');
    }
    return path;
  }
  function validateConfig(cfg) {
    assertTestRoot(cfg);
    if (cfg.storageMode !== 'shared') return;
    for (const key of ['driveId', 'siteId', 'stateListId']) {
      if (!String(cfg[key] || '').trim() || /[/?#]/.test(cfg[key])) {
        throw new StorageError('CONFIG_REQUIRED', 'Die gemeinsame Testablage ist noch nicht vollständig eingerichtet (' + key + ').');
      }
    }
  }
  function graphPath(path, cfg) {
    if (path.startsWith('https://')) {
      const url = new URL(path);
      if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/')) {
        throw new StorageError('INVALID_GRAPH_URL', 'Ungültiger Microsoft-Graph-Folgepfad.');
      }
      path = url.pathname.slice('/v1.0'.length) + url.search;
    }
    if (path === '/me/drive' || path.startsWith('/me/drive/')) {
      validateConfig(cfg);
      if (cfg.storageMode === 'shared') return '/drives/' + encodeURIComponent(cfg.driveId) + path.slice('/me/drive'.length);
    }
    return path;
  }
  async function collectPages(graph, path, cfg, options = {}) {
    let first = null, items = [], seen = new Set();
    while (path) {
      path = graphPath(path, cfg);
      if (seen.has(path) || seen.size >= 1000) throw new StorageError('INVALID_PAGING', 'Ungültige oder zyklische Microsoft-Seitenfolge.');
      seen.add(path);
      const page = await graph(path, options);
      if (!first) first = page;
      if (!Array.isArray(page?.value)) {
        if (seen.size > 1 || page?.['@odata.nextLink']) throw new StorageError('INVALID_PAGING', 'Unvollständige Microsoft-Seite.');
        return page;
      }
      items.push(...page.value);
      path = page['@odata.nextLink'] || '';
    }
    const result = {...first, value:items}; delete result['@odata.nextLink']; return result;
  }
  function scopes(cfg) {
    return ['openid', 'profile', 'offline_access', 'User.Read', cfg.storageMode === 'shared' ? 'Sites.Selected' : 'Files.ReadWrite'];
  }
  function fingerprint(cfg) {
    return [cfg.storageMode, cfg.tenant, cfg.driveId, cfg.siteId, cfg.stateListId, cfg.root].join('|');
  }
  function recordKey(cfg, caseId, name) {
    validateConfig(cfg);
    if ((!DOCUMENTS.has(name) && !isChunk(name)) || !caseId || /[/?#]/.test(caseId)) throw new StorageError('INVALID_RECORD', 'Ungültige Aktzuordnung.');
    const key = cfg.driveId + ':' + caseId + ':' + name;
    if (key.length > 255) throw new StorageError('INVALID_RECORD', 'Die Aktkennung ist zu lang.');
    return key;
  }
  const clone = value => JSON.parse(JSON.stringify(value));
  function parseRecord(item, key) {
    if (!item?.id || !item.eTag || item.fields?.RecordKey !== key || typeof item.fields.Payload !== 'string') {
      throw new StorageError('INVALID_RECORD', 'Der gemeinsame Datensatz ist unvollständig. Es wurde nichts überschrieben.');
    }
    let data;
    try { data = JSON.parse(item.fields.Payload); } catch (_) { throw new StorageError('INVALID_JSON', 'Der gemeinsame Arbeitsstand enthält ungültiges JSON.'); }
    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new StorageError('INVALID_JSON', 'Der gemeinsame Arbeitsstand ist kein Objekt.');
    return { id: item.id, eTag: item.eTag, data, lastModifiedDateTime: item.lastModifiedDateTime || '' };
  }
  function createStore({ graph, config, onState = () => {}, persistPending = () => {} }) {
    const baselines = new Map(), queues = new Map(), draftTokens = new Map();
    const cfg = Object.freeze({ ...config });
    validateConfig(cfg);
    const base = '/sites/' + encodeURIComponent(cfg.siteId) + '/lists/' + encodeURIComponent(cfg.stateListId);
    const pathFor = id => base + '/items/' + encodeURIComponent(id) + '?$expand=fields';
    async function lookup(key) {
      const filter = "fields/RecordKey eq '" + key.replace(/'/g, "''") + "'";
      let path = base + '/items?$expand=fields&$filter=' + encodeURIComponent(filter), result = [];
      const seen = new Set();
      while (path) {
        if (seen.has(path)) throw new StorageError('INVALID_PAGING', 'Zyklischer Microsoft-Folgepfad.');
        seen.add(path);
        const page = await graph(graphPath(path, cfg));
        if (!Array.isArray(page?.value)) throw new StorageError('INVALID_RECORD', 'Die Datenbankantwort ist unvollständig.');
        result.push(...page.value);
        path = page['@odata.nextLink'] || '';
      }
      if (result.length > 1) throw new StorageError('DUPLICATE_RECORD', 'Die Aktkennung ist in der Datenbank nicht eindeutig.');
      return result.length ? parseRecord(result[0], key) : null;
    }
    async function materialize(caseId, row) {
      if (!row || !Object.hasOwn(row.data, '_studiofChunks')) return row;
      const m = row.data._studiofChunks;
      if (m?.version !== 1 || !Array.isArray(m.names) || !m.names.length || m.names.length > MAX_CHUNKS || !/^[a-f0-9-]{36}$/.test(m.revision) || !/^[a-f0-9]{64}$/.test(m.sha256)) throw new StorageError('INVALID_CHUNKS', 'Ungültiger Arbeitsstand-Verweis.');
      let text = '';
      for (let i = 0; i < m.names.length; i++) {
        if (m.names[i] !== 'Arbeitsstand_' + m.revision + '_' + i + '.json') throw new StorageError('INVALID_CHUNKS', 'Ungültige Reihenfolge der Arbeitsstand-Teile.');
        const part = await lookup(recordKey(cfg, caseId, m.names[i]));
        if (!part || part.data.revision !== m.revision || part.data.index !== i || typeof part.data.text !== 'string' || part.data.text.length > CHUNK_SIZE) throw new StorageError('INCOMPLETE_WORK_STATE', 'Der gemeinsame Arbeitsstand ist unvollständig. Lokale Daten bleiben erhalten.');
        text += part.data.text;
      }
      if (await digest(text) !== m.sha256) throw new StorageError('INVALID_CHUNKS', 'Die Prüfsumme des Arbeitsstands stimmt nicht.');
      let data; try { data = JSON.parse(text); } catch (_) { throw new StorageError('INVALID_JSON', 'Ungültiges JSON im aufgeteilten Arbeitsstand.'); }
      if (!data || typeof data !== 'object' || Array.isArray(data) || Object.hasOwn(data, '_studiofChunks')) throw new StorageError('INVALID_CHUNKS', 'Ungültiger aufgeteilter Arbeitsstand.');
      return { ...row, data };
    }
    async function preview(caseId, name) {
      return materialize(caseId, await lookup(recordKey(cfg, caseId, name)));
    }
    async function resolveReviewed(caseId, name, reviewed, localData, actor) {
      const key = recordKey(cfg, caseId, name);
      if (queues.has(key)) throw new StorageError('SAVE_IN_FLIGHT', 'Ein Speichervorgang ist noch aktiv.');
      const localToken = draftTokens.get(key);
      const current = await preview(caseId, name);
      if(draftTokens.get(key) !== localToken) throw new StorageError('LOCAL_DRAFT_CHANGED', 'Der lokale Entwurf hat sich während des Vergleichs geändert. Bitte erneut vergleichen.');
      if ((current?.id || null) !== (reviewed?.id || null) || (current?.eTag || null) !== (reviewed?.eTag || null)) throw new StorageError('REVIEW_STALE', 'Die zentrale Fassung wurde seit dem Vergleich geändert. Bitte erneut vergleichen.');
      if (queues.has(key)) throw new StorageError('SAVE_IN_FLIGHT', 'Ein Speichervorgang ist noch aktiv.');
      baselines.set(key, current ? { id: current.id, eTag: current.eTag } : null);
      if (localData !== undefined) return write(caseId, name, localData, actor);
      persistPending(key, null); onState(key, 'loaded'); return current;
    }
    async function read(caseId, name) {
      const key = recordKey(cfg, caseId, name), row = await preview(caseId, name);
      // A background read must never silently acknowledge a newer remote version.
      if (!baselines.has(key)) baselines.set(key, row ? { id: row.id, eTag: row.eTag } : null);
      return row;
    }
    async function acceptRemote(caseId, name) {
      const key = recordKey(cfg, caseId, name);
      if (queues.has(key)) throw new StorageError('SAVE_IN_FLIGHT', 'Ein Speichervorgang ist noch aktiv.');
      const row = await preview(caseId, name);
      baselines.set(key, row ? { id: row.id, eTag: row.eTag } : null);
      persistPending(key, null);
      onState(key, 'loaded');
      return row;
    }
    function stage(caseId, name, data) {
      if (!DOCUMENTS.has(name) || !data || typeof data !== 'object' || Array.isArray(data) || Object.hasOwn(data, '_studiofChunks')) throw new StorageError('INVALID_RECORD', 'Ungültiger lokaler Arbeitsstand.');
      const key = recordKey(cfg, caseId, name), token = crypto.randomUUID();
      persistPending(key, {name, caseId, data:clone(data), updatedAt:new Date().toISOString(), draftToken:token});
      draftTokens.set(key, token); return token;
    }
    function write(caseId, name, data, actor, stagedToken) {
      const key = recordKey(cfg, caseId, name);
      const snapshot = clone(data), payload = JSON.stringify(snapshot);
      // Keep the pending draft before network access, including during offline use.
      if(stagedToken && draftTokens.get(key) !== stagedToken) return Promise.reject(new StorageError('LOCAL_DRAFT_CHANGED', 'Ein neuerer lokaler Entwurf liegt vor. Die ältere Fassung wird nicht übertragen.'));
      const token = stagedToken || stage(caseId, name, snapshot);
      if (payload.length > MAX_PAYLOAD && (name !== 'AppArbeitsstand.json' || payload.length > CHUNK_SIZE * MAX_CHUNKS)) return Promise.reject(new StorageError('PAYLOAD_TOO_LARGE', 'Dieser Arbeitsstand überschreitet die Testgrenze. Der lokale Entwurf ist gesichert.'));
      const previous = queues.get(key) || Promise.resolve();
      const task = previous.catch(() => {}).then(async () => {
        onState(key, 'saving');
        try {
          if (!baselines.has(key)) {
            const current = await lookup(key);
            if (current) throw new StorageError('READ_REQUIRED', 'Vor dem Bearbeiten bitte den gemeinsamen Akt laden. Der vorhandene Arbeitsstand wurde nicht überschrieben.');
            baselines.set(key, null);
          }
          const expected = baselines.get(key);
          let storedPayload = payload;
          if (payload.length > MAX_PAYLOAD) {
            const revision = crypto.randomUUID(), names = [];
            for (let offset = 0, index = 0; offset < payload.length; offset += CHUNK_SIZE, index++) {
              const partName = 'Arbeitsstand_' + revision + '_' + index + '.json';
              const chunkFields = { RecordKey: recordKey(cfg, caseId, partName), DocumentName: partName, CaseDriveItemId: caseId, Payload: JSON.stringify({revision, index, text:payload.slice(offset, offset + CHUNK_SIZE)}), UpdatedBy:String(actor || ''), UpdatedAt:new Date().toISOString(), SchemaVersion:2, WriteToken:crypto.randomUUID() };
              const created = await graph(base + '/items', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({fields:chunkFields})});
              if (!created?.id || !created.eTag) throw new StorageError('VERSION_UNCONFIRMED', 'Ein Teil des Arbeitsstands wurde nicht bestätigt. Der bisherige zentrale Stand bleibt gültig.');
              names.push(partName);
            }
            storedPayload = JSON.stringify({_studiofChunks:{version:1, revision, names, sha256:await digest(payload)}});
          }
          const fields = { RecordKey: key, DocumentName: name, CaseDriveItemId: caseId, Payload: storedPayload, UpdatedBy: String(actor || ''), UpdatedAt: new Date().toISOString(), SchemaVersion: 1, WriteToken: crypto.randomUUID() };
          let item;
          if (expected) {
            const updated = await graph(base + '/items/' + encodeURIComponent(expected.id) + '/fields', {
              method: 'PATCH', headers: { 'Content-Type': 'application/json', 'If-Match': expected.eTag }, body: JSON.stringify(fields)
            });
            // Read the exact version returned by PATCH, never a later user's version.
            const confirmed = await graph(pathFor(expected.id));
            const eTag = confirmed?.fields?.WriteToken === fields.WriteToken ? confirmed.eTag : null;
            if (!eTag) throw new StorageError('VERSION_UNCONFIRMED', 'Speichern wurde angenommen, die neue Version aber nicht bestätigt. Bitte den Akt erneut laden; nicht automatisch wiederholen.');
            item = { id: expected.id, eTag };
          } else {
            const created = await graph(base + '/items', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
            if (!created?.id || !created.eTag) throw new StorageError('VERSION_UNCONFIRMED', 'Die Anlage wurde nicht vollständig bestätigt. Bitte den Akt erneut laden.');
            item = { id: created.id, eTag: created.eTag };
          }
          baselines.set(key, item);
          // A later queued edit may already have replaced this draft; only clear the matching one.
          if (queues.get(key) === task && draftTokens.get(key) === token) { persistPending(key, null); draftTokens.delete(key); }
          onState(key, 'saved');
          return { ...item, data: snapshot };
        } catch (error) {
          if ([409, 412].includes(error.status) || error.code === 'READ_REQUIRED') {
            onState(key, 'conflict');
            throw new StorageError('CONFLICT', 'Der Akt wurde zwischenzeitlich geändert. Dein lokaler Entwurf bleibt erhalten. Bitte beide Fassungen vergleichen.');
          }
          // If the request outcome is unknown, never retry automatically using a fresh version.
          onState(key, 'pending');
          throw error;
        }
      });
      queues.set(key, task);
      task.finally(() => { if (queues.get(key) === task) queues.delete(key); }).catch(() => {});
      return task;
    }
    async function verify() {
      const [drive, list, columns] = await Promise.all([
        graph('/drives/' + encodeURIComponent(cfg.driveId) + '?$select=id,driveType'),
        graph(base + '?$select=id,displayName'), graph(base + '/columns')
      ]);
      if (drive?.id !== cfg.driveId || list?.id !== cfg.stateListId) throw new StorageError('CONFIG_MISMATCH', 'Die gemeinsame Ablage stimmt nicht mit der Konfiguration überein.');
      const byName = new Map((columns?.value || []).map(x => [x.name, x]));
      for (const name of ['RecordKey', 'DocumentName', 'CaseDriveItemId', 'Payload', 'UpdatedBy', 'UpdatedAt', 'SchemaVersion', 'WriteToken']) {
        if (!byName.has(name)) throw new StorageError('SCHEMA_REQUIRED', 'Die Datenbankspalte ' + name + ' fehlt.');
      }
      if (!byName.get('RecordKey').indexed || !byName.get('RecordKey').enforceUniqueValues) throw new StorageError('SCHEMA_REQUIRED', 'RecordKey muss eindeutig und indiziert sein.');
      if (!byName.get('Payload').text?.allowMultipleLines || byName.get('Payload').text?.textType !== 'plain') throw new StorageError('SCHEMA_REQUIRED', 'Payload muss mehrzeiliger Klartext sein.');
      return { driveId: drive.id, listId: list.id, label: list.displayName };
    }
    return { read, preview, resolveReviewed, stage, write, acceptRemote, verify, fingerprint: fingerprint(cfg), pendingCount: () => queues.size };
  }
  function assertWorkflowsDisabled() {
    throw new StorageError('TEST_FLOW_DISABLED', 'In dieser Testversion ist das Senden an Power Automate gesperrt. Zuerst separate Testflows einrichten und prüfen.');
  }
  function assertWordTestWorkflow(cfg) {
    if (cfg.wordTestFlowEnabled !== true) assertWorkflowsDisabled();
    validateConfig(cfg);
    if (cfg.storageMode !== 'shared' || cfg.root !== 'Gutachten_Mehrbenutzer_TEST_20261004' ||
        cfg.driveId !== 'b!kac40iyIuUq2m1cyr9YBjiQ3PrlENNBFnhADDVWqiVkQBRKI_bZ2Q40FFUr0Gkux' ||
        cfg.clientId !== '86f20e39-d4ae-4a7b-bf71-588de01d2e16' ||
        cfg.tenant !== 'fa2f2c93-cb81-46a9-ac8f-2c3e9efd5fd4') {
      throw new StorageError('TEST_FLOW_BINDING', 'Der Word-Testflow ist ausschließlich mit der geprüften gemeinsamen Testablage verbunden.');
    }
  }
  return { assertTestRoot, validateConfig, graphPath, collectPages, scopes, fingerprint, recordKey, createStore, StorageError, assertWorkflowsDisabled, assertWordTestWorkflow };
});
