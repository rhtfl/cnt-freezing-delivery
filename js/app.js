    // Настройки режимов и URL — в config/app.config.js (window.APP_CONFIG)
    const APP = window.APP_CONFIG;
    if (!APP || !APP.modes || typeof APP.modes !== 'object') {
      throw new Error('Подключите config/app.config.js перед js/app.js');
    }

    function buildRuntimeFromAppConfig() {
      const weekdays = { ...(APP.sheetNames && APP.sheetNames.weekdays ? APP.sheetNames.weekdays : {}) };
      const modeIds = Object.keys(APP.modes);
      const modeConfig = {};
      modeIds.forEach((id) => {
        const m = APP.modes[id];
        const auto = m.autoSelection || {};
        modeConfig[id] = {
          id,
          label: m.label || id,
          sheetEndpoint: m.sheetEndpoint || '',
          exportFilePrefix: m.exportFilePrefix || 'Заказы',
          sheetNameByDay: m.sheetNameByDay ? { ...m.sheetNameByDay } : { ...weekdays },
          rules: {
            priorityDepotId: auto.priorityDepotId ?? '',
            priorityStartLocationId: auto.priorityStartLocationId ?? '',
            defaultStartLocationId: auto.defaultStartLocationId ?? '',
            defaultExtraStartId: auto.defaultExtraStartId ?? '',
            defaultTimeWindow: m.defaultTimeWindow || '10:00-21:00'
          },
          seed: {
            startLocations: (m.seed && m.seed.startLocations) || [],
            vehicles: (m.seed && m.seed.vehicles) || [],
            depots: (m.seed && m.seed.depots) || []
          },
          startsEmptyUntilSheetSync: m.startsEmptyUntilSheetSync === true,
          defaultExportVehicleIds: Array.isArray(m.defaultExportVehicleIds)
            ? m.defaultExportVehicleIds.map((id) => String(id).trim()).filter(Boolean)
            : null
        };
      });
      const allMeta = APP.allMode || {};
      modeConfig.all = {
        id: 'all',
        label: allMeta.label || 'Все',
        sheetEndpoint: '',
        exportFilePrefix: allMeta.exportFilePrefix || 'Комбинированный_заказы',
        sheetNameByDay: { ...weekdays },
        rules: {
          priorityDepotId: '',
          priorityStartLocationId: '',
          defaultStartLocationId: '',
          defaultExtraStartId: '',
          defaultTimeWindow: allMeta.defaultTimeWindow || '10:00-21:00'
        },
        seed: { startLocations: [], vehicles: [], depots: [] },
        startsEmptyUntilSheetSync: false,
        defaultExportVehicleIds: null,
        isCombined: true
      };
      const defaultMode = APP.defaultMode && modeConfig[APP.defaultMode]
        ? APP.defaultMode
        : modeIds[0];
      return {
        DEFAULT_MODE: defaultMode,
        MODE_IDS: modeIds,
        SOURCE_MODE_IDS: modeIds,
        ALL_MODE: 'all',
        MODE_CONFIG: modeConfig,
        SHEET_NAME_BY_DAY: weekdays,
        SHEET_META: APP.sheetNames || {}
      };
    }

    const {
      DEFAULT_MODE,
      MODE_IDS,
      SOURCE_MODE_IDS,
      ALL_MODE,
      MODE_CONFIG,
      SHEET_NAME_BY_DAY,
      SHEET_META
    } = buildRuntimeFromAppConfig();

    // ===== Надёжная загрузка XLSX с запасными CDN =====
    function loadScript(src) {
      return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.async = true;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error('Failed ' + src));
        document.head.appendChild(script);
      });
    }

    function makeConfigField(label, key, value, options = {}) {
      const placeholderAttr = options.placeholder ? ` placeholder="${escapeHtml(options.placeholder)}"` : '';
      const dataAttr = `data-field="${key}"`;
      if (options.type === 'checkbox') {
        return `
          <div class="config-flag">
            <label>
              <input type="checkbox" ${dataAttr} ${boolFrom(value) ? 'checked' : ''} />
              <span>${escapeHtml(label)}</span>
            </label>
          </div>`;
      }
      const type = options.type === 'number' ? 'number' : 'text';
      const stepAttr = options.step ? ` step="${options.step}"` : '';
      const rawValue = value ?? '';
      const stringValue = rawValue === '' ? '' : String(rawValue);
      return `
        <div class="config-field">
          <label>${escapeHtml(label)}</label>
          <input type="${type}" ${dataAttr} value="${escapeHtml(stringValue)}"${placeholderAttr}${stepAttr} />
        </div>`;
    }

    async function ensureXlsxReady() {
      if (window.XLSX && XLSX.utils && XLSX.writeFile) return true;
      const tag = document.getElementById('xlsx-lib');
      if (tag && !window.XLSX) {
        await new Promise((res) => {
          tag.addEventListener('load', res, { once: true });
          tag.addEventListener('error', res, { once: true });
          setTimeout(res, 1500);
        });
        if (window.XLSX && XLSX.utils && XLSX.writeFile) return true;
      }
      try { await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js'); } catch (_) {}
      if (window.XLSX && XLSX.utils && XLSX.writeFile) return true;
      try { await loadScript('https://unpkg.com/xlsx@0.19.3/dist/xlsx.full.min.js'); } catch (_) {}
      return !!(window.XLSX && XLSX.utils && XLSX.writeFile);
    }

    function pickValue(source, ...keys) {
      if (!source) return '';
      for (const key of keys) {
        if (key in source && source[key] != null && source[key] !== '') {
          return source[key];
        }
      }
      return '';
    }

    // ===== Константы и состояние =====
    const WEEKDAY_LABELS = {
      monday: 'Понедельник',
      tuesday: 'Вторник',
      wednesday: 'Среда',
      thursday: 'Четверг',
      friday: 'Пятница',
      saturday: 'Суббота',
      sunday: 'Воскресенье'
    };
    const DAY_KEYS = Object.keys(WEEKDAY_LABELS);
    const CANON_TO_JS_DAY = {
      monday: 1,
      tuesday: 2,
      wednesday: 3,
      thursday: 4,
      friday: 5,
      saturday: 6,
      sunday: 0
    };
    function buildSheetNameToCanon(sheetNameByDay) {
      return Object.entries(sheetNameByDay).reduce((acc, [canon, sheet]) => {
        acc[String(sheet).toLowerCase()] = canon;
        acc[canon.toLowerCase()] = canon;
        return acc;
      }, {});
    }

    function createEmptySchedule() {
      return DAY_KEYS.reduce((acc, key) => {
        acc[key] = [];
        return acc;
      }, {});
    }

    function createEmptyExtraOrders() {
      return DAY_KEYS.reduce((acc, key) => {
        acc[key] = [];
        return acc;
      }, {});
    }

    function emptyStores() {
      return {
        scheduleData: createEmptySchedule(),
        extraOrders: createEmptyExtraOrders(),
        vehicles: [],
        depots: [],
        startLocations: [],
        selected: new Set(),
        selectedVehicleIds: new Set()
      };
    }

    function vehicleUid(row, index) {
      const id = row && row.id != null ? String(row.id).trim() : '';
      return id ? `veh:id::${id}` : `veh:idx::${index}`;
    }

    function getSelectedVehicleIds() {
      return getActiveStores().selectedVehicleIds;
    }

    function getVehiclesForExport() {
      if (isAllMode()) {
        return getMergedUniqueVehicles()
          .filter((entry) => isVehicleExportSelectedInAllMode(entry))
          .map((entry) => entry.row);
      }
      const store = getActiveStores();
      const rows = store.vehicles || [];
      const selected = store.selectedVehicleIds;
      return rows.filter((row, index) => selected.has(vehicleUid(row, index)));
    }

    function getDepotsForExport() {
      if (!isAllMode()) return dataStore.depots || [];
      return getMergedUniqueDepotsWithMeta().map((entry) => entry.depot);
    }

    function countConfiguredVehicles() {
      if (!isAllMode()) return (dataStore.vehicles || []).length;
      return getMergedUniqueVehicles().length;
    }

    function pruneVehicleSelection() {
      const store = getActiveStores();
      const rows = dataStore.vehicles || [];
      const valid = new Set(rows.map((row, index) => vehicleUid(row, index)));
      const next = new Set();
      store.selectedVehicleIds.forEach((uid) => {
        if (valid.has(uid)) next.add(uid);
      });
      store.selectedVehicleIds = next;
    }

    function setAllVehiclesSelected(selectAll) {
      if (isAllMode()) {
        setAllUniqueVehiclesExportSelected(selectAll);
        return;
      }
      const store = getActiveStores();
      const rows = dataStore.vehicles || [];
      if (selectAll) {
        store.selectedVehicleIds = new Set(rows.map((row, index) => vehicleUid(row, index)));
      } else {
        store.selectedVehicleIds = new Set();
      }
      saveLocal();
    }

    function applyDefaultVehicleExportSelection(mode) {
      const modeId = mode && MODE_CONFIG[mode] ? mode : getActiveMode();
      if (isAllMode(modeId)) return;
      const store = storesByMode[modeId] || emptyStores();
      const rows = store.vehicles || [];
      const defaultIds = getModeConfig(modeId).defaultExportVehicleIds;

      if (Array.isArray(defaultIds)) {
        const idSet = new Set(defaultIds);
        const next = new Set();
        rows.forEach((row, index) => {
          const vid = row && row.id != null ? String(row.id).trim() : '';
          if (vid && idSet.has(vid)) {
            next.add(vehicleUid(row, index));
          }
        });
        store.selectedVehicleIds = next;
      } else {
        store.selectedVehicleIds = new Set(rows.map((row, index) => vehicleUid(row, index)));
      }
      if (modeId === getActiveMode()) saveLocal();
    }

    function rebuildVehicleSelectionAfterLoad() {
      applyDefaultVehicleExportSelection(getActiveMode());
    }

    function ensureVehicleSelectionDefault() {
      if (!shouldLoadPersistedForMode(getActiveMode())) return;
      if ((dataStore.vehicles || []).length && getSelectedVehicleIds().size === 0) {
        applyDefaultVehicleExportSelection(getActiveMode());
      }
    }

    function updateVehicleCountBadge() {
      if (!dom.vehCount) return;
      const total = countConfiguredVehicles();
      const selected = getVehiclesForExport().length;
      if (!total) {
        dom.vehCount.textContent = '0 записей';
        return;
      }
      dom.vehCount.textContent = `${total} записей · в экспорт: ${selected}`;
    }

    const storesByMode = {
      horeca: emptyStores(),
      gallery: emptyStores(),
      all: emptyStores()
    };

    function getActiveMode() {
      return state.activeMode;
    }

    function isAllMode(mode) {
      return (mode || getActiveMode()) === ALL_MODE;
    }

    function isSourceMode(mode) {
      return SOURCE_MODE_IDS.includes(mode || getActiveMode());
    }

    function isValidBusinessMode(mode) {
      return Boolean(mode && MODE_CONFIG[mode]);
    }

    function getModeConfig(mode) {
      const id = mode && MODE_CONFIG[mode] ? mode : state.activeMode;
      return MODE_CONFIG[id] || MODE_CONFIG[DEFAULT_MODE];
    }

    function getStoreForMode(mode) {
      const id = mode || getActiveMode();
      if (!storesByMode[id]) storesByMode[id] = emptyStores();
      return storesByMode[id];
    }

    function getActiveStores() {
      return getStoreForMode(getActiveMode());
    }

    function isStartSelectionUid(uid) {
      if (typeof uid !== 'string') return false;
      return uid.startsWith('shared::start::') || uid.startsWith('start::') || uid.includes('::start::');
    }

    function formatSourceLabels(labels) {
      return (labels || []).filter(Boolean).join(' · ');
    }

    function startDedupeKey(record) {
      const startId = normalizeStartId(record && record.id);
      if (startId) return `id::${startId}`;
      const lat = toNumOrNull(record['point.lat']);
      const lon = toNumOrNull(record['point.lon']);
      if (lat != null && lon != null) return `coord::${lat}::${lon}`;
      return '';
    }

    function sharedStartUid(key, record) {
      const startId = normalizeStartId(record && record.id);
      if (startId) return `shared::start::${startId}`;
      return `shared::start::${key}`;
    }

    function getMergedUniqueStartLocations() {
      const byKey = new Map();
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!shouldLoadPersistedForMode(sourceMode)) return;
        const label = getModeConfig(sourceMode).label;
        (getStoreForMode(sourceMode).startLocations || []).forEach((record) => {
          const key = startDedupeKey(record);
          if (!key) return;
          if (!byKey.has(key)) {
            byKey.set(key, { record, sourceModes: [sourceMode], sourceLabels: [label] });
          } else {
            const entry = byKey.get(key);
            if (!entry.sourceModes.includes(sourceMode)) {
              entry.sourceModes.push(sourceMode);
              entry.sourceLabels.push(label);
            }
          }
        });
      });
      return Array.from(byKey.entries()).map(([key, data]) => ({
        key,
        ...data,
        uid: sharedStartUid(key, data.record)
      }));
    }

    function getMergedUniqueVehicles() {
      const byKey = new Map();
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!shouldLoadPersistedForMode(sourceMode)) return;
        const label = getModeConfig(sourceMode).label;
        (getStoreForMode(sourceMode).vehicles || []).forEach((row, index) => {
          const id = row && row.id != null ? String(row.id).trim() : '';
          const key = id ? `veh:id::${id}` : `veh:anon::${sourceMode}::${index}`;
          if (!byKey.has(key)) {
            byKey.set(key, { row, key, sourceModes: [sourceMode], sourceLabels: [label] });
          } else {
            const entry = byKey.get(key);
            if (!entry.sourceModes.includes(sourceMode)) {
              entry.sourceModes.push(sourceMode);
              entry.sourceLabels.push(label);
            }
          }
        });
      });
      return Array.from(byKey.values());
    }

    function getMergedUniqueDepotsWithMeta() {
      const byKey = new Map();
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!shouldLoadPersistedForMode(sourceMode)) return;
        const label = getModeConfig(sourceMode).label;
        (getStoreForMode(sourceMode).depots || []).forEach((depot) => {
          const id = depot && depot.id != null ? String(depot.id).trim() : '';
          const key = id || `depot:anon::${sourceMode}::${JSON.stringify(depot)}`;
          if (!byKey.has(key)) {
            byKey.set(key, { depot, sourceLabels: [label] });
          } else {
            const entry = byKey.get(key);
            if (!entry.sourceLabels.includes(label)) entry.sourceLabels.push(label);
          }
        });
      });
      return Array.from(byKey.values());
    }

    function isVehicleExportSelectedInAllMode(entry) {
      const id = entry.row && entry.row.id != null ? String(entry.row.id).trim() : '';
      if (id) {
        return SOURCE_MODE_IDS.some((sourceMode) => {
          const store = getStoreForMode(sourceMode);
          return (store.vehicles || []).some((row, index) => {
            const vid = row && row.id != null ? String(row.id).trim() : '';
            if (vid !== id) return false;
            return store.selectedVehicleIds.has(vehicleUid(row, index));
          });
        });
      }
      return entry.sourceModes.some((sourceMode) => {
        const store = getStoreForMode(sourceMode);
        return (store.vehicles || []).some((row, index) => (
          vehicleUid(row, index) === entry.key && store.selectedVehicleIds.has(entry.key)
        ));
      });
    }

    function setVehicleExportSelectedInAllMode(entry, selected) {
      const id = entry.row && entry.row.id != null ? String(entry.row.id).trim() : '';
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        const store = getStoreForMode(sourceMode);
        (store.vehicles || []).forEach((row, index) => {
          const vid = row && row.id != null ? String(row.id).trim() : '';
          if (id) {
            if (vid !== id) return;
          } else if (vehicleUid(row, index) !== entry.key) {
            return;
          }
          const uid = vehicleUid(row, index);
          if (selected) store.selectedVehicleIds.add(uid);
          else store.selectedVehicleIds.delete(uid);
        });
        saveLocalForMode(sourceMode);
      });
    }

    function setAllUniqueVehiclesExportSelected(selectAll) {
      getMergedUniqueVehicles().forEach((entry) => {
        setVehicleExportSelectedInAllMode(entry, selectAll);
      });
    }

    function forEachMatchingStartRecord(entry, callback) {
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!entry.sourceModes.includes(sourceMode)) return;
        const store = getStoreForMode(sourceMode);
        (store.startLocations || []).forEach((record, index) => {
          if (startDedupeKey(record) !== entry.key) return;
          callback(record, index, sourceMode, store);
        });
      });
    }

    function setMergedStartField(entry, key, value) {
      entry.record[key] = value;
      forEachMatchingStartRecord(entry, (record) => {
        record[key] = value;
      });
      entry.sourceModes.forEach((mode) => saveLocalForMode(mode));
    }

    function deleteMergedStart(entry) {
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!entry.sourceModes.includes(sourceMode)) return;
        const store = getStoreForMode(sourceMode);
        store.startLocations = (store.startLocations || []).filter(
          (record) => startDedupeKey(record) !== entry.key
        );
        saveLocalForMode(sourceMode);
      });
      state.selected.delete(entry.uid);
      renderStartLocations();
      render();
    }

    function forEachMatchingVehicleRecord(entry, callback) {
      const id = entry.row && entry.row.id != null ? String(entry.row.id).trim() : '';
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!entry.sourceModes.includes(sourceMode)) return;
        const store = getStoreForMode(sourceMode);
        (store.vehicles || []).forEach((row, index) => {
          if (id) {
            const vid = row && row.id != null ? String(row.id).trim() : '';
            if (vid !== id) return;
          } else if (vehicleUid(row, index) !== entry.key) {
            return;
          }
          callback(row, index, sourceMode, store);
        });
      });
    }

    function setMergedVehicleField(entry, key, value, inputType) {
      entry.row[key] = value;
      forEachMatchingVehicleRecord(entry, (row, index, sourceMode, store) => {
        const prevUid = vehicleUid(row, index);
        const wasSelected = store.selectedVehicleIds.has(prevUid);
        if (inputType === 'checkbox') {
          row[key] = value;
        } else if (inputType === 'number') {
          row[key] = value === '' ? '' : Number(value);
        } else {
          row[key] = value;
        }
        if (key === 'id' && wasSelected) {
          store.selectedVehicleIds.delete(prevUid);
          store.selectedVehicleIds.add(vehicleUid(row, index));
        }
      });
      if (key === 'id') {
        const newId = value != null ? String(value).trim() : '';
        if (newId) entry.key = `veh:id::${newId}`;
      }
      entry.sourceModes.forEach((mode) => saveLocalForMode(mode));
    }

    function deleteMergedVehicle(entry) {
      const id = entry.row && entry.row.id != null ? String(entry.row.id).trim() : '';
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!entry.sourceModes.includes(sourceMode)) return;
        const store = getStoreForMode(sourceMode);
        if (id) {
          store.vehicles = (store.vehicles || []).filter((row) => {
            const vid = row && row.id != null ? String(row.id).trim() : '';
            return vid !== id;
          });
        } else {
          store.vehicles = (store.vehicles || []).filter((row, index) => (
            vehicleUid(row, index) !== entry.key
          ));
        }
        const valid = new Set((store.vehicles || []).map((row, index) => vehicleUid(row, index)));
        const next = new Set();
        store.selectedVehicleIds.forEach((uid) => {
          if (valid.has(uid)) next.add(uid);
        });
        store.selectedVehicleIds = next;
        saveLocalForMode(sourceMode);
      });
      renderVehicles();
    }

    function addStartToAllSources() {
      const record = {
        id: '',
        ref: '',
        'point.lat': '',
        'point.lon': '',
        time_window: '07:00:00-22:00:00',
        type: 'garage',
        address: '',
        comments: '',
        phone: ''
      };
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!shouldLoadPersistedForMode(sourceMode)) return;
        const store = getStoreForMode(sourceMode);
        if (!store.startLocations) store.startLocations = [];
        store.startLocations.push({ ...record });
        saveLocalForMode(sourceMode);
      });
      renderStartLocations();
      render();
    }

    function addVehicleToAllSources() {
      const record = {
        id: '',
        ref: '',
        'capacity.weight_kg': '',
        tags: '',
        start_at: '',
        finish_at: '',
        visit_depot_at_start: false,
        return_to_depot: false,
        depot_id: '',
        'shifts.0.id': '',
        'shifts.0.time_window': '',
        allow_different_depots_in_route: false,
        max_middle_depots: '',
        depots_only_at_run_beginning: false,
        starting_depot_id: '',
        middle_depot_id: ''
      };
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!shouldLoadPersistedForMode(sourceMode)) return;
        const store = getStoreForMode(sourceMode);
        if (!store.vehicles) store.vehicles = [];
        store.vehicles.push({ ...record });
        const idx = store.vehicles.length - 1;
        store.selectedVehicleIds.add(vehicleUid(store.vehicles[idx], idx));
        saveLocalForMode(sourceMode);
      });
      renderVehicles();
    }

    function getModeRules() {
      return getModeConfig().rules || {};
    }

    function getScheduleData() {
      return getActiveStores().scheduleData;
    }

    function setScheduleData(next) {
      const store = getActiveStores();
      if (next && typeof next === 'object') {
        store.scheduleData = next;
      } else {
        store.scheduleData = createEmptySchedule();
      }
    }

    function getDefaultTimeWindow() {
      const tw = getModeRules().defaultTimeWindow;
      return tw || '10:00-21:00';
    }

    function migrateLegacyLocalStorage() {
      try {
        const legacyVehicles = localStorage.getItem('vehicles');
        const legacyDepots = localStorage.getItem('depots');
        if (legacyVehicles && !localStorage.getItem('vehicles_gallery')) {
          localStorage.setItem('vehicles_gallery', legacyVehicles);
          localStorage.removeItem('vehicles');
        }
        if (legacyDepots && !localStorage.getItem('depots_gallery')) {
          localStorage.setItem('depots_gallery', legacyDepots);
          localStorage.removeItem('depots');
        }
      } catch (_) {}
    }

    /** Однократно сбрасывает кэш HoReCa (убирает тестовые данные из старых версий). */
    function migrateHorecaStaleCache() {
      try {
        if (localStorage.getItem('horeca_fresh_store_v1') === '1') return;
        clearModePersistedData('horeca');
        localStorage.setItem('horeca_fresh_store_v1', '1');
      } catch (_) {}
    }

    function storageKey(suffix, mode) {
      return `${suffix}_${mode != null ? mode : getActiveMode()}`;
    }

    function sheetSyncedStorageKey(mode) {
      return `sheetSynced_${mode || getActiveMode()}`;
    }

    function sheetCacheStorageKey(mode) {
      return `sheetImportCache_${mode || getActiveMode()}`;
    }

    function getSheetCacheTtlMs() {
      const minutes = Number(APP.sheetCacheTtlMinutes);
      return (Number.isFinite(minutes) && minutes > 0 ? minutes : 30) * 60 * 1000;
    }

    let sheetCacheUiTimer = null;

    function readSheetCache(mode) {
      try {
        const raw = localStorage.getItem(sheetCacheStorageKey(mode));
        if (!raw) return null;
        return JSON.parse(raw);
      } catch (_) {
        return null;
      }
    }

    function isSheetCacheValid(cache) {
      if (!cache || typeof cache.importedAt !== 'number') return false;
      return Date.now() - cache.importedAt < getSheetCacheTtlMs();
    }

    function hasValidSheetCache(mode) {
      return isSheetCacheValid(readSheetCache(mode));
    }

    function getSheetCacheImportedAt(mode) {
      const cache = readSheetCache(mode);
      return isSheetCacheValid(cache) ? cache.importedAt : null;
    }

    function normalizeScheduleCache(schedule) {
      const base = createEmptySchedule();
      if (!schedule || typeof schedule !== 'object') return base;
      DAY_KEYS.forEach((day) => {
        base[day] = (schedule[day] || []).map((entry) => normalizeOrderRecord(entry)).filter(Boolean);
      });
      return base;
    }

    function applySheetCacheToStore(mode, cache) {
      const store = getStoreForMode(mode);
      store.scheduleData = normalizeScheduleCache(cache.scheduleData);
      if (Array.isArray(cache.vehicles)) {
        store.vehicles = cache.vehicles.map((item) => normalizeVehicleRecord(item)).filter(Boolean);
      }
      if (Array.isArray(cache.depots)) {
        store.depots = cache.depots.map((item) => normalizeDepotRecord(item)).filter(Boolean);
      }
      if (Array.isArray(cache.startLocations)) {
        store.startLocations = cache.startLocations.map((item) => normalizeStartRecord(item)).filter(Boolean);
      }
      store.selected = new Set(Array.isArray(cache.selected) ? cache.selected.map(String) : []);
      if (Array.isArray(cache.selectedVehicleIds)) {
        store.selectedVehicleIds = new Set(cache.selectedVehicleIds.map(String));
      }
      const prevMode = state.activeMode;
      state.activeMode = mode;
      pruneVehicleSelection();
      state.activeMode = prevMode;
    }

    function clearSheetImportCache(mode) {
      try {
        localStorage.removeItem(sheetCacheStorageKey(mode || getActiveMode()));
      } catch (_) {}
    }

    function allModeSelectionCacheKey() {
      return 'sheetImportCache_all_selection';
    }

    function persistAllModeSelectionCache() {
      if (!isAllMode()) return;
      const importedAt = Math.min(...SOURCE_MODE_IDS.map(getSheetCacheImportedAt).filter(Boolean));
      if (!importedAt) return;
      try {
        localStorage.setItem(allModeSelectionCacheKey(), JSON.stringify({
          importedAt,
          selected: Array.from(state.selected)
        }));
      } catch (_) {}
      updateSheetCacheUi();
    }

    function loadAllModeSelectionCache() {
      if (!isAllMode()) return false;
      try {
        const raw = localStorage.getItem(allModeSelectionCacheKey());
        if (!raw) return false;
        const parsed = JSON.parse(raw);
        if (!isSheetCacheValid(parsed)) {
          localStorage.removeItem(allModeSelectionCacheKey());
          return false;
        }
        getStoreForMode(ALL_MODE).selected = new Set(
          Array.isArray(parsed.selected) ? parsed.selected.map(String) : []
        );
        return true;
      } catch (_) {
        return false;
      }
    }

    function clearAllModeSelectionCache() {
      try {
        localStorage.removeItem(allModeSelectionCacheKey());
      } catch (_) {}
    }

    function saveSheetImportCache(mode, options = {}) {
      const id = mode || getActiveMode();
      if (isAllMode(id)) return;
      const store = getStoreForMode(id);
      const prev = readSheetCache(id);
      const preserveImportedAt = options.preserveImportedAt !== false && !options.freshImport && isSheetCacheValid(prev);
      const payload = {
        importedAt: preserveImportedAt ? prev.importedAt : Date.now(),
        scheduleData: store.scheduleData,
        vehicles: store.vehicles,
        depots: store.depots,
        startLocations: store.startLocations,
        selected: Array.from(store.selected),
        selectedVehicleIds: Array.from(store.selectedVehicleIds)
      };
      try {
        localStorage.setItem(sheetCacheStorageKey(id), JSON.stringify(payload));
        markSheetSynced(id);
      } catch (_) {}
      updateSheetCacheUi();
    }

    function loadSheetImportCache(mode) {
      const id = mode || getActiveMode();
      if (isAllMode(id)) return false;
      const cache = readSheetCache(id);
      if (!isSheetCacheValid(cache)) {
        if (cache) clearSheetImportCache(id);
        return false;
      }
      applySheetCacheToStore(id, cache);
      markSheetSynced(id);
      return true;
    }

    function loadExtraOrdersOnly(mode) {
      try {
        const storedExtra = localStorage.getItem(storageKey('extraOrders', mode));
        if (storedExtra) {
          getStoreForMode(mode).extraOrders = normalizeExtraOrdersStructure(JSON.parse(storedExtra));
        }
      } catch (_) {}
    }

    function loadLegacyReferenceData(mode) {
      const store = getStoreForMode(mode);
      let loaded = false;
      try {
        const storedVehicles = localStorage.getItem(storageKey('vehicles', mode));
        const storedDepots = localStorage.getItem(storageKey('depots', mode));
        const storedStart = localStorage.getItem(storageKey('startLocations', mode));
        if (storedVehicles) {
          const parsed = JSON.parse(storedVehicles);
          if (Array.isArray(parsed)) {
            store.vehicles = parsed.map((item) => normalizeVehicleRecord(item)).filter(Boolean);
            loaded = loaded || store.vehicles.length > 0;
          }
        }
        if (storedDepots) {
          const parsed = JSON.parse(storedDepots);
          if (Array.isArray(parsed)) {
            store.depots = parsed.map((item) => normalizeDepotRecord(item)).filter(Boolean);
            loaded = loaded || store.depots.length > 0;
          }
        }
        if (storedStart) {
          const parsed = JSON.parse(storedStart);
          if (Array.isArray(parsed)) {
            store.startLocations = parsed.map((item) => normalizeStartRecord(item)).filter(Boolean);
            loaded = loaded || store.startLocations.length > 0;
          }
        }
        const storedVehicleSelection = localStorage.getItem(storageKey('selectedVehicleIds', mode));
        if (storedVehicleSelection) {
          const parsed = JSON.parse(storedVehicleSelection);
          if (Array.isArray(parsed)) {
            store.selectedVehicleIds = new Set(parsed.map((id) => String(id)));
          }
        }
      } catch (_) {}
      pruneVehicleSelection();
      return loaded;
    }

    function persistActiveSheetCacheSelection() {
      if (isAllMode()) {
        persistAllModeSelectionCache();
        return;
      }
      if (hasValidSheetCache(getActiveMode())) {
        saveSheetImportCache(getActiveMode(), { preserveImportedAt: true });
      }
    }

    function formatSheetCacheAge(importedAt) {
      const minutes = Math.floor((Date.now() - importedAt) / 60000);
      if (minutes <= 0) return 'только что';
      if (minutes === 1) return '1 мин назад';
      if (minutes < 60) return `${minutes} мин назад`;
      const hours = Math.floor(minutes / 60);
      const rem = minutes % 60;
      if (rem === 0) return `${hours} ч назад`;
      return `${hours} ч ${rem} мин назад`;
    }

    function formatSheetCacheRemaining(importedAt) {
      const remainingMs = getSheetCacheTtlMs() - (Date.now() - importedAt);
      if (remainingMs <= 0) return '';
      const minutes = Math.ceil(remainingMs / 60000);
      if (minutes <= 1) return ' · осталось <1 мин';
      return ` · ещё ~${minutes} мин`;
    }

    function getSheetCacheUiMeta() {
      if (isAllMode()) {
        const entries = SOURCE_MODE_IDS.map((modeId) => {
          const importedAt = getSheetCacheImportedAt(modeId);
          if (!importedAt) return null;
          return `${getModeConfig(modeId).label}: ${formatSheetCacheAge(importedAt)}`;
        }).filter(Boolean);
        if (!entries.length) return null;
        const importedAt = Math.min(...SOURCE_MODE_IDS.map(getSheetCacheImportedAt).filter(Boolean));
        return {
          importedAt,
          text: `Локальная копия таблицы · ${entries.join(' · ')}${formatSheetCacheRemaining(importedAt)}`
        };
      }
      const importedAt = getSheetCacheImportedAt(getActiveMode());
      if (!importedAt) return null;
      return {
        importedAt,
        text: `Локальная копия таблицы · ${formatSheetCacheAge(importedAt)}${formatSheetCacheRemaining(importedAt)}`
      };
    }

    function updateLoadSheetButtonCacheHint(meta) {
      const btn = dom.loadSheetBtn;
      if (!btn) return;
      let sub = btn.querySelector('.btn-action__sub');
      if (!meta) {
        sub?.remove();
        btn.classList.remove('btn-action--cached', 'sheet-cache-stale');
        btn.removeAttribute('data-cache-age');
        return;
      }
      if (!sub) {
        sub = document.createElement('span');
        sub.className = 'btn-action__sub';
        btn.appendChild(sub);
      }
      const ageText = formatSheetCacheAge(meta.importedAt);
      sub.textContent = ageText;
      btn.classList.add('btn-action--cached');
      btn.dataset.cacheAge = ageText;
      const stale = Date.now() - meta.importedAt >= getSheetCacheTtlMs() * 0.66;
      btn.classList.toggle('sheet-cache-stale', stale);
    }

    function updateSheetCacheUi() {
      const meta = getSheetCacheUiMeta();
      if (dom.sheetCacheStatus) {
        if (!meta) {
          dom.sheetCacheStatus.hidden = true;
          dom.sheetCacheStatus.textContent = '';
          dom.sheetCacheStatus.classList.remove('sheet-cache-status--stale');
        } else {
          dom.sheetCacheStatus.hidden = false;
          dom.sheetCacheStatus.textContent = meta.text;
          dom.sheetCacheStatus.classList.toggle(
            'sheet-cache-status--stale',
            Date.now() - meta.importedAt >= getSheetCacheTtlMs() * 0.66
          );
        }
      }
      if (dom.loadSheetBtn && !dom.loadSheetBtn.classList.contains('btn-action--loading')) {
        dom.loadSheetBtn.title = meta
          ? `${meta.text}. Нажмите, чтобы загрузить свежие данные из Google Sheets.`
          : 'Загрузить расписание и справочники из Google Sheets';
        updateLoadSheetButtonCacheHint(meta);
      }
    }

    function handleSheetCacheExpired(modeId) {
      clearSheetImportCache(modeId);
      clearAllModeSelectionCache();
      try {
        localStorage.removeItem(sheetSyncedStorageKey(modeId));
      } catch (_) {}
      loadExtraOrdersOnly(modeId);
      if (modeStartsEmptyUntilSync(modeId)) {
        resetModeStore(modeId);
        loadExtraOrdersOnly(modeId);
        getStoreForMode(modeId).scheduleData = createEmptySchedule();
        getStoreForMode(modeId).selected = new Set();
      } else {
        loadLegacyReferenceData(modeId);
        getStoreForMode(modeId).scheduleData = createEmptySchedule();
        getStoreForMode(modeId).selected = new Set();
      }
      if (getActiveMode() === modeId || isAllMode()) {
        showNotify(
          `Локальная копия «${getModeConfig(modeId).label}» устарела — обновите данные из Google Sheets.`,
          'info',
          8000
        );
        refreshModeUiAfterDataChange();
      }
    }

    function expireStaleSheetCaches() {
      SOURCE_MODE_IDS.forEach((modeId) => {
        const cache = readSheetCache(modeId);
        if (!cache || isSheetCacheValid(cache)) return;
        handleSheetCacheExpired(modeId);
      });
    }

    function startSheetCacheUiTimer() {
      if (sheetCacheUiTimer) clearInterval(sheetCacheUiTimer);
      sheetCacheUiTimer = setInterval(() => {
        expireStaleSheetCaches();
        updateSheetCacheUi();
      }, 30000);
    }

    function hasSheetSynced(mode) {
      const id = mode || getActiveMode();
      try {
        return localStorage.getItem(sheetSyncedStorageKey(id)) === '1';
      } catch (_) {
        return false;
      }
    }

    function markSheetSynced(mode) {
      try {
        localStorage.setItem(sheetSyncedStorageKey(mode || getActiveMode()), '1');
      } catch (_) {}
    }

    function modeStartsEmptyUntilSync(mode) {
      if (isAllMode(mode)) return false;
      const cfg = getModeConfig(mode);
      return cfg.startsEmptyUntilSheetSync === true;
    }

    function hasLoadedSheetDataForMode(mode) {
      const store = getStoreForMode(mode);
      if ((store.startLocations || []).length > 0) return true;
      if ((store.vehicles || []).length > 0) return true;
      return DAY_KEYS.some((day) => (store.scheduleData[day] || []).length > 0);
    }

    function shouldLoadPersistedForMode(mode) {
      const id = mode || getActiveMode();
      if (isAllMode(id)) {
        return SOURCE_MODE_IDS.some((sourceId) => shouldLoadPersistedForMode(sourceId));
      }
      if (!modeStartsEmptyUntilSync(id)) return true;
      return hasValidSheetCache(id) || hasLoadedSheetDataForMode(id);
    }

    function resetModeStore(mode) {
      const id = mode || getActiveMode();
      storesByMode[id] = emptyStores();
    }

    function clearModePersistedData(mode) {
      const id = mode || getActiveMode();
      ['vehicles', 'depots', 'startLocations', 'selectedVehicleIds', 'extraOrders'].forEach((suffix) => {
        try {
          localStorage.removeItem(`${suffix}_${id}`);
        } catch (_) {}
      });
      try {
        localStorage.removeItem(sheetSyncedStorageKey(id));
      } catch (_) {}
      clearSheetImportCache(id);
    }

    function ensureEmptyModeUntilSheetSync(mode) {
      if (isAllMode(mode)) return;
      if (!modeStartsEmptyUntilSync(mode)) return;
      if (hasValidSheetCache(mode)) return;
      if (hasLoadedSheetDataForMode(mode)) return;
      loadExtraOrdersOnly(mode);
      const extraOrders = getStoreForMode(mode).extraOrders;
      clearSheetImportCache(mode);
      ['vehicles', 'depots', 'startLocations', 'selectedVehicleIds'].forEach((suffix) => {
        try {
          localStorage.removeItem(`${suffix}_${mode}`);
        } catch (_) {}
      });
      try {
        localStorage.removeItem(sheetSyncedStorageKey(mode));
      } catch (_) {}
      resetModeStore(mode);
      getStoreForMode(mode).extraOrders = extraOrders || createEmptyExtraOrders();
      getStoreForMode(mode).scheduleData = createEmptySchedule();
      if (getActiveMode() === mode) {
        state.selected.clear();
        state.query = '';
      }
    }

    function loadLocalForMode(mode) {
      const prev = state.activeMode;
      state.activeMode = mode;
      loadLocal();
      state.activeMode = prev;
    }

    function saveLocalForMode(mode) {
      const prev = state.activeMode;
      state.activeMode = mode;
      saveLocal();
      state.activeMode = prev;
    }

    function sheetNameMatches(lower, names) {
      const list = Array.isArray(names) ? names : [names];
      return list.some((name) => String(name).trim().toLowerCase() === lower);
    }

    function applyConfigSeedIfEmpty() {
      if (!shouldLoadPersistedForMode(getActiveMode())) return;
      const seed = getModeConfig().seed;
      if (!seed) return;
      if (!dataStore.vehicles.length && seed.vehicles.length) {
        dataStore.vehicles = seed.vehicles.map((item) => normalizeVehicleRecord(item)).filter(Boolean);
      }
      if (!dataStore.depots.length && seed.depots.length) {
        dataStore.depots = seed.depots.map((item) => normalizeDepotRecord(item)).filter(Boolean);
      }
      if (!dataStore.startLocations.length && seed.startLocations.length) {
        dataStore.startLocations = seed.startLocations.map((item) => normalizeStartRecord(item)).filter(Boolean);
      }
    }

    function applyConfigSeedForAllModes() {
      const prev = state.activeMode;
      SOURCE_MODE_IDS.forEach((modeId) => {
        state.activeMode = modeId;
        if (shouldLoadPersistedForMode(modeId)) applyConfigSeedIfEmpty();
      });
      state.activeMode = prev;
    }

    const EXTRA_ORDER_FORM_FIELDS = [
      { key: 'title', label: 'Наименование клиента', placeholder: 'например Магазин на Ленина', required: true },
      { key: 'address', label: 'Адрес', placeholder: 'полный адрес доставки' },
      { key: 'coords_paste', label: 'Координаты с карты', placeholder: '55.706284, 37.781800' },
      { key: 'lat', label: 'Широта', type: 'number', step: 'any', placeholder: '55.751244', required: true },
      { key: 'lng', label: 'Долгота', type: 'number', step: 'any', placeholder: '37.618423', required: true },
      { key: 'time_window', label: 'Временное окно', placeholder: '10:00-21:00' },
      { key: 'delivery_minutes', label: 'Время обслуживания, мин', type: 'number', step: '1', placeholder: 'например 10' },
      { key: 'phone', label: 'Телефон', placeholder: '+7 (___) ___-__-__', phoneMask: true },
      { key: 'depot_id', label: 'ID склада', placeholder: 'если нужно' },
      { key: 'depot_name', label: 'Название склада', placeholder: 'если нужно' },
      { key: 'comments', label: 'Комментарий', placeholder: 'заметка для водителя' }
    ];

    let extraOrderModalScrollY = 0;

    function lockPageScrollForExtraOrderModal() {
      extraOrderModalScrollY = window.scrollY || document.documentElement.scrollTop || 0;
      document.body.classList.add('extra-order-modal-open');
      document.body.style.top = `-${extraOrderModalScrollY}px`;
      document.body.style.position = 'fixed';
      document.body.style.width = '100%';
    }

    function unlockPageScrollForExtraOrderModal() {
      document.body.classList.remove('extra-order-modal-open');
      document.body.style.position = '';
      document.body.style.top = '';
      document.body.style.width = '';
      window.scrollTo(0, extraOrderModalScrollY);
    }

    const START_LOCATIONS_FIELDS = [
      { key: 'id', label: 'ID' },
      { key: 'ref', label: 'Название' },
      { key: 'point.lat', label: 'Широта', type: 'number', step: 'any' },
      { key: 'point.lon', label: 'Долгота', type: 'number', step: 'any' },
      { key: 'time_window', label: 'Временное окно', placeholder: 'например 07:00:00-22:00:00' },
      { key: 'type', label: 'Тип (garage/...)', placeholder: 'например garage' },
      { key: 'comments', label: 'Комментарий', placeholder: 'например Комментарий к точке' }
    ];

    const VEHICLE_FIELD_CONFIG = [
      { key: 'id', label: 'ID (Идентификатор машины)', placeholder: 'например V-1' },
      { key: 'ref', label: 'Имя курьера', placeholder: 'например Иван' },
      { key: 'capacity.weight_kg', label: 'Грузоподъемность, кг', type: 'number', step: 'any', placeholder: 'например 1000' },
      { key: 'tags', label: 'Свойства', placeholder: 'например изотерм, GPS' },
      { key: 'start_at', label: 'Начальная точка', placeholder: 'например depot:1' },
      { key: 'finish_at', label: 'Конечная точка', placeholder: 'например depot:2' },
      { key: 'depot_id', label: 'Идентификаторы складов', placeholder: 'список id через запятую' },
      { key: 'shifts.0.id', label: 'Смена 1. ID', placeholder: 'например shift:1' },
      { key: 'shifts.0.time_window', label: 'Смена 1. Временное окно', placeholder: 'например 09:00-18:00' },
      { key: 'max_middle_depots', label: 'Макс. количество промежуточных складов', type: 'number', step: '1', placeholder: 'например 2' },
      { key: 'starting_depot_id', label: 'Стартовый склад', placeholder: 'например depot:1' },
      { key: 'middle_depot_id', label: 'Промежуточный склад', placeholder: 'например depot:2' }
    ];

    const VEHICLE_FLAG_CONFIG = [
      { key: 'visit_depot_at_start', label: 'Посетить склад в начале рейса' },
      { key: 'return_to_depot', label: 'Надо вернуться на склад, Да/Нет' },
      { key: 'allow_different_depots_in_route', label: 'Разрешить заезды на склады для дозагрузки' },
      { key: 'depots_only_at_run_beginning', label: 'Заезжать сначала только на склады' }
    ];

    const EXPORT_SCHEMAS = {
      orders: {
        sheetName: 'Orders',
        title: 'Orders — заказы',
        description: 'Основная таблица заказов для маршрутизации.',
        includeKeys: true,
        columns: [
          { key: 'id', label: 'Номер заказа', width: 14, description: 'Порядковый номер строки в экспортируемом файле.', getValue: (row) => row.id },
          { key: 'point.lat', label: 'Широта', width: 10, description: 'Широта точки доставки.', getValue: (row) => row['point.lat'] ?? '' },
          { key: 'point.lon', label: 'Долгота', width: 10, description: 'Долгота точки доставки.', getValue: (row) => row['point.lon'] ?? '' },
          { key: 'title', label: 'Наименование клиента', width: 26, description: 'Название магазина или клиента.', getValue: (row) => row.title || '' },
          { key: 'address', label: 'Адрес клиента', width: 32, description: 'Полный адрес точки доставки.', getValue: (row) => row.address || '' },
          { key: 'phone', label: 'Телефон', width: 18, description: 'Контактный телефон точки.', getValue: (row) => toStrPhone(row.phone) || '' },
          { key: 'time_window', label: 'Временное окно *', width: 16, description: 'Интервал обслуживания в формате HH:MM-HH:MM.', getValue: (row) => row.time_window || getDefaultTimeWindow() },
          { key: 'comments', label: 'Комментарии к заказу', width: 24, description: 'Дополнительные заметки для водителя.', getValue: (row) => row.comments || '' },
          { key: 'hard_window', label: 'Признак жесткого окна', width: 20, description: 'TRUE, если временное окно нельзя нарушать.', getValue: (row) => row.hard_window !== false },
          { key: 'shared_service_duration_s', label: 'Время обслуживания *', width: 18, description: 'Длительность обслуживания (сек).', getValue: (row) => pickValue(row, 'shared_service_duration_s', 'service_duration_s') || '' },
          { key: 'service_duration_s', label: '', width: 3, description: 'Дублирующий столбец для совместимости с шаблоном.', getValue: (row) => pickValue(row, 'service_duration_s', 'shared_service_duration_s') || '' },
          { key: 'shipment_size.weight_kg', label: 'Вес (брутто), кг', width: 14, description: 'Вес заказа, если требуется.', getValue: (row) => row.weight ?? '' },
          { key: 'shipment_size.units', label: 'Кол-во мест', width: 16, description: 'Количество мест, паллет или коробов.', getValue: (row) => row.units ?? '' },
          { key: 'shipment_size.volume_cbm', label: 'Объем, м³', width: 12, description: 'Объем поставки в кубометрах.', getValue: (row) => row.volume ?? '' },
          { key: 'type', label: 'Тип', width: 14, description: 'Тип точки (например garage).', getValue: (row) => row.type || '' },
          { key: 'depot_id', label: 'Наличие на складах', width: 18, description: 'ID склада, откуда отгружается заказ.', getValue: (row) => row.depot_id || '' }
        ]
      },
      vehicles: {
        sheetName: 'Vehicles',
        title: 'Vehicles — транспорт',
        description: 'Настройки автомобилей и водителей из вкладки Vehicles.',
        includeKeys: true,
        columns: [
          { key: 'id', label: 'ID (Идентификатор машины)', width: 18, description: 'Уникальный идентификатор машины.', getValue: (item) => item.id || '' },
          { key: 'ref', label: 'Имя курьера', width: 22, description: 'Отображаемое имя или позывной.', getValue: (item) => item.ref || '' },
          { key: 'capacity.weight_kg', label: 'Грузоподъемность, кг', width: 20, description: 'Максимальный вес груза для машины.', getValue: (item) => item['capacity.weight_kg'] ?? '' },
          { key: 'tags', label: 'Свойства', width: 24, description: 'Дополнительные свойства/теги машины.', getValue: (item) => item.tags || '' },
          { key: 'start_at', label: 'Начальная точка', width: 20, description: 'Старт маршрута (например depot:1).', getValue: (item) => item.start_at || '' },
          { key: 'finish_at', label: 'Конечная точка', width: 20, description: 'Завершение маршрута.', getValue: (item) => item.finish_at || '' },
          { key: 'visit_depot_at_start', label: 'Посетить склад в начале рейса', width: 22, description: 'TRUE, если перед началом рейса нужно заехать на склад.', getValue: (item) => boolFrom(item.visit_depot_at_start) },
          { key: 'return_to_depot', label: 'Надо вернуться на склад, Да/Нет', width: 24, description: 'TRUE, если по завершении маршрута нужно вернуться на склад.', getValue: (item) => boolFrom(item.return_to_depot) },
          { key: 'depot_id', label: 'Идентификаторы складов', width: 28, description: 'Склады, обслуживаемые машиной.', getValue: (item) => item.depot_id || '' },
          { key: 'shifts.0.id', label: 'Смена 1. ID', width: 18, description: 'Идентификатор первой смены.', getValue: (item) => item['shifts.0.id'] || '' },
          { key: 'shifts.0.time_window', label: 'Смена 1. Временное окно', width: 24, description: 'Рабочее окно первой смены.', getValue: (item) => item['shifts.0.time_window'] || '' },
          { key: 'allow_different_depots_in_route', label: 'Разрешить заезды на склады для дозагрузки', width: 30, description: 'TRUE, если допускаются промежуточные склады.', getValue: (item) => boolFrom(item.allow_different_depots_in_route) },
          { key: 'max_middle_depots', label: 'Максимальное количество промежуточных складов', width: 30, description: 'Лимит на дополнительные склады в маршруте.', getValue: (item) => item.max_middle_depots ?? '' },
          { key: 'depots_only_at_run_beginning', label: 'Заезжать сначала только на склады', width: 28, description: 'TRUE, если склады можно посещать только в начале.', getValue: (item) => boolFrom(item.depots_only_at_run_beginning) },
          { key: 'starting_depot_id', label: 'Стартовый склад', width: 22, description: 'ID склада, с которого начинается маршрут.', getValue: (item) => item.starting_depot_id || '' },
          { key: 'middle_depot_id', label: 'Промежуточный склад', width: 22, description: 'ID склада для промежуточной дозагрузки.', getValue: (item) => item.middle_depot_id || '' }
        ]
      },
      depots: {
        sheetName: 'Depot',
        title: 'Depot — склады',
        description: 'Справочник складов и координат.',
        includeKeys: true,
        columns: [
          { key: 'id', label: 'ID', width: 14, description: 'Идентификатор склада.', getValue: (item) => item.id || '' },
          { key: 'ref', label: 'Название', width: 20, description: 'Название склада.', getValue: (item) => item.ref || '' },
          { key: 'point.lat', label: 'Широта *', width: 10, description: 'Широта склада.', getValue: (item) => item['point.lat'] ?? '' },
          { key: 'point.lon', label: 'Долгота *', width: 10, description: 'Долгота склада.', getValue: (item) => item['point.lon'] ?? '' },
          { key: 'time_window', label: 'Время работы', width: 22, description: 'Рабочий интервал склада.', getValue: (item) => item.time_window || '' }
        ]
      }
    };

    const dataStore = {
      get vehicles() { return getActiveStores().vehicles; },
      set vehicles(v) { getActiveStores().vehicles = v; },
      get depots() { return getActiveStores().depots; },
      set depots(v) { getActiveStores().depots = v; },
      get startLocations() { return getActiveStores().startLocations; },
      set startLocations(v) { getActiveStores().startLocations = v; }
    };
    const state = {
      activeMode: DEFAULT_MODE,
      activeDay: 'monday',
      lastIndexByDay: {},
      query: '',
      suppressClickAfterDrag: false
    };
    Object.defineProperty(state, 'selected', {
      get() { return getActiveStores().selected; },
      enumerable: true
    });
    const dom = {
      root: document.getElementById('root'),
      schemaDocs: document.getElementById('schemaDocs'),
      vehTableWrap: document.getElementById('vehTableWrap'),
      depTableWrap: document.getElementById('depTableWrap'),
      startTableWrap: document.getElementById('startTableWrap'),
      startAddBtn: document.getElementById('startAdd'),
      startClearBtn: document.getElementById('startClear'),
      loadSheetBtn: document.getElementById('loadSheet'),
      sheetCacheStatus: document.getElementById('sheetCacheStatus'),
      vehAddBtn: document.getElementById('vehAdd'),
      vehClearBtn: document.getElementById('vehClear'),
      depAddBtn: document.getElementById('depAdd'),
      depClearBtn: document.getElementById('depClear'),
      vehCount: document.getElementById('vehCount'),
      vehSelectAllBtn: document.getElementById('vehSelectAll'),
      vehClearSelectionBtn: document.getElementById('vehClearSelection'),
      depCount: document.getElementById('depCount'),
      startCount: document.getElementById('startCount'),
      settingsExportBtn: document.getElementById('settingsExport'),
      settingsImportBtn: document.getElementById('settingsImportBtn'),
      settingsImportInput: document.getElementById('settingsImport'),
      tabs: Array.from(document.querySelectorAll('.tab')),
      jsonFile: document.getElementById('jsonFile'),
      loadJsonBtn: document.getElementById('loadJson'),
      errors: document.getElementById('errors'),
      errorMessage: document.getElementById('errorMessage'),
      errorClose: document.getElementById('errorClose'),
      notify: document.getElementById('notify'),
      notifyMessage: document.getElementById('notifyMessage'),
      notifyIcon: document.getElementById('notifyIcon'),
      notifyClose: document.getElementById('notifyClose'),
      storesSection: document.getElementById('storesSection'),
      dataSection: document.getElementById('dataSection'),
      modeButtons: Array.from(document.querySelectorAll('.mode-btn')),
      dataSourceLabel: document.getElementById('dataSourceLabel'),
      sheetOnboarding: document.getElementById('sheetOnboarding'),
      sheetOnboardingSpotlight: document.getElementById('sheetOnboardingSpotlight'),
      sheetOnboardingHint: document.getElementById('sheetOnboardingHint'),
      sheetOnboardingBackdrop: document.getElementById('sheetOnboardingBackdrop'),
      sheetOnboardingDismiss: document.getElementById('sheetOnboardingDismiss'),
      extraOrderModal: document.getElementById('extraOrderModal'),
      extraOrderModalBackdrop: document.getElementById('extraOrderModalBackdrop'),
      extraOrderModalClose: document.getElementById('extraOrderModalClose'),
      extraOrderModalDay: document.getElementById('extraOrderModalDay'),
      extraOrderFormBlocks: document.getElementById('extraOrderFormBlocks'),
      extraOrderAddBlockBtn: document.getElementById('extraOrderAddBlock'),
      extraOrderCancelBtn: document.getElementById('extraOrderCancel'),
      extraOrderSubmitBtn: document.getElementById('extraOrderSubmit')
    };

    const SHEET_ONBOARDING_KEY = 'sheet_onboarding_done';

    function shouldShowSheetOnboarding() {
      try {
        if (localStorage.getItem(SHEET_ONBOARDING_KEY) === '1') return false;
      } catch (_) {}
      if (isAllMode()) {
        return SOURCE_MODE_IDS.some((id) => modeStartsEmptyUntilSync(id) && !hasSheetSynced(id));
      }
      if (!shouldLoadPersistedForMode(getActiveMode())) return true;
      return !hasSheetSynced(getActiveMode());
    }

    function dismissSheetOnboarding(persist) {
      document.body.classList.remove('sheet-onboarding-active');
      if (dom.sheetOnboarding) {
        dom.sheetOnboarding.hidden = true;
        dom.sheetOnboarding.setAttribute('aria-hidden', 'true');
      }
      if (persist !== false) {
        try {
          localStorage.setItem(SHEET_ONBOARDING_KEY, '1');
        } catch (_) {}
      }
    }

    function positionSheetOnboarding() {
      if (!dom.loadSheetBtn || !dom.sheetOnboardingSpotlight || !dom.sheetOnboardingHint) return;
      const rect = dom.loadSheetBtn.getBoundingClientRect();
      const pad = 10;
      dom.sheetOnboardingSpotlight.style.top = `${Math.max(8, rect.top - pad)}px`;
      dom.sheetOnboardingSpotlight.style.left = `${Math.max(8, rect.left - pad)}px`;
      dom.sheetOnboardingSpotlight.style.width = `${rect.width + pad * 2}px`;
      dom.sheetOnboardingSpotlight.style.height = `${rect.height + pad * 2}px`;

      const hintGap = 14;
      let hintTop = rect.bottom + hintGap;
      const hintEl = dom.sheetOnboardingHint;
      const hintHeight = hintEl.offsetHeight || 120;
      if (hintTop + hintHeight > window.innerHeight - 16) {
        hintTop = Math.max(16, rect.top - hintGap - hintHeight);
      }
      let hintLeft = rect.left;
      const hintWidth = hintEl.offsetWidth || 300;
      if (hintLeft + hintWidth > window.innerWidth - 16) {
        hintLeft = window.innerWidth - hintWidth - 16;
      }
      hintEl.style.top = `${hintTop}px`;
      hintEl.style.left = `${Math.max(16, hintLeft)}px`;
    }

    function showSheetOnboarding() {
      if (!dom.sheetOnboarding || !dom.loadSheetBtn) return;
      if (!shouldShowSheetOnboarding()) return;
      document.body.classList.add('sheet-onboarding-active');
      dom.sheetOnboarding.hidden = false;
      dom.sheetOnboarding.setAttribute('aria-hidden', 'false');
      positionSheetOnboarding();
    }

    function initSheetOnboarding() {
      if (!dom.sheetOnboarding) return;
      if (dom.sheetOnboardingBackdrop) {
        dom.sheetOnboardingBackdrop.addEventListener('click', () => dismissSheetOnboarding(true));
      }
      if (dom.sheetOnboardingDismiss) {
        dom.sheetOnboardingDismiss.addEventListener('click', () => dismissSheetOnboarding(true));
      }
      window.addEventListener('resize', () => {
        if (!dom.sheetOnboarding.hidden) positionSheetOnboarding();
      });
      window.addEventListener('scroll', () => {
        if (!dom.sheetOnboarding.hidden) positionSheetOnboarding();
      }, true);
      document.addEventListener('keydown', (ev) => {
        if (ev.key === 'Escape' && dom.sheetOnboarding && !dom.sheetOnboarding.hidden) {
          dismissSheetOnboarding(true);
        }
      });
      requestAnimationFrame(() => {
        requestAnimationFrame(() => showSheetOnboarding());
      });
    }

    // ===== Утилиты =====
    function saveLocal() {
      const mode = getActiveMode();
      if (isAllMode(mode)) {
        try {
          localStorage.setItem(
            storageKey('extraOrders', ALL_MODE),
            JSON.stringify(getStoreForMode(ALL_MODE).extraOrders)
          );
        } catch (_) {}
        return;
      }
      if (!shouldLoadPersistedForMode(mode)) return;
      try {
        const store = getStoreForMode(mode);
        const payload = {
          vehicles: store.vehicles,
          depots: store.depots,
          startLocations: store.startLocations,
          extraOrders: store.extraOrders || createEmptyExtraOrders()
        };
        localStorage.setItem(storageKey('vehicles', mode), JSON.stringify(payload.vehicles));
        localStorage.setItem(storageKey('depots', mode), JSON.stringify(payload.depots));
        localStorage.setItem(storageKey('startLocations', mode), JSON.stringify(payload.startLocations));
        localStorage.setItem(storageKey('extraOrders', mode), JSON.stringify(payload.extraOrders));
        localStorage.setItem(
          storageKey('selectedVehicleIds', mode),
          JSON.stringify(Array.from(store.selectedVehicleIds))
        );
        if (hasValidSheetCache(mode)) {
          saveSheetImportCache(mode, { preserveImportedAt: true });
        }
        try {
          window.name = JSON.stringify({
            mode,
            ...payload,
            selectedVehicleIds: Array.from(store.selectedVehicleIds)
          });
        } catch (_) {}
        const status = document.getElementById('saveStatus');
        if (status) {
          status.textContent = 'Сохранено';
          setTimeout(() => {
            if (status.textContent === 'Сохранено') status.textContent = '';
          }, 2000);
        }
      } catch (err) {
        const status = document.getElementById('saveStatus');
        if (status) status.textContent = '⚠︎ Не удалось сохранить в этом браузере';
      }
    }

    function loadLocal() {
      const mode = getActiveMode();
      if (isAllMode(mode)) {
        loadExtraOrdersOnly(ALL_MODE);
        loadAllModeSelectionCache();
        updateSheetCacheUi();
        return;
      }

      loadExtraOrdersOnly(mode);

      if (loadSheetImportCache(mode)) {
        const status = document.getElementById('saveStatus');
        if (status) {
          status.textContent = 'Данные восстановлены';
          setTimeout(() => { status.textContent = ''; }, 2000);
        }
        updateSheetCacheUi();
        return;
      }

      if (!shouldLoadPersistedForMode(mode)) {
        return;
      }

      const loaded = loadLegacyReferenceData(mode);
      loadExtraOrdersOnly(mode);
      if (!loaded) {
        try {
          const stash = JSON.parse(window.name || '{}');
          if (stash.mode === getActiveMode() || !stash.mode) {
            if (Array.isArray(stash.vehicles)) {
              dataStore.vehicles = stash.vehicles.map((item) => normalizeVehicleRecord(item)).filter(Boolean);
            }
            if (Array.isArray(stash.depots)) {
              dataStore.depots = stash.depots.map((item) => normalizeDepotRecord(item)).filter(Boolean);
            }
            if (Array.isArray(stash.startLocations)) {
              dataStore.startLocations = stash.startLocations.map((item) => normalizeStartRecord(item)).filter(Boolean);
            }
          }
        } catch (_) {}
      }
      const status = document.getElementById('saveStatus');
      if (status && loaded) {
        status.textContent = 'Настройки загружены';
        setTimeout(() => { status.textContent = ''; }, 2000);
      }
      updateSheetCacheUi();
    }

    function boolFrom(v) {
      if (typeof v === 'boolean') return v;
      if (v == null) return false;
      const norm = String(v).trim().toLowerCase();
      return ['true', '1', 'yes', 'да', 'y', 'on'].includes(norm);
    }

    function showError(msg) {
      if (!dom.errors || !dom.errorMessage) return;
      dom.errorMessage.textContent = msg;
      dom.errors.setAttribute('aria-hidden', 'false');
    }

    function clearError() {
      if (!dom.errors || !dom.errorMessage) return;
      dom.errors.setAttribute('aria-hidden', 'true');
      dom.errorMessage.textContent = '';
    }


    const LIST_LAZY_BATCH = 40;
    let listRenderToken = 0;
    let listLazyObserver = null;
    let notifyHideTimer = null;

    const MODE_LOAD_LABELS = { horeca: 'HoReCa', gallery: 'Галереи' };

    function ensureActionButton(btn) {
      if (!btn) return null;
      if (!btn.classList.contains('btn-action')) {
        btn.classList.add('btn-action');
        const label = (btn.textContent || '').trim() || 'Действие';
        btn.textContent = '';
        btn.innerHTML = `
          <span class="btn-action__spinner" aria-hidden="true"></span>
          <span class="btn-action__icon" aria-hidden="true"></span>
          <span class="btn-action__label">${escapeHtml(label)}</span>
        `;
        btn.dataset.defaultLabel = label;
      }
      return btn;
    }

    function setActionButtonState(btn, state, options = {}) {
      if (!btn) return;
      ensureActionButton(btn);
      btn.classList.remove('btn-action--loading', 'btn-action--success');
      const labelEl = btn.querySelector('.btn-action__label');
      const defaultLabel = btn.dataset.defaultLabel || (labelEl ? labelEl.textContent : '') || '';
      if (!btn.dataset.defaultLabel && defaultLabel) btn.dataset.defaultLabel = defaultLabel;
      if (state === 'loading') {
        btn.disabled = true;
        btn.classList.add('btn-action--loading');
        btn.setAttribute('aria-busy', 'true');
        if (labelEl) labelEl.textContent = options.loadingText || 'Загрузка…';
        return;
      }
      btn.removeAttribute('aria-busy');
      if (state === 'success') {
        btn.disabled = false;
        btn.classList.add('btn-action--success');
        if (labelEl) labelEl.textContent = options.successText || 'Готово';
        return;
      }
      btn.disabled = options.disabled === true;
      if (labelEl) labelEl.textContent = options.label || defaultLabel;
    }

    function setSyncingUi(syncing) {
      dom.dataSection?.classList.toggle('is-syncing', syncing);
    }

    function hideNotify() {
      if (notifyHideTimer) {
        clearTimeout(notifyHideTimer);
        notifyHideTimer = null;
      }
      if (!dom.notify) return;
      dom.notify.setAttribute('aria-hidden', 'true');
    }

    function showNotify(message, type = 'success', duration = 4500) {
      if (!dom.notify || !dom.notifyMessage) return;
      hideNotify();
      dom.notifyMessage.textContent = message;
      dom.notify.classList.remove('toast--success', 'toast--info');
      dom.notify.classList.add(type === 'info' ? 'toast--info' : 'toast--success');
      if (dom.notifyIcon) dom.notifyIcon.textContent = type === 'info' ? 'i' : '✓';
      dom.notify.setAttribute('aria-hidden', 'false');
      if (duration > 0) {
        notifyHideTimer = setTimeout(hideNotify, duration);
      }
    }

    function celebrateExport(btn) {
      const target = btn || document.getElementById('exportXlsx');
      if (!target) return;
      target.classList.add('export-celebrate');
      setTimeout(() => target.classList.remove('export-celebrate'), 900);
    }

    function disconnectListLazyObserver() {
      if (listLazyObserver) {
        listLazyObserver.disconnect();
        listLazyObserver = null;
      }
    }

    function createStoreListItem(row, idx, day) {
      const li = document.createElement('li');
      li.className = row.__isExtra ? 'store store--extra store--lazy-enter' : 'store store--lazy-enter';
      li.dataset.index = String(idx);
      li.dataset.uid = row.uid || '';
      if (row.__isExtra) {
        li.dataset.isExtra = 'true';
        li.dataset.extraId = row.id || '';
      }
      if (row.__isStart) {
        li.dataset.isStart = 'true';
        const startId = normalizeStartId(row.id);
        if (startId) li.dataset.startId = startId;
      }
      if (row.__sourceMode) li.dataset.sourceMode = row.__sourceMode;

      const hasCoords = row.lat != null && row.lng != null;
      const addrHtml = row.address
        ? hasCoords
          ? `<a class="muted addr-link" href="http://maps.yandex.ru/?text=${encodeURIComponent(row.lat + ',' + row.lng)}" target="_blank" rel="noopener">${highlight(escapeHtml(row.address), state.query)}</a><span style="opacity:.6">&nbsp;↗</span>`
          : `<span class="muted">${highlight(escapeHtml(row.address), state.query)}</span>`
        : '';
      const mins = Number.isFinite(row.delivery_seconds) ? Math.round(row.delivery_seconds / 60) : null;
      const timeWindow = row.time_window ? `<span class="pill">⏰ ${escapeHtml(row.time_window)}</span>` : '';
      const depot = row.depot_name || row.depot_id ? `<span class="pill" title="Склад">🏷️ ${escapeHtml(row.depot_name ? `${row.depot_name}` : `ID ${row.depot_id}`)}</span>` : '';
      const serviceTime = mins !== null ? `<span class="pill" title="Время на доставку (сек): ${row.delivery_seconds}">⏱ ${mins} мин</span>` : '';
      const typePill = row.__isStart && row.type ? `<span class="pill">🏁 ${escapeHtml(row.type)}</span>` : '';
      const extraPill = row.__isExtra ? '<span class="pill" title="Разовая точка, не из Google Sheets">✨ Разовая</span>' : '';
      const sourcePill = row.__sourceLabel
        ? `<span class="pill pill--source" title="Источник данных">${escapeHtml(row.__sourceLabel)}</span>`
        : '';
      const phoneHtml = row.phone ? `<span class="muted">☎ ${highlight(escapeHtml(toStrPhone(row.phone)), state.query)}</span>` : '';
      const pillsHtml = [sourcePill, timeWindow, serviceTime, depot, typePill, extraPill].filter(Boolean).join('');
      const extraDeleteBtn = row.__isExtra
        ? `<div class="store-extra-actions"><button type="button" class="mini-btn" data-act="del-extra" title="Удалить разовую точку">Удалить</button></div>`
        : '';
      const detailParts = [];
      if (addrHtml) detailParts.push(addrHtml);
      if (phoneHtml) detailParts.push(phoneHtml);
      const detailsHtml = detailParts.length ? detailParts.join('<span class="sep">•</span>') : '';
      const metaHtml = (pillsHtml || detailsHtml) ? `
        <div class="store-meta">
          ${pillsHtml ? `<div class="store-meta-pills">${pillsHtml}</div>` : ''}
          ${detailsHtml ? `<div class="store-meta-details">${detailsHtml}</div>` : ''}
        </div>` : '';

      li.innerHTML = `
        <div class="row-item">
          <input type="checkbox" data-uid="${row.uid}" ${state.selected.has(row.uid) ? 'checked' : ''} />
          <div class="store-line">
            <span class="title">${highlight(escapeHtml(row.title || row.store || 'Без названия'), state.query)}</span>
            ${metaHtml}
          </div>
          ${extraDeleteBtn}
        </div>
      `;
      if (row.__isExtra) {
        const delBtn = li.querySelector('[data-act="del-extra"]');
        if (delBtn) {
          delBtn.addEventListener('click', (ev) => {
            ev.stopPropagation();
            if (confirm('Удалить эту разовую точку?')) {
              removeExtraOrder(day, row.id, row.__sourceMode);
            }
          });
        }
      }
      return li;
    }

    function mountStoreLazySentinel(listEl, items, day, nextIndex, token) {
      const sentinel = document.createElement('li');
      sentinel.className = 'stores-lazy-sentinel';
      sentinel.innerHTML = `
        <div class="lazy-sentinel-inner">
          <div class="store-skeleton"></div>
          <div class="store-skeleton"></div>
        </div>
      `;
      listEl.appendChild(sentinel);
      disconnectListLazyObserver();
      listLazyObserver = new IntersectionObserver((entries) => {
        entries.forEach((entry) => {
          if (!entry.isIntersecting || token !== listRenderToken) return;
          disconnectListLazyObserver();
          sentinel.remove();
          appendStoreLazyBatch(listEl, items, day, nextIndex, token);
        });
      }, { root: null, rootMargin: '240px 0px', threshold: 0 });
      listLazyObserver.observe(sentinel);
    }

    function appendStoreLazyBatch(listEl, items, day, startIndex, token) {
      if (token !== listRenderToken) return startIndex;
      const end = Math.min(startIndex + LIST_LAZY_BATCH, items.length);
      const frag = document.createDocumentFragment();
      for (let i = startIndex; i < end; i += 1) {
        frag.appendChild(createStoreListItem(items[i], i, day));
      }
      const sentinel = listEl.querySelector('.stores-lazy-sentinel');
      if (sentinel) sentinel.remove();
      listEl.appendChild(frag);
      if (end < items.length) mountStoreLazySentinel(listEl, items, day, end, token);
      return end;
    }

    function mountStoreListLazy(listEl, items, day) {
      disconnectListLazyObserver();
      const token = ++listRenderToken;
      listEl.innerHTML = '';
      if (!items.length) return;
      const end = appendStoreLazyBatch(listEl, items, day, 0, token);
      if (end < items.length) mountStoreLazySentinel(listEl, items, day, end, token);
    }

    function escapeHtml(value) {
      return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
    }

    function highlight(text, query) {
      if (!query) return String(text);
      const esc = String(query).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      try {
        return String(text).replace(new RegExp(esc, 'ig'), (m) => `<mark>${m}</mark>`);
      } catch (err) {
        console.warn('Highlight regex failed for query', query, err);
        return String(text);
      }
    }

    function updateSelectedHighlight() {
      document.querySelectorAll('.store').forEach((item) => {
        const checkbox = item.querySelector('input[type="checkbox"]');
        if (checkbox && checkbox.checked) item.classList.add('selected');
        else item.classList.remove('selected');
      });
    }

    function updateSelectedCount() {
      const counter = document.getElementById('selectedCount');
      if (!counter) return;

      if (isAllMode()) {
        let visibleCount = 0;
        state.selected.forEach((uid) => {
          if (typeof uid !== 'string') return;
          if (isStartSelectionUid(uid)) return;
          visibleCount += 1;
        });
        counter.textContent = `Выбрано: ${visibleCount}`;
        updateExportButtonState();
        return;
      }

      // Считаем выбранные точки, исключая три специальные стартовые:
      // 1111 (ИМ Восток), 1110 (ИМ Центр), 1114 (Дом).
      const rules = getModeRules();
      const excludedStartIds = new Set([
        normalizeStartId(rules.priorityStartLocationId),
        normalizeStartId(rules.defaultStartLocationId),
        normalizeStartId(rules.defaultExtraStartId)
      ].filter(Boolean));

      let visibleCount = 0;
      const startRows = dataStore.startLocations || [];

      state.selected.forEach((uid) => {
        if (typeof uid !== 'string') return;
        if (uid.startsWith('start::')) {
          const idx = Number(uid.split('::')[1]);
          if (!Number.isFinite(idx)) return;
          const item = startRows[idx];
          const idValue = normalizeStartId(item && item.id);
          if (!idValue || excludedStartIds.has(idValue)) return;
          visibleCount += 1;
        } else {
          visibleCount += 1;
        }
      });

      counter.textContent = `Выбрано: ${visibleCount}`;
      updateExportButtonState();
    }

    /** Точки для экспорта: заказы и разовые; стартовые не считаются. */
    function countExportableSelections() {
      let count = 0;
      state.selected.forEach((uid) => {
        if (typeof uid !== 'string') return;
        if (isStartSelectionUid(uid)) return;
        count += 1;
      });
      return count;
    }

    function hasExportableSelection() {
      return countExportableSelections() > 0;
    }

    function updateExportButtonState() {
      const btn = document.getElementById('exportXlsx');
      if (!btn) return;
      if (btn.classList.contains('btn-action--loading') || btn.classList.contains('btn-action--success')) return;
      const ready = hasExportableSelection();
      btn.disabled = !ready;
      btn.classList.toggle('export-btn--ready', ready);
      btn.title = ready ? '' : 'Выберите хотя бы одну точку доставки (не стартовую)';
    }

    function toStrPhone(phone) {
      if (phone == null) return '';
      try { return String(phone).trim(); } catch (_) { return ''; }
    }

    function extractRuPhoneDigits(value) {
      let digits = String(value ?? '').replace(/\D/g, '');
      if (!digits) return '';
      if (digits.length >= 11 && digits.startsWith('8')) digits = `7${digits.slice(1)}`;
      if (digits.startsWith('7')) digits = digits.slice(1);
      return digits.slice(0, 10);
    }

    function formatRuPhoneMaskFromDigits(digits) {
      const d = String(digits ?? '').replace(/\D/g, '').slice(0, 10);
      if (!d.length) return '';
      let out = `+7 (${d.slice(0, 3)}`;
      if (d.length <= 3) return out;
      out += `) ${d.slice(3, 6)}`;
      if (d.length <= 6) return out;
      out += `-${d.slice(6, 8)}`;
      if (d.length <= 8) return out;
      out += `-${d.slice(8, 10)}`;
      return out;
    }

    function formatRuPhoneInputValue(value) {
      return formatRuPhoneMaskFromDigits(extractRuPhoneDigits(value));
    }

    function handleExtraOrderPhoneInput(input) {
      if (!input) return;
      const prevLen = input.value.length;
      const sel = input.selectionStart ?? prevLen;
      const formatted = formatRuPhoneMaskFromDigits(extractRuPhoneDigits(input.value));
      input.value = formatted;
      if (!formatted) return;
      let newSel = sel + (formatted.length - prevLen);
      if (newSel < 4) newSel = formatted.length;
      if (newSel > formatted.length) newSel = formatted.length;
      input.setSelectionRange(newSel, newSel);
    }

    function handleExtraOrderPhoneFocus(input) {
      if (!input || input.value.trim()) return;
      input.value = '+7 (';
      input.setSelectionRange(input.value.length, input.value.length);
    }

    function handleExtraOrderPhoneBlur(input) {
      if (!input) return;
      const digits = extractRuPhoneDigits(input.value);
      input.value = digits.length ? formatRuPhoneMaskFromDigits(digits) : '';
    }

    function formatDateForFile(date) {
      const dd = String(date.getDate()).padStart(2, '0');
      const mm = String(date.getMonth() + 1).padStart(2, '0');
      const yyyy = date.getFullYear();
      return `${dd}.${mm}.${yyyy}`;
    }

    function getDeliveryDateForDay(dayKey) {
      const today = new Date();
      const todayJs = today.getDay(); // 0 (вс) – 6 (сб)
      const targetJs = CANON_TO_JS_DAY[dayKey] ?? todayJs;
      let diff = (targetJs - todayJs + 7) % 7;
      const result = new Date(today);
      result.setHours(0, 0, 0, 0);
      if (diff !== 0) {
        result.setDate(result.getDate() + diff);
      }
      return result;
    }

    function toNumOrNull(value) {
      if (value == null || value === '') return null;
      let candidate = value;
      if (typeof candidate === 'string') {
        candidate = candidate.replace(',', '.').trim();
      }
      const num = Number(candidate);
      return Number.isFinite(num) ? num : null;
    }

    /** Парсит «55.706284, 37.781865» или «55.7 37.7» из буфера Яндекс.Карт. */
    function parseLatLngPair(text) {
      if (text == null) return null;
      const cleaned = String(text).trim();
      if (!cleaned) return null;

      const tryPair = (a, b) => {
        const lat = toNumOrNull(a);
        const lng = toNumOrNull(b);
        if (lat == null || lng == null) return null;
        if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
        return { lat, lng };
      };

      const commaParts = cleaned.split(/[,;]/).map((p) => p.trim()).filter(Boolean);
      if (commaParts.length >= 2) {
        const pair = tryPair(commaParts[0], commaParts[1]);
        if (pair) return pair;
      }

      const spaceParts = cleaned.split(/\s+/).filter(Boolean);
      if (spaceParts.length >= 2) {
        const pair = tryPair(spaceParts[0], spaceParts[1]);
        if (pair) return pair;
      }

      return null;
    }

    function applyParsedCoordsToBlock(block, parsed) {
      if (!block || !parsed) return false;
      const latInput = block.querySelector('input[data-field="lat"]');
      const lngInput = block.querySelector('input[data-field="lng"]');
      const pasteInput = block.querySelector('input[data-field="coords_paste"]');
      if (latInput) latInput.value = String(parsed.lat);
      if (lngInput) lngInput.value = String(parsed.lng);
      if (pasteInput) pasteInput.value = `${parsed.lat}, ${parsed.lng}`;
      return true;
    }

    function tryApplyCoordsTextToBlock(block, text) {
      const parsed = parseLatLngPair(text);
      if (!parsed) return false;
      return applyParsedCoordsToBlock(block, parsed);
    }

    function isRowEmpty(row) {
      if (!Array.isArray(row)) return true;
      return row.every((cell) => cell == null || String(cell).trim() === '');
    }

    function normalizeOrderRecord(raw) {
      if (!raw || typeof raw !== 'object') return null;
      const lat = toNumOrNull(raw['point.lat'] ?? raw.lat ?? raw.latitude);
      const lon = toNumOrNull(raw['point.lon'] ?? raw.lon ?? raw.lng ?? raw.longitude);
      let serviceSeconds = null;
      if (raw.delivery_seconds != null && raw.delivery_seconds !== '') {
        serviceSeconds = toNumOrNull(raw.delivery_seconds);
      } else {
        const rawDuration = toNumOrNull(
          raw.service_duration_s ??
          raw.service_seconds ??
          raw.shared_service_duration_s ??
          raw.shared_service_duration_minutes
        );
        serviceSeconds = rawDuration;
        if (serviceSeconds != null && serviceSeconds < 1000) {
          serviceSeconds = serviceSeconds * 60;
        }
      }
      const record = {
        title: raw.title || raw.name || '',
        address: raw.address || raw.Адрес || raw.Адреса || '',
        lat,
        lng: lon,
        phone: toStrPhone(raw.phone),
        depot_name: raw.depot_name || raw.depot || '',
        depot_id: raw.depot_id != null ? String(raw.depot_id) : '',
        time_window: raw.time_window || getDefaultTimeWindow(),
        comments: raw.comments || '',
        delivery_seconds: serviceSeconds
      };
      return record;
    }

    function generateExtraOrderId() {
      if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return crypto.randomUUID();
      }
      return `extra-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
    }

    function normalizeExtraOrderRecord(raw) {
      if (!raw || typeof raw !== 'object') return null;
      const id = raw.id != null ? String(raw.id).trim() : '';
      if (!id) return null;
      const base = normalizeOrderRecord(raw);
      if (!base) return null;
      return { ...base, id };
    }

    function normalizeExtraOrdersStructure(raw) {
      const base = createEmptyExtraOrders();
      if (!raw || typeof raw !== 'object') return base;
      DAY_KEYS.forEach((day) => {
        if (!Array.isArray(raw[day])) return;
        base[day] = raw[day].map((item) => normalizeExtraOrderRecord(item)).filter(Boolean);
      });
      return base;
    }

    function getExtraOrders() {
      const store = getActiveStores();
      if (!store.extraOrders || typeof store.extraOrders !== 'object') {
        store.extraOrders = createEmptyExtraOrders();
      }
      return store.extraOrders;
    }

    function extraOrderUid(day, id, sourceMode) {
      if (sourceMode === ALL_MODE) return `${ALL_MODE}::extra::${day}::${id}`;
      if (sourceMode) return `${sourceMode}::extra::${day}::${id}`;
      return `extra::${day}::${id}`;
    }

    function scheduleOrderUid(day, record, index, sourceMode) {
      const tail = `${record.title || record.store || ''}::${record.address || ''}::${toStrPhone(record.phone)}`;
      if (sourceMode) return `${sourceMode}::order::${day}::${index}::${tail}`;
      return `${day}::${index}::${tail}`;
    }

    function startSelectionUid(sourceMode, index) {
      if (sourceMode) return `${sourceMode}::start::${index}`;
      return `start::${index}`;
    }

    function mapExtraOrder(record, day, sourceMode) {
      return {
        ...record,
        uid: extraOrderUid(day, record.id, sourceMode),
        __isExtra: true,
        __sourceMode: sourceMode || getActiveMode(),
        __sourceLabel: sourceMode ? getModeConfig(sourceMode).label : ''
      };
    }

    function buildDayListItems(day, options = {}) {
      const includeStarts = options.includeStarts !== false
        && shouldLoadPersistedForMode(getActiveMode());

      if (isAllMode()) {
        const items = [];
        if (includeStarts) {
          getMergedUniqueStartLocations().forEach((entry, index) => {
            const mapped = mapStartLocation(entry.record, index);
            const sources = formatSourceLabels(entry.sourceLabels);
            items.push({
              ...mapped,
              uid: entry.uid,
              __isStart: true,
              __sourceMode: entry.sourceModes[0],
              __sourceLabel: sources
            });
          });
        }
        SOURCE_MODE_IDS.forEach((sourceMode) => {
          if (!shouldLoadPersistedForMode(sourceMode)) return;
          const store = getStoreForMode(sourceMode);
          const label = getModeConfig(sourceMode).label;
          (store.scheduleData[day] || []).forEach((record, index) => {
            items.push({
              ...record,
              uid: scheduleOrderUid(day, record, index, sourceMode),
              __sourceMode: sourceMode,
              __sourceLabel: label
            });
          });
          const extras = store.extraOrders || createEmptyExtraOrders();
          (extras[day] || []).forEach((record) => {
            items.push(mapExtraOrder(record, day, sourceMode));
          });
        });
        (getExtraOrders()[day] || []).forEach((record) => {
          items.push(mapExtraOrder(record, day, ALL_MODE));
        });
        return items;
      }

      const startEntries = includeStarts
        ? (dataStore.startLocations || []).map((record, index) => mapStartLocation(record, index))
        : [];
      const listRaw = (getScheduleData()[day] || []).map((record, index) => ({
        ...record,
        uid: scheduleOrderUid(day, record, index)
      }));
      const extraRaw = (getExtraOrders()[day] || []).map((record) => mapExtraOrder(record, day));
      return [...startEntries, ...listRaw, ...extraRaw];
    }

    function removeExtraOrder(day, id, sourceMode) {
      const mode = sourceMode || getActiveMode();
      const store = getStoreForMode(mode);
      if (!store.extraOrders) store.extraOrders = createEmptyExtraOrders();
      const orders = store.extraOrders[day];
      if (!Array.isArray(orders)) return;
      const idx = orders.findIndex((item) => item && item.id === id);
      if (idx === -1) return;
      orders.splice(idx, 1);
      state.selected.delete(extraOrderUid(day, id, mode));
      saveLocal();
      persistActiveSheetCacheSelection();
      render();
    }

    function normalizeVehicleRecord(raw) {
      if (!raw || typeof raw !== 'object') return null;
      const record = {
        id: raw.id != null ? String(raw.id) : '',
        ref: raw.ref != null ? String(raw.ref) : '',
        'capacity.weight_kg': toNumOrNull(raw['capacity.weight_kg'] ?? raw.capacity_weight ?? raw.capacityWeight),
        tags: raw.tags != null ? String(raw.tags) : '',
        start_at: raw.start_at != null ? String(raw.start_at) : '',
        finish_at: raw.finish_at != null ? String(raw.finish_at) : '',
        visit_depot_at_start: boolFrom(raw.visit_depot_at_start),
        return_to_depot: boolFrom(raw.return_to_depot),
        depot_id: raw.depot_id != null ? String(raw.depot_id) : '',
        'shifts.0.id': raw['shifts.0.id'] != null ? String(raw['shifts.0.id']) : '',
        'shifts.0.time_window': raw['shifts.0.time_window'] != null ? String(raw['shifts.0.time_window']) : '',
        allow_different_depots_in_route: boolFrom(raw.allow_different_depots_in_route),
        max_middle_depots: toNumOrNull(raw.max_middle_depots),
        depots_only_at_run_beginning: boolFrom(raw.depots_only_at_run_beginning),
        starting_depot_id: raw.starting_depot_id != null ? String(raw.starting_depot_id) : '',
        middle_depot_id: raw.middle_depot_id != null ? String(raw.middle_depot_id) : ''
      };
      return record;
    }

    function normalizeDepotRecord(raw) {
      if (!raw || typeof raw !== 'object') return null;
      const lat = toNumOrNull(raw['point.lat'] ?? raw.lat ?? raw.latitude);
      const lon = toNumOrNull(raw['point.lon'] ?? raw.lon ?? raw.lng ?? raw.longitude);
      const record = {
        id: raw.id != null ? String(raw.id) : '',
        ref: raw.ref != null ? String(raw.ref) : '',
        'point.lat': lat == null ? '' : lat,
        'point.lon': lon == null ? '' : lon,
        time_window: raw.time_window != null ? String(raw.time_window) : ''
      };
      return record;
    }

    function normalizeStartRecord(raw) {
      if (!raw || typeof raw !== 'object') return null;
      const lat = toNumOrNull(raw['point.lat'] ?? raw.lat ?? raw.latitude);
      const lon = toNumOrNull(raw['point.lon'] ?? raw.lon ?? raw.lng ?? raw.longitude);
      return {
        id: raw.id != null ? String(raw.id) : '',
        ref: raw.ref != null ? String(raw.ref) : '',
        'point.lat': lat == null ? '' : lat,
        'point.lon': lon == null ? '' : lon,
        time_window: raw.time_window != null ? String(raw.time_window) : '',
        type: raw.type != null ? String(raw.type) : '',
        address: raw.address != null ? String(raw.address) : '',
        comments: raw.comments != null ? String(raw.comments) : '',
        phone: raw.phone != null ? String(raw.phone) : ''
      };
    }

    function parseStructuredSheet(sheetData, normalizer) {
      if (Array.isArray(sheetData) && sheetData.length && typeof sheetData[0] === 'object' && !Array.isArray(sheetData[0])) {
        return sheetData.map((item) => normalizer(item)).filter(Boolean);
      }
      if (!Array.isArray(sheetData) || sheetData.length < 2) return [];
      const keysRow = sheetData[1] || [];
      const entries = [];
      for (let i = 2; i < sheetData.length; i += 1) {
        const row = sheetData[i];
        if (!row || isRowEmpty(row)) continue;
        const raw = {};
        keysRow.forEach((key, idx) => {
          if (!key) return;
          raw[key] = row[idx];
        });
        const normalized = normalizer(raw);
        if (normalized) entries.push(normalized);
      }
      return entries;
    }

    function startLocationUid(record, index) {
      return `start::${index}`;
    }

    function mapStartLocation(record, index) {
      const lat = toNumOrNull(record['point.lat']);
      const lon = toNumOrNull(record['point.lon']);
      return {
        ...record,
        lat,
        lng: lon,
        title: record.ref || record.id || `Старт #${index + 1}`,
        address: record.address || '',
        time_window: record.time_window || getDefaultTimeWindow(),
        comments: record.comments || '',
        phone: record.phone || '',
        depot_id: '',
        type: record.type || '',
        uid: startLocationUid(record, index),
        __isStart: true
      };
    }

    function normalizeStartId(value) {
      if (value == null) return '';
      try {
        return String(value).trim();
      } catch (_) {
        return '';
      }
    }

    function ensureDefaultExtraStartSelected() {
      const defaultId = normalizeStartId(getModeRules().defaultExtraStartId);
      if (!defaultId) return;
      const startRows = dataStore.startLocations || [];
      for (let index = 0; index < startRows.length; index += 1) {
        const item = startRows[index];
        if (normalizeStartId(item && item.id) !== defaultId) continue;
        const uid = startLocationUid(item, index);
        state.selected.add(uid);
        return;
      }
    }

    function isDefaultExtraSelected() {
      const defaultId = normalizeStartId(getModeRules().defaultExtraStartId);
      if (!defaultId) return false;
      const startRows = dataStore.startLocations || [];
      return Array.from(state.selected).some((uid) => {
        if (!uid || typeof uid !== 'string' || !uid.startsWith('start::')) return false;
        const idx = Number(uid.split('::')[1]);
        if (!Number.isFinite(idx)) return false;
        const item = startRows[idx];
        return normalizeStartId(item && item.id) === defaultId;
      });
    }

    function hasOrdersWithDepotId(targetDepotId, dayKey, sourceMode) {
      const normalizedTarget = targetDepotId != null ? String(targetDepotId).trim() : '';
      if (!normalizedTarget) return false;
      const daysToCheck = dayKey && DAY_KEYS.includes(dayKey) ? [dayKey] : DAY_KEYS;
      const scheduleSource = sourceMode
        ? getStoreForMode(sourceMode).scheduleData
        : getScheduleData();
      return daysToCheck.some((day) => {
        const rows = scheduleSource[day] || [];
        return rows.some((order) => {
          if (!order || typeof order !== 'object') return false;
          const rawDepot = order.depot_id != null ? String(order.depot_id).trim() : '';
          return rawDepot === normalizedTarget;
        });
      });
    }

    function determinePreferredStartLocationId(options = {}) {
      const dayKey = options.dayKey && DAY_KEYS.includes(options.dayKey) ? options.dayKey : state.activeDay;
      const rules = options.sourceMode ? getModeConfig(options.sourceMode).rules : getModeRules();
      const priorityDepot = rules.priorityDepotId;
      if (priorityDepot && hasOrdersWithDepotId(priorityDepot, dayKey, options.sourceMode)) {
        return rules.priorityStartLocationId || rules.defaultStartLocationId || '';
      }
      return rules.defaultStartLocationId || '';
    }

    function applyAutoStartSelectionForSource(sourceMode, options = {}) {
      if (!shouldLoadPersistedForMode(sourceMode)) return;
      const store = getStoreForMode(sourceMode);
      const rules = getModeConfig(sourceMode).rules;
      const dayKey = options.dayKey && DAY_KEYS.includes(options.dayKey) ? options.dayKey : state.activeDay;
      const preferredStartId = normalizeStartId(
        options.preferredStartId != null
          ? options.preferredStartId
          : determinePreferredStartLocationId({ dayKey, sourceMode })
      );
      const defaultExtraId = normalizeStartId(rules.defaultExtraStartId);
      const startRows = store.startLocations || [];

      if (options.resetSelection === true) {
        const toRemove = [];
        state.selected.forEach((uid) => {
          if (typeof uid === 'string' && uid.startsWith(`${sourceMode}::start::`)) {
            toRemove.push(uid);
          }
        });
        toRemove.forEach((uid) => state.selected.delete(uid));
      } else if (!options.keepOtherStarts) {
        const toRemove = [];
        state.selected.forEach((uid) => {
          if (!uid || typeof uid !== 'string' || !uid.startsWith(`${sourceMode}::start::`)) return;
          const idx = Number(uid.split('::')[2]);
          if (!Number.isFinite(idx)) {
            toRemove.push(uid);
            return;
          }
          const item = startRows[idx];
          const idValue = normalizeStartId(item && item.id);
          if (idValue === defaultExtraId) return;
          toRemove.push(uid);
        });
        toRemove.forEach((uid) => state.selected.delete(uid));
      }

      if (options.selectDefaultExtra) {
        startRows.forEach((item, index) => {
          if (normalizeStartId(item && item.id) !== defaultExtraId) return;
          state.selected.add(startSelectionUid(sourceMode, index));
        });
      }

      const targetIndex = startRows.findIndex((item) => {
        if (!item || typeof item !== 'object') return false;
        return normalizeStartId(item.id) === preferredStartId;
      });
      if (targetIndex !== -1) {
        state.selected.add(startSelectionUid(sourceMode, targetIndex));
      }
    }

    function applyAutoStartSelectionAll(options = {}) {
      if (!shouldLoadPersistedForMode(ALL_MODE)) return;
      if (options.resetSelection === true) {
        const toRemove = [];
        state.selected.forEach((uid) => {
          if (isStartSelectionUid(uid)) toRemove.push(uid);
        });
        toRemove.forEach((uid) => state.selected.delete(uid));
      }
      const selectDefaultExtra = options.selectDefaultExtra ?? options.resetSelection === true;
      const dayKey = options.dayKey && DAY_KEYS.includes(options.dayKey) ? options.dayKey : state.activeDay;
      SOURCE_MODE_IDS.forEach((sourceMode) => {
        if (!shouldLoadPersistedForMode(sourceMode)) return;
        const rules = getModeConfig(sourceMode).rules;
        const defaultExtraId = normalizeStartId(rules.defaultExtraStartId);
        if (selectDefaultExtra && defaultExtraId) {
          state.selected.add(`shared::start::${defaultExtraId}`);
        }
        const preferredId = normalizeStartId(
          determinePreferredStartLocationId({ dayKey, sourceMode })
        );
        if (preferredId) state.selected.add(`shared::start::${preferredId}`);
      });
    }

    function applyAutoStartSelection(options = {}) {
      if (isAllMode()) {
        applyAutoStartSelectionAll(options);
        return;
      }
      if (!shouldLoadPersistedForMode(getActiveMode())) return;
      const dayKey = options.dayKey && DAY_KEYS.includes(options.dayKey) ? options.dayKey : state.activeDay;
      const preferredStartId = normalizeStartId(
        options.preferredStartId != null ? options.preferredStartId : determinePreferredStartLocationId({ dayKey })
      );
      const defaultExtraId = normalizeStartId(getModeRules().defaultExtraStartId);
      const selectDefaultExtra = options.selectDefaultExtra ?? options.resetSelection === true;
      const startRows = dataStore.startLocations || [];

      if (options.resetSelection === true) {
        state.selected.clear();
      } else {
        const toRemove = [];
        Array.from(state.selected).forEach((uid) => {
          if (!uid || typeof uid !== 'string' || !uid.startsWith('start::')) return;
          const idx = Number(uid.split('::')[1]);
          if (!Number.isFinite(idx)) {
            toRemove.push(uid);
            return;
          }
          const item = startRows[idx];
          const idValue = normalizeStartId(item && item.id);
          if (idValue === defaultExtraId) return;
          toRemove.push(uid);
        });
        toRemove.forEach((uid) => state.selected.delete(uid));
      }

      if (selectDefaultExtra) {
        ensureDefaultExtraStartSelected();
      }

      const targetIndex = startRows.findIndex((item) => {
        if (!item || typeof item !== 'object') return false;
        return normalizeStartId(item.id) === preferredStartId;
      });

      const startLocationExists = targetIndex !== -1;
      if (startLocationExists) {
        const uid = startLocationUid(startRows[targetIndex], targetIndex);
        state.selected.add(uid);
      }

      if (!startLocationExists || !preferredStartId) return;

      (dataStore.vehicles || []).forEach((vehicle) => {
        if (!vehicle || typeof vehicle !== 'object') return;
        const currentValue = vehicle.start_at != null ? String(vehicle.start_at).trim() : '';
        const desiredStartAt = currentValue.toLowerCase().startsWith('start:')
          ? `start:${preferredStartId}`
          : preferredStartId;
        if (currentValue !== desiredStartAt) {
          vehicle.start_at = desiredStartAt;
        }
      });
      saveLocal();
    }

    function parseSheetDataset(payload, sheetNameByDay) {
      const schedule = createEmptySchedule();
      let vehicles = [];
      let depots = [];
      let startLocations = [];
      const sheetNameToCanon = buildSheetNameToCanon(sheetNameByDay || SHEET_NAME_BY_DAY);

      if (!payload || typeof payload !== 'object') {
        return { schedule, vehicles, depots, startLocations };
      }

      Object.entries(payload).forEach(([sheetName, sheetData]) => {
        const lower = String(sheetName || '').trim().toLowerCase();
        if (sheetNameMatches(lower, SHEET_META.vehicles || 'vehicles')) {
          vehicles = parseStructuredSheet(sheetData, normalizeVehicleRecord);
          return;
        }
        if (sheetNameMatches(lower, SHEET_META.depots || ['depot', 'depots'])) {
          depots = parseStructuredSheet(sheetData, normalizeDepotRecord);
          return;
        }
        if (sheetNameMatches(lower, SHEET_META.startLocations || ['startdata', 'startlocations', 'start_points'])) {
          startLocations = parseStructuredSheet(sheetData, normalizeStartRecord);
          return;
        }
        const canon = sheetNameToCanon[lower];
        if (!canon) return;
        schedule[canon] = parseStructuredSheet(sheetData, normalizeOrderRecord);
      });

      return { schedule, vehicles, depots, startLocations };
    }

    function renderSchemaDocs() {
      if (!dom.schemaDocs) return;
      dom.schemaDocs.innerHTML = '';
      Object.values(EXPORT_SCHEMAS).forEach((sheet) => {
        const block = document.createElement('div');
        block.className = 'schema-block';
        block.innerHTML = `
          <h3>${sheet.title}</h3>
          <p class="muted">${sheet.description}</p>
        `;
        const grid = document.createElement('div');
        grid.className = 'schema-grid';
        sheet.columns.forEach((column) => {
          const item = document.createElement('div');
          item.className = 'schema-item';
          item.innerHTML = `
            <strong>${column.label || '—'}</strong>
            <div class="muted">Поле: <code>${column.key || '—'}</code></div>
            ${column.description ? `<div class="muted">${column.description}</div>` : ''}
          `;
          grid.appendChild(item);
        });
        block.appendChild(grid);
        dom.schemaDocs.appendChild(block);
      });
    }

    function getExtraOrderFieldConfig(key) {
      return EXTRA_ORDER_FORM_FIELDS.find((cfg) => cfg.key === key);
    }

    function makeExtraOrderField(label, key, value, options = {}) {
      const placeholderAttr = options.placeholder ? ` placeholder="${escapeHtml(options.placeholder)}"` : '';
      const dataAttr = `data-field="${key}"`;
      const type = options.type === 'number' ? 'number' : 'text';
      const stepAttr = options.step ? ` step="${options.step}"` : '';
      let stringValue = value ?? '';
      stringValue = stringValue === '' ? '' : String(stringValue);
      if (options.phoneMask && stringValue) {
        stringValue = formatRuPhoneInputValue(stringValue);
      }
      const requiredAttr = options.required ? ' required aria-required="true"' : '';
      const fullClass = options.fullWidth ? ' eo-field--full' : '';
      const phoneClass = options.phoneMask ? ' eo-input--phone' : '';
      const inputModeAttr = options.phoneMask ? ' inputmode="tel" autocomplete="tel"' : '';
      const maxLengthAttr = options.phoneMask ? ' maxlength="18"' : '';
      const requiredMark = options.required ? '<span class="eo-required" aria-hidden="true">*</span>' : '';
      return `
        <div class="eo-field${fullClass}">
          <label class="eo-field__label" for="eo-${key}-${options.blockIndex ?? 0}">${escapeHtml(label)}${requiredMark}</label>
          <input
            id="eo-${key}-${options.blockIndex ?? 0}"
            class="eo-input${phoneClass}"
            type="${type}"
            ${dataAttr}${placeholderAttr}${stepAttr}${requiredAttr}${inputModeAttr}${maxLengthAttr}
            value="${escapeHtml(stringValue)}"
          />
        </div>`;
    }

    function makeExtraOrderFieldHtml(key, values, blockIndex) {
      const cfg = getExtraOrderFieldConfig(key);
      if (!cfg) return '';
      const val = values[cfg.key] ?? (cfg.key === 'time_window' ? getDefaultTimeWindow() : '');
      return makeExtraOrderField(cfg.label, cfg.key, val, { ...cfg, blockIndex });
    }

    function getDepotsForExtraOrderPicker() {
      if (isAllMode()) {
        return getMergedUniqueDepotsWithMeta()
          .map((entry) => ({
            id: entry.depot && entry.depot.id != null ? String(entry.depot.id).trim() : '',
            ref: entry.depot && entry.depot.ref != null ? String(entry.depot.ref).trim() : '',
            sourceLabels: entry.sourceLabels || []
          }))
          .filter((d) => d.id);
      }
      return (dataStore.depots || [])
        .map((depot) => ({
          id: depot && depot.id != null ? String(depot.id).trim() : '',
          ref: depot && depot.ref != null ? String(depot.ref).trim() : '',
          sourceLabels: []
        }))
        .filter((d) => d.id);
    }

    function makeExtraOrderDepotPickerHtml(blockIndex, values = {}) {
      const depots = getDepotsForExtraOrderPicker();
      const selectedId = values.depot_id != null ? String(values.depot_id).trim() : '';
      if (!depots.length) {
        return '<p class="eo-depot-picker__empty eo-depot-picker__empty--inline">Склады появятся после загрузки из Google Sheets</p>';
      }
      const chips = depots.map((depot) => {
        const label = depot.ref || depot.id;
        const meta = depot.ref && depot.ref !== depot.id ? depot.id : '';
        const source = depot.sourceLabels.length ? formatSourceLabels(depot.sourceLabels) : '';
        const titleParts = [label, meta, source].filter(Boolean);
        const isActive = selectedId && depot.id === selectedId;
        return `<button
          type="button"
          class="eo-depot-chip${isActive ? ' eo-depot-chip--active' : ''}"
          data-act="pick-depot"
          data-depot-id="${escapeHtml(depot.id)}"
          data-depot-ref="${escapeHtml(depot.ref)}"
          title="${escapeHtml(titleParts.join(' · '))}"
        >${escapeHtml(label)}${meta ? `<span class="eo-depot-chip__meta">${escapeHtml(meta)}</span>` : ''}</button>`;
      }).join('');
      return `<div class="eo-depot-picker eo-depot-picker--inline" role="group" aria-label="Быстрый выбор склада">${chips}</div>`;
    }

    function applyDepotPickToBlock(block, depotId, depotRef) {
      if (!block) return;
      const idInput = block.querySelector('[data-field="depot_id"]');
      const nameInput = block.querySelector('[data-field="depot_name"]');
      if (idInput) idInput.value = depotId || '';
      if (nameInput) nameInput.value = depotRef || '';
      idInput?.closest('.eo-field')?.classList.remove('eo-field--error');
      block.querySelectorAll('[data-act="pick-depot"]').forEach((chip) => {
        chip.classList.toggle('eo-depot-chip--active', chip.dataset.depotId === depotId);
      });
    }

    function makeExtraOrderDepotFieldHtml(values, blockIndex) {
      const cfg = getExtraOrderFieldConfig('depot_id');
      const val = values.depot_id ?? '';
      const pickerHtml = makeExtraOrderDepotPickerHtml(blockIndex, values);
      const isEmptyMsg = pickerHtml.includes('eo-depot-picker__empty');
      return `
        <div class="eo-field eo-field--depot">
          <label class="eo-field__label" for="eo-depot_id-${blockIndex}">${escapeHtml(cfg.label)}</label>
          <div class="eo-depot-inline${isEmptyMsg ? ' eo-depot-inline--empty' : ''}">
            <input
              id="eo-depot_id-${blockIndex}"
              class="eo-input eo-input--depot-id"
              type="text"
              data-field="depot_id"
              placeholder="${escapeHtml(cfg.placeholder || '')}"
              value="${escapeHtml(val === '' ? '' : String(val))}"
            />
            ${pickerHtml}
          </div>
        </div>`;
    }

    function renderExtraOrderPointGrid(values, blockIndex) {
      return `
        <div class="eo-point-grid">
          <div class="eo-point-row eo-point-row--2">
            ${makeExtraOrderFieldHtml('title', values, blockIndex)}
            ${makeExtraOrderFieldHtml('address', values, blockIndex)}
          </div>
          <div class="eo-point-row eo-point-row--2">
            ${makeExtraOrderFieldHtml('coords_paste', values, blockIndex)}
            <div class="eo-coords-pair">
              ${makeExtraOrderFieldHtml('lat', values, blockIndex)}
              ${makeExtraOrderFieldHtml('lng', values, blockIndex)}
            </div>
          </div>
          <div class="eo-point-row eo-point-row--2">
            ${makeExtraOrderFieldHtml('time_window', values, blockIndex)}
            ${makeExtraOrderFieldHtml('delivery_minutes', values, blockIndex)}
          </div>
          <div class="eo-point-row eo-point-row--2">
            ${makeExtraOrderDepotFieldHtml(values, blockIndex)}
            ${makeExtraOrderFieldHtml('depot_name', values, blockIndex)}
          </div>
          <div class="eo-point-row eo-point-row--2">
            ${makeExtraOrderFieldHtml('phone', values, blockIndex)}
            ${makeExtraOrderFieldHtml('comments', values, blockIndex)}
          </div>
        </div>`;
    }

    function syncExtraOrderModalLayout() {
      if (!dom.extraOrderModal || !dom.extraOrderFormBlocks) return;
      const count = dom.extraOrderFormBlocks.querySelectorAll('.extra-order-block').length;
      dom.extraOrderModal.classList.toggle('extra-order-modal--single', count <= 1);
      dom.extraOrderModal.classList.toggle('extra-order-modal--multi', count > 1);
    }

    function makeExtraOrderFormBlockHtml(blockIndex, values = {}) {
      const removeBtn = blockIndex > 0
        ? `<button type="button" class="eo-card-remove" data-act="remove-block">Убрать</button>`
        : '';
      return `
        <div class="extra-order-block" data-block-index="${blockIndex}">
          <div class="extra-order-block__head">
            <h3 class="extra-order-block__title">Точка ${blockIndex + 1}</h3>
            ${removeBtn}
          </div>
          <div class="extra-order-block__body">
            ${renderExtraOrderPointGrid(values, blockIndex)}
          </div>
          <p class="eo-coords-hint" data-coords-hint hidden>Формат: 55.706284, 37.781800 или через пробел</p>
        </div>`;
    }

    function clearExtraOrderFieldErrors() {
      if (!dom.extraOrderFormBlocks) return;
      dom.extraOrderFormBlocks.querySelectorAll('.eo-field--error').forEach((el) => {
        el.classList.remove('eo-field--error');
      });
      dom.extraOrderFormBlocks.querySelectorAll('[data-coords-hint]').forEach((el) => {
        el.hidden = true;
      });
    }

    function markExtraOrderFieldError(block, fieldKey) {
      if (!block) return;
      const input = block.querySelector(`[data-field="${fieldKey}"]`);
      const field = input && input.closest('.eo-field');
      if (field) field.classList.add('eo-field--error');
    }

    function handleExtraOrderCoordsInput(block, text) {
      if (!block) return;
      const hint = block.querySelector('[data-coords-hint]');
      const field = block.querySelector('[data-field="coords_paste"]')?.closest('.eo-field');
      const trimmed = (text || '').trim();
      if (!trimmed) {
        if (hint) hint.hidden = true;
        if (field) field.classList.remove('eo-field--error');
        return;
      }
      if (tryApplyCoordsTextToBlock(block, trimmed)) {
        if (hint) hint.hidden = true;
        if (field) field.classList.remove('eo-field--error');
      } else {
        if (hint) hint.hidden = false;
        if (field) field.classList.add('eo-field--error');
      }
    }

    function readExtraOrderBlockValues(blockEl) {
      const values = {};
      blockEl.querySelectorAll('[data-field]').forEach((input) => {
        const key = input.getAttribute('data-field');
        if (!key) return;
        values[key] = input.type === 'number' ? input.value.trim() : input.value.trim();
      });
      return values;
    }

    function resetExtraOrderModalForm() {
      if (!dom.extraOrderFormBlocks) return;
      dom.extraOrderFormBlocks.innerHTML = makeExtraOrderFormBlockHtml(0);
      syncExtraOrderModalLayout();
    }

    function appendExtraOrderFormBlock() {
      if (!dom.extraOrderFormBlocks) return;
      const index = dom.extraOrderFormBlocks.querySelectorAll('.extra-order-block').length;
      dom.extraOrderFormBlocks.insertAdjacentHTML('beforeend', makeExtraOrderFormBlockHtml(index));
      syncExtraOrderModalLayout();
    }

    function openExtraOrderModal() {
      if (!shouldLoadPersistedForMode(getActiveMode())) {
        showError('Сначала обновите данные из Google Sheets.');
        return;
      }
      if (!dom.extraOrderModal) return;
      clearError();
      resetExtraOrderModalForm();
      if (dom.extraOrderModalDay) {
        dom.extraOrderModalDay.textContent = WEEKDAY_LABELS[state.activeDay] || state.activeDay;
      }
      dom.extraOrderModal.hidden = false;
      dom.extraOrderModal.setAttribute('aria-hidden', 'false');
      lockPageScrollForExtraOrderModal();
      syncExtraOrderModalLayout();
      const firstInput = dom.extraOrderFormBlocks && dom.extraOrderFormBlocks.querySelector('input[data-field="title"]');
      if (firstInput) firstInput.focus();
    }

    function closeExtraOrderModal() {
      if (!dom.extraOrderModal) return;
      dom.extraOrderModal.hidden = true;
      dom.extraOrderModal.setAttribute('aria-hidden', 'true');
      unlockPageScrollForExtraOrderModal();
      if (dom.extraOrderSubmitBtn) dom.extraOrderSubmitBtn.disabled = false;
      clearExtraOrderFieldErrors();
    }

    function collectExtraOrdersFromModal() {
      if (!dom.extraOrderFormBlocks) return [];
      clearExtraOrderFieldErrors();
      const blocks = Array.from(dom.extraOrderFormBlocks.querySelectorAll('.extra-order-block'));
      const records = [];
      const errors = [];
      blocks.forEach((block, blockIndex) => {
        const raw = readExtraOrderBlockValues(block);
        const title = raw.title || '';
        const lat = toNumOrNull(raw.lat);
        const lon = toNumOrNull(raw.lng);
        if (!title.trim()) {
          errors.push(`Точка ${blockIndex + 1}: укажите наименование.`);
          markExtraOrderFieldError(block, 'title');
          return;
        }
        if (lat == null || lon == null) {
          errors.push(`Точка ${blockIndex + 1}: укажите широту и долготу.`);
          markExtraOrderFieldError(block, 'lat');
          markExtraOrderFieldError(block, 'lng');
          return;
        }
        let deliverySeconds = toNumOrNull(raw.delivery_minutes);
        if (deliverySeconds != null) {
          deliverySeconds = deliverySeconds * 60;
        }
        const record = normalizeExtraOrderRecord({
          id: generateExtraOrderId(),
          title: title.trim(),
          address: raw.address || '',
          lat,
          lng: lon,
          phone: raw.phone || '',
          time_window: raw.time_window || getDefaultTimeWindow(),
          comments: raw.comments || '',
          depot_id: raw.depot_id || '',
          depot_name: raw.depot_name || '',
          delivery_seconds: deliverySeconds
        });
        if (record) records.push(record);
      });
      if (errors.length) {
        showError(errors.join(' '));
        return null;
      }
      if (!records.length) {
        showError('Заполните хотя бы одну точку.');
        return null;
      }
      return records;
    }

    function submitExtraOrdersFromModal() {
      const submitBtn = dom.extraOrderSubmitBtn;
      if (submitBtn) submitBtn.disabled = true;
      const records = collectExtraOrdersFromModal();
      if (!records) {
        if (submitBtn) submitBtn.disabled = false;
        return;
      }
      const day = state.activeDay;
      const bucket = getExtraOrders()[day] || (getExtraOrders()[day] = []);
      records.forEach((record) => {
        bucket.push(record);
        state.selected.add(extraOrderUid(day, record.id));
      });
      saveLocal();
      persistActiveSheetCacheSelection();
      closeExtraOrderModal();
      if (submitBtn) submitBtn.disabled = false;
      render();
    }

    function initExtraOrderModal() {
      if (!dom.extraOrderModal) return;
      const closeHandlers = [
        dom.extraOrderModalBackdrop,
        dom.extraOrderModalClose,
        dom.extraOrderCancelBtn
      ];
      closeHandlers.forEach((el) => {
        if (el) el.addEventListener('click', () => closeExtraOrderModal());
      });
      if (dom.extraOrderAddBlockBtn) {
        dom.extraOrderAddBlockBtn.addEventListener('click', () => appendExtraOrderFormBlock());
      }
      if (dom.extraOrderSubmitBtn) {
        dom.extraOrderSubmitBtn.addEventListener('click', () => submitExtraOrdersFromModal());
      }
      if (dom.extraOrderFormBlocks) {
        dom.extraOrderFormBlocks.addEventListener('click', (event) => {
          const depotChip = event.target.closest('[data-act="pick-depot"]');
          if (depotChip) {
            const block = depotChip.closest('.extra-order-block');
            applyDepotPickToBlock(block, depotChip.dataset.depotId || '', depotChip.dataset.depotRef || '');
            return;
          }
          const btn = event.target.closest('[data-act="remove-block"]');
          if (!btn) return;
          const block = btn.closest('.extra-order-block');
          if (!block) return;
          block.remove();
          dom.extraOrderFormBlocks.querySelectorAll('.extra-order-block').forEach((el, index) => {
            el.dataset.blockIndex = String(index);
            const titleEl = el.querySelector('.extra-order-block__title');
            if (titleEl) titleEl.textContent = `Точка ${index + 1}`;
          });
          syncExtraOrderModalLayout();
        });

        dom.extraOrderFormBlocks.addEventListener('input', (event) => {
          const phoneInput = event.target.closest('input[data-field="phone"]');
          if (phoneInput) {
            handleExtraOrderPhoneInput(phoneInput);
            phoneInput.closest('.eo-field')?.classList.remove('eo-field--error');
            return;
          }
          const depotInput = event.target.closest('input[data-field="depot_id"]');
          if (depotInput) {
            const block = depotInput.closest('.extra-order-block');
            const currentId = depotInput.value.trim();
            block?.querySelectorAll('[data-act="pick-depot"]').forEach((chip) => {
              chip.classList.toggle('eo-depot-chip--active', chip.dataset.depotId === currentId);
            });
          }
          const input = event.target.closest('input[data-field="coords_paste"]');
          if (!input) return;
          const block = input.closest('.extra-order-block');
          handleExtraOrderCoordsInput(block, input.value);
        });

        dom.extraOrderFormBlocks.addEventListener('paste', (event) => {
          const phoneInput = event.target.closest('input[data-field="phone"]');
          if (phoneInput) {
            event.preventDefault();
            const pasted = event.clipboardData && event.clipboardData.getData('text');
            phoneInput.value = formatRuPhoneMaskFromDigits(extractRuPhoneDigits(pasted));
            phoneInput.setSelectionRange(phoneInput.value.length, phoneInput.value.length);
            return;
          }
          const input = event.target.closest('input[data-field]');
          if (!input) return;
          const field = input.getAttribute('data-field');
          if (field !== 'coords_paste' && field !== 'lat' && field !== 'lng') return;
          const block = input.closest('.extra-order-block');
          if (!block) return;
          const text = event.clipboardData && event.clipboardData.getData('text');
          if (!text || !parseLatLngPair(text)) return;
          event.preventDefault();
          tryApplyCoordsTextToBlock(block, text);
          handleExtraOrderCoordsInput(block, text);
        });

        dom.extraOrderFormBlocks.addEventListener('focusin', (event) => {
          const phoneInput = event.target.closest('input[data-field="phone"]');
          if (phoneInput) handleExtraOrderPhoneFocus(phoneInput);
        });

        dom.extraOrderFormBlocks.addEventListener('focusout', (event) => {
          const phoneInput = event.target.closest('input[data-field="phone"]');
          if (phoneInput) handleExtraOrderPhoneBlur(phoneInput);
        });

        dom.extraOrderFormBlocks.addEventListener('keydown', (event) => {
          const phoneInput = event.target.closest('input[data-field="phone"]');
          if (!phoneInput || event.key !== 'Backspace') return;
          const digits = extractRuPhoneDigits(phoneInput.value);
          if (digits.length <= 1) {
            event.preventDefault();
            phoneInput.value = '';
          }
        });

        dom.extraOrderFormBlocks.addEventListener('blur', (event) => {
          const input = event.target.closest('input[data-field="coords_paste"]');
          if (!input) return;
          const block = input.closest('.extra-order-block');
          handleExtraOrderCoordsInput(block, input.value);
        }, true);

        dom.extraOrderFormBlocks.addEventListener('input', (event) => {
          const input = event.target.closest('.eo-input[data-field]');
          if (!input) return;
          input.closest('.eo-field')?.classList.remove('eo-field--error');
        });
      }
      document.addEventListener('keydown', (ev) => {
        if (ev.key === 'Escape' && dom.extraOrderModal && !dom.extraOrderModal.hidden) {
          closeExtraOrderModal();
        }
      });
    }

    // ===== Рендер =====
    function setActiveDay(day) {
      state.activeDay = day;
      dom.tabs.forEach((btn) => {
        const active = btn.dataset.day === day;
        btn.classList.toggle('active', active);
        btn.setAttribute('aria-selected', active ? 'true' : 'false');
      });
      state.selected.clear();
      if (shouldLoadPersistedForMode(getActiveMode())) {
        applyAutoStartSelection({ dayKey: day, resetSelection: false, selectDefaultExtra: true });
      }
      persistActiveSheetCacheSelection();
      updateSelectedCount();
      state.lastIndexByDay[day] = undefined;
      dom.storesSection?.classList.add('is-day-switching');
      render();
      requestAnimationFrame(() => {
        dom.storesSection?.classList.remove('is-day-switching');
      });
      renderVehicles();
    }

    function refreshModeUiAfterDataChange() {
      render();
      renderVehicles();
      renderDepots();
      renderStartLocations();
      updateSelectedCount();
    }

    function render() {
      if (!dom.root) return;
      const day = state.activeDay;
      const combinedList = buildDayListItems(day);
      const canAddExtra = shouldLoadPersistedForMode(getActiveMode());
      const query = state.query.trim().toLowerCase();
      const list = query
        ? combinedList.filter((item) =>
            `${item.title || item.store || ''} ${item.address || ''} ${toStrPhone(item.phone)} ${item.__sourceLabel || ''}`
              .toLowerCase()
              .includes(query)
          )
        : combinedList;

      const totalVisible = list.reduce((acc, item) => {
        if (item && item.__isStart) {
          const rules = item.__sourceMode ? getModeConfig(item.__sourceMode).rules : getModeRules();
          const excludedStartIds = new Set([
            normalizeStartId(rules.priorityStartLocationId),
            normalizeStartId(rules.defaultStartLocationId),
            normalizeStartId(rules.defaultExtraStartId)
          ].filter(Boolean));
          const idValue = normalizeStartId(item.id);
          if (!idValue || excludedStartIds.has(idValue)) return acc;
        }
        return acc + 1;
      }, 0);

      dom.root.innerHTML = '';

      const headingWrap = document.createElement('div');
      headingWrap.className = 'sticky-header';

      const topControls = document.createElement('div');
      topControls.className = 'header-controls';
      topControls.innerHTML = `
        <div class="header-controls__row header-controls__row--title">
          <h2 class="day-heading">
            <span>${WEEKDAY_LABELS[day]}</span>
            <span class="badge muted">Всего ${totalVisible}</span>
            <span class="badge" id="selectedCount">Выбрано: ${state.selected.size}</span>
          </h2>
        </div>
        <div class="header-controls__row header-controls__row--actions">
          <input id="searchSticky" type="search" placeholder="Поиск: название, адрес или телефон…" />
          <button id="addExtraOrder" class="mini-btn" type="button" ${canAddExtra ? '' : 'disabled title="Сначала обновите данные из Google Sheets"'}>+ Разовая точка</button>
          <button id="selectAll">Выбрать все</button>
          <button id="clearAll">Снять выбор</button>
          <button id="exportXlsx" type="button" class="mini-btn export-btn btn-action${hasExportableSelection() ? ' export-btn--ready' : ''}"${hasExportableSelection() ? '' : ' disabled'}>
            <span class="btn-action__spinner" aria-hidden="true"></span>
            <span class="btn-action__icon" aria-hidden="true"></span>
            <span class="btn-action__label">Скачать таблицу для планирования</span>
          </button>
        </div>
      `;
      headingWrap.appendChild(topControls);
      dom.root.appendChild(headingWrap);

      const searchSticky = topControls.querySelector('#searchSticky');
      if (searchSticky) {
        searchSticky.value = state.query;
        searchSticky.addEventListener('input', () => {
          const { selectionStart, selectionEnd } = searchSticky;
          state.query = searchSticky.value;
          render();
          const next = document.getElementById('searchSticky');
          if (next) {
            next.focus();
            try { next.setSelectionRange(selectionStart, selectionEnd); } catch (_) {}
          }
        });
      }
      const exportStickyBtn = topControls.querySelector('#exportXlsx');
      if (exportStickyBtn) {
        exportStickyBtn.addEventListener('click', exportSelectedXlsx);
      }
      const addExtraBtn = topControls.querySelector('#addExtraOrder');
      if (addExtraBtn && canAddExtra) {
        addExtraBtn.addEventListener('click', () => openExtraOrderModal());
      }

      const listEl = document.createElement('ul');
      listEl.className = 'stores';
      if (!list.length) {
        const empty = document.createElement('p');
        empty.className = 'empty-hint stores-empty';
        empty.textContent = query
          ? 'Ничего не найдено. Попробуйте другой запрос.'
          : 'На этот день пока нет точек. Обновите данные из Google Sheets.';
        dom.root.appendChild(empty);
      } else {
        mountStoreListLazy(listEl, list, day);
        dom.root.appendChild(listEl);
        if (list.length > LIST_LAZY_BATCH) {
          const lazyHint = document.createElement('p');
          lazyHint.className = 'stores-lazy-hint muted';
          lazyHint.textContent = `Всего ${list.length} · подгружаем список при прокрутке`;
          dom.root.appendChild(lazyHint);
        }
      }

      topControls.querySelector('#selectAll').addEventListener('click', () => bulkSelect(true));
      topControls.querySelector('#clearAll').addEventListener('click', () => bulkSelect(false));

      if (!list.length) {
        updateSelectedCount();
        updateSelectedHighlight();
        return;
      }

      listEl.addEventListener('mousedown', (event) => {
        if (event.button !== 0) return;
        if (event.target.closest('a')) return;
        if (event.target.closest('[data-act="del-extra"]')) return;
        const item = event.target.closest('li.store');
        if (!item) return;
        event.preventDefault();
        document.body.classList.add('no-select');

        const dayKey = state.activeDay;
        const startIdx = Number(item.dataset.index);
        const startBox = item.querySelector('input[type="checkbox"][data-uid]');
        if (!startBox) return;
        const targetChecked = !startBox.checked;
        let lastIdx = startIdx;

        const applyRange = (from, to) => {
          const start = Math.min(from, to);
          const end = Math.max(from, to);
          for (let i = start; i <= end; i += 1) {
            const candidate = listEl.querySelector(`li.store[data-index="${i}"]`);
            if (!candidate) continue;
            const box = candidate.querySelector('input[type="checkbox"][data-uid]');
            if (!box) continue;
            candidate.classList.add('range-hl');
            setTimeout(() => candidate.classList.remove('range-hl'), 500);
            if (box.checked !== targetChecked) {
              box.checked = targetChecked;
              box.dispatchEvent(new Event('change', { bubbles: true }));
            }
          }
        };

        applyRange(startIdx, startIdx);

        const onMove = (moveEvent) => {
          const element = document.elementFromPoint(moveEvent.clientX, moveEvent.clientY);
          const li = element && element.closest ? element.closest('li.store') : null;
          if (!li) return;
          const idx = Number(li.dataset.index);
          if (!Number.isFinite(idx) || idx === lastIdx) return;
          applyRange(startIdx, idx);
          lastIdx = idx;
        };

        const onUp = () => {
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);
          document.body.classList.remove('no-select');
          state.lastIndexByDay[dayKey] = lastIdx;
          state.suppressClickAfterDrag = true;
          setTimeout(() => {
            state.suppressClickAfterDrag = false;
          }, 0);
        };

        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });

      listEl.addEventListener('click', (event) => {
        if (state.suppressClickAfterDrag) return;
        if (event.target.closest('a')) return;
        if (event.target.closest('[data-act="del-extra"]')) return;
        const rowEl = event.target.closest('li.store');
        if (!rowEl) return;
        const checkbox = rowEl.querySelector('input[type="checkbox"][data-uid]');
        if (!checkbox) return;

        const dayKey = state.activeDay;
        const idx = Number(rowEl.dataset.index);
        const isShift = event.shiftKey === true;
        let targetChecked;

        if (event.target.closest('input[type="checkbox"]')) {
          targetChecked = checkbox.checked;
        } else {
          targetChecked = !checkbox.checked;
          checkbox.checked = targetChecked;
        }

        if (isShift) {
          const selection = window.getSelection && window.getSelection();
          if (selection && selection.removeAllRanges) selection.removeAllRanges();
          document.body.classList.add('no-select');
          event.preventDefault();
        }

        if (isShift && Number.isFinite(state.lastIndexByDay[dayKey])) {
          const start = Math.min(state.lastIndexByDay[dayKey], idx);
          const end = Math.max(state.lastIndexByDay[dayKey], idx);
          for (let i = start; i <= end; i += 1) {
            const candidate = listEl.querySelector(`li.store[data-index="${i}"]`);
            if (!candidate) continue;
            const box = candidate.querySelector('input[type="checkbox"][data-uid]');
            if (!box) continue;
            candidate.classList.add('range-hl');
            setTimeout(() => candidate.classList.remove('range-hl'), 700);
            if (box.checked !== targetChecked) {
              box.checked = targetChecked;
              box.dispatchEvent(new Event('change', { bubbles: true }));
            }
          }
        } else {
          checkbox.dispatchEvent(new Event('change', { bubbles: true }));
        }

        state.lastIndexByDay[dayKey] = idx;
        if (isShift) setTimeout(() => document.body.classList.remove('no-select'), 150);
      });

      listEl.addEventListener('change', (event) => {
        const checkbox = event.target.closest('input[type="checkbox"][data-uid]');
        if (!checkbox) return;
        const rowEl = checkbox.closest('li.store');
        const uid = checkbox.dataset.uid;
        const isStart = rowEl && rowEl.dataset.isStart === 'true';
        if (isStart) {
          const uid = checkbox.dataset.uid;
          if (isAllMode() && uid && uid.startsWith('shared::start::')) {
            if (checkbox.checked) state.selected.add(uid);
            else state.selected.delete(uid);
            updateSelectedCount();
            updateSelectedHighlight();
            persistActiveSheetCacheSelection();
            return;
          }
          const sourceMode = rowEl && rowEl.dataset.sourceMode ? rowEl.dataset.sourceMode : '';
          const rules = sourceMode && isAllMode() ? getModeConfig(sourceMode).rules : getModeRules();
          const startId = normalizeStartId(rowEl ? rowEl.dataset.startId : '');
          const defaultExtraId = normalizeStartId(rules.defaultExtraStartId);
          if (startId && startId === defaultExtraId) {
            if (checkbox.checked) state.selected.add(uid);
            else state.selected.delete(uid);
            updateSelectedCount();
            updateSelectedHighlight();
            persistActiveSheetCacheSelection();
          } else if (isAllMode() && sourceMode) {
            if (checkbox.checked) {
              applyAutoStartSelectionForSource(sourceMode, {
                preferredStartId: startId,
                dayKey: state.activeDay,
                resetSelection: false,
                selectDefaultExtra: false
              });
            } else {
              applyAutoStartSelectionForSource(sourceMode, {
                dayKey: state.activeDay,
                resetSelection: false,
                selectDefaultExtra: false
              });
            }
            render();
            renderVehicles();
            persistActiveSheetCacheSelection();
          } else {
            if (checkbox.checked) {
              applyAutoStartSelection({
                preferredStartId: startId,
                dayKey: state.activeDay,
                resetSelection: false,
                selectDefaultExtra: false
              });
            } else {
              applyAutoStartSelection({
                dayKey: state.activeDay,
                resetSelection: false,
                selectDefaultExtra: false
              });
            }
            render();
            renderVehicles();
            persistActiveSheetCacheSelection();
          }
          return;
        }
        if (checkbox.checked) state.selected.add(uid);
        else state.selected.delete(uid);
        updateSelectedCount();
        updateSelectedHighlight();
        persistActiveSheetCacheSelection();
      });

      updateSelectedCount();
      updateSelectedHighlight();
    }

    function bulkSelect(selectAll) {
      const day = state.activeDay;
      const query = state.query.trim().toLowerCase();
      const combined = buildDayListItems(day);
      const filtered = query
        ? combined.filter((item) =>
            `${item.title || item.store || ''} ${item.address || ''} ${toStrPhone(item.phone)}`
              .toLowerCase()
              .includes(query)
          )
        : combined;
      filtered.forEach((item) => {
        if (selectAll) state.selected.add(item.uid);
        else state.selected.delete(item.uid);
      });
      if (shouldLoadPersistedForMode(getActiveMode())) {
        applyAutoStartSelection({ dayKey: state.activeDay, resetSelection: false, selectDefaultExtra: false });
      }
      persistActiveSheetCacheSelection();
      render();
    }

    // ===== Экспорт =====
    function pushSelectedStartRow(rows, entry, uid, index, timeWindow) {
      if (!state.selected.has(uid)) return;
      const lat = entry.lat ?? toNumOrNull(entry['point.lat']);
      const lon = entry.lng ?? toNumOrNull(entry['point.lon']);
      const idValue = entry.id != null && entry.id !== '' ? entry.id : `start-${index + 1}`;
      rows.push({
        id: idValue,
        'point.lat': lat ?? '',
        'point.lon': lon ?? '',
        title: entry.title || entry.ref || entry.id || `Старт #${index + 1}`,
        address: entry.address || '',
        phone: toStrPhone(entry.phone),
        time_window: entry.time_window || timeWindow,
        comments: entry.comments || '',
        hard_window: true,
        shared_service_duration_s: '',
        service_duration_s: '',
        depot_id: entry.depot_id || '',
        type: entry.type || ''
      });
    }

    function pushSelectedOrderRow(rows, counterRef, record, timeWindow) {
      const lat = toNumOrNull(record.lat);
      const lon = toNumOrNull(record.lng);
      const seconds = Number.isFinite(record.delivery_seconds) ? Number(record.delivery_seconds) : '';
      rows.push({
        id: counterRef.value++,
        'point.lat': lat ?? '',
        'point.lon': lon ?? '',
        title: record.title || record.store || '',
        address: record.address || '',
        phone: toStrPhone(record.phone),
        time_window: record.time_window || timeWindow,
        comments: record.comments || '',
        hard_window: true,
        shared_service_duration_s: seconds,
        service_duration_s: seconds,
        depot_id: record.depot_id || '',
        type: record.type || ''
      });
    }

    function gatherSelectedRows() {
      const rows = [];
      const counter = { value: 1 };

      if (isAllMode()) {
        getMergedUniqueStartLocations().forEach((entry, index) => {
          const tw = getModeConfig(entry.sourceModes[0]).rules.defaultTimeWindow || '10:00-21:00';
          const mapped = mapStartLocation(entry.record, index);
          pushSelectedStartRow(rows, mapped, entry.uid, index, tw);
        });
        SOURCE_MODE_IDS.forEach((sourceMode) => {
          if (!shouldLoadPersistedForMode(sourceMode)) return;
          const store = getStoreForMode(sourceMode);
          const tw = getModeConfig(sourceMode).rules.defaultTimeWindow || '10:00-21:00';
          for (const day of DAY_KEYS) {
            (store.scheduleData[day] || []).forEach((record, index) => {
              const uid = scheduleOrderUid(day, record, index, sourceMode);
              if (!state.selected.has(uid)) return;
              pushSelectedOrderRow(rows, counter, record, tw);
            });
            const extras = store.extraOrders || createEmptyExtraOrders();
            (extras[day] || []).forEach((record) => {
              const uid = extraOrderUid(day, record.id, sourceMode);
              if (!state.selected.has(uid)) return;
              pushSelectedOrderRow(rows, counter, record, tw);
            });
          }
        });
        const allStore = getStoreForMode(ALL_MODE);
        const allTw = getDefaultTimeWindow();
        for (const day of DAY_KEYS) {
          (allStore.extraOrders[day] || []).forEach((record) => {
            const uid = extraOrderUid(day, record.id, ALL_MODE);
            if (!state.selected.has(uid)) return;
            pushSelectedOrderRow(rows, counter, record, allTw);
          });
        }
        return rows;
      }

      const startEntries = shouldLoadPersistedForMode(getActiveMode())
        ? (dataStore.startLocations || []).map((record, index) => mapStartLocation(record, index))
        : [];
      startEntries.forEach((entry, index) => {
        const uid = entry.uid || startLocationUid(entry, index);
        pushSelectedStartRow(rows, entry, uid, index, getDefaultTimeWindow());
      });
      for (const day of DAY_KEYS) {
        const source = getScheduleData()[day] || [];
        source.forEach((record, index) => {
          const uid = scheduleOrderUid(day, record, index);
          if (!state.selected.has(uid)) return;
          pushSelectedOrderRow(rows, counter, record, getDefaultTimeWindow());
        });
        (getExtraOrders()[day] || []).forEach((record) => {
          const uid = extraOrderUid(day, record.id);
          if (!state.selected.has(uid)) return;
          pushSelectedOrderRow(rows, counter, record, getDefaultTimeWindow());
        });
      }
      return rows;
    }

    function makeAoA(data, schema) {
      const headerLabels = schema.columns.map((column) => column.label ?? '');
      const body = data.map((item) => schema.columns.map((column) => column.getValue(item)));
      if (schema.includeKeys) {
        const headerKeys = schema.columns.map((column) => column.key ?? '');
        return [headerLabels, headerKeys, ...body];
      }
      return [headerLabels, ...body];
    }

    function buildOrdersAoA(rows) {
      return makeAoA(rows, EXPORT_SCHEMAS.orders);
    }

    function buildVehiclesAoA() {
      return makeAoA(getVehiclesForExport(), EXPORT_SCHEMAS.vehicles);
    }

    function buildDepotAoA() {
      return makeAoA(getDepotsForExport(), EXPORT_SCHEMAS.depots);
    }

    async function exportSelectedXlsx() {
      const btn = ensureActionButton(document.getElementById('exportXlsx'));
      if (!hasExportableSelection()) {
        showError('Выберите хотя бы одну точку доставки (не стартовую).');
        return;
      }
      const rows = gatherSelectedRows();
      if (!rows.length) {
        showError('Выберите хотя бы одну точку для экспорта.');
        return;
      }
      const vehiclesForExport = getVehiclesForExport();
      if (countConfiguredVehicles() && !vehiclesForExport.length) {
        showError('Выберите хотя бы одного водителя в настройках Vehicles для экспорта.');
        return;
      }
      clearError();
      setActionButtonState(btn, 'loading', { loadingText: 'Собираем таблицу…' });
      if (btn) btn.classList.add('export-btn--ready');

      const ok = await ensureXlsxReady();
      if (!ok) {
        showError('Не удалось загрузить библиотеку XLSX.');
        setActionButtonState(btn, 'idle');
        updateExportButtonState();
        return;
      }

      const deliveryDate = getDeliveryDateForDay(state.activeDay);
      const prefix = getModeConfig().exportFilePrefix || 'Заказы';
      const baseFileName = `${prefix}_${formatDateForFile(deliveryDate)}`;
      const fileName = `${baseFileName}.xlsx`;

      try {
        const wb = XLSX.utils.book_new();

        const ordersSheet = XLSX.utils.aoa_to_sheet(buildOrdersAoA(rows));
        ordersSheet['!cols'] = EXPORT_SCHEMAS.orders.columns.map((column) => ({ wch: column.width || 18 }));
        XLSX.utils.book_append_sheet(wb, ordersSheet, EXPORT_SCHEMAS.orders.sheetName);

        const vehiclesSheet = XLSX.utils.aoa_to_sheet(buildVehiclesAoA());
        vehiclesSheet['!cols'] = EXPORT_SCHEMAS.vehicles.columns.map((column) => ({ wch: column.width || 18 }));
        XLSX.utils.book_append_sheet(wb, vehiclesSheet, EXPORT_SCHEMAS.vehicles.sheetName);

        const depotsSheet = XLSX.utils.aoa_to_sheet(buildDepotAoA());
        depotsSheet['!cols'] = EXPORT_SCHEMAS.depots.columns.map((column) => ({ wch: column.width || 18 }));
        XLSX.utils.book_append_sheet(wb, depotsSheet, EXPORT_SCHEMAS.depots.sheetName);

        try {
          XLSX.writeFile(wb, fileName, { compression: true, bookSST: true });
        } catch (writeErr) {
          const blob = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          });
          const link = document.createElement('a');
          link.href = URL.createObjectURL(blob);
          link.download = fileName;
          document.body.appendChild(link);
          link.click();
          URL.revokeObjectURL(link.href);
          link.remove();
        }
        setActionButtonState(btn, 'success', { successText: 'Скачано!' });
        celebrateExport(btn);
        showNotify(`Файл «${fileName}» скачан — можно загружать в планировщик.`, 'success', 5500);
        if (APP.yandexPlanningUrl) {
          window.open(APP.yandexPlanningUrl, '_blank', 'noopener,noreferrer');
        }
        setTimeout(() => {
          setActionButtonState(btn, 'idle');
          updateExportButtonState();
        }, 2200);
      } catch (err) {
        console.error(err);
        showError('Не удалось сформировать Excel-файл. Откройте консоль для подробностей.');
        setActionButtonState(btn, 'idle');
      } finally {
        updateExportButtonState();
      }
    }

    // ===== Импорт данных =====
    async function loadLocalJson(file) {
      try {
        clearError();
        const text = await file.text();
        const json = JSON.parse(text);
        const keys = DAY_KEYS;
        keys.forEach((key) => {
          if (!(key in json)) json[key] = [];
        });
        const nextData = {};
        keys.forEach((key) => {
          nextData[key] = (json[key] || []).map((entry) => normalizeOrderRecord(entry)).filter(Boolean);
        });
        setScheduleData(nextData);
        if (Array.isArray(json.vehicles)) {
          dataStore.vehicles = json.vehicles.map((item) => normalizeVehicleRecord(item)).filter(Boolean);
          if (Array.isArray(json.selectedVehicleIds)) {
            getActiveStores().selectedVehicleIds = new Set(json.selectedVehicleIds.map((id) => String(id)));
          } else {
            rebuildVehicleSelectionAfterLoad();
          }
        }
        if (Array.isArray(json.depots)) {
          dataStore.depots = json.depots.map((item) => normalizeDepotRecord(item)).filter(Boolean);
        }
        const jsonStart = json.startLocations || json.start_locations || json.startData || json.StartData;
        dataStore.startLocations = Array.isArray(jsonStart)
          ? jsonStart.map((item) => normalizeStartRecord(item)).filter(Boolean)
          : [];
        markSheetSynced(getActiveMode());
        state.selected.clear();
        state.query = '';
        applyAutoStartSelection({ resetSelection: true, dayKey: state.activeDay, selectDefaultExtra: true });
        render();
        renderVehicles();
        renderDepots();
        renderStartLocations();
        saveLocal();
      } catch (err) {
        showError('Не удалось прочитать JSON: ' + err.message);
      }
    }

    async function loadSheetForMode(sourceMode) {
      const config = getModeConfig(sourceMode);
      const endpoint = config.sheetEndpoint;
      if (!endpoint) {
        throw new Error(
          `Для режима «${config.label}» не настроен Google Sheets (modes.${sourceMode}.sheetEndpoint).`
        );
      }
      const response = await fetch(endpoint, { cache: 'no-store' });
      if (!response.ok) {
        throw new Error(`«${config.label}»: HTTP ${response.status}`);
      }
      const payload = await response.json();
      const parsed = parseSheetDataset(payload, config.sheetNameByDay);
      const store = getStoreForMode(sourceMode);
      store.scheduleData = parsed.schedule;
      store.vehicles = parsed.vehicles;
      store.depots = parsed.depots;
      store.startLocations = parsed.startLocations || [];
      markSheetSynced(sourceMode);
      applyDefaultVehicleExportSelection(sourceMode);
      saveLocalForMode(sourceMode);
    }

    async function loadFromGoogleSheets() {
      const btn = ensureActionButton(dom.loadSheetBtn);
      const modesToLoad = isAllMode() ? SOURCE_MODE_IDS : [getActiveMode()];
      let loadSucceeded = false;
      try {
        clearError();
        setSyncingUi(true);
        setActionButtonState(btn, 'loading', { loadingText: 'Подключаемся…' });
        for (let i = 0; i < modesToLoad.length; i += 1) {
          const sourceMode = modesToLoad[i];
          const label = MODE_LOAD_LABELS[sourceMode] || sourceMode;
          const loadingText = modesToLoad.length > 1
            ? `Загрузка: ${label} (${i + 1}/${modesToLoad.length})…`
            : 'Загружаем из таблицы…';
          setActionButtonState(btn, 'loading', { loadingText });
          await loadSheetForMode(sourceMode);
        }
        loadSucceeded = true;
        dismissSheetOnboarding(true);
        if (isAllMode()) {
          state.selected.clear();
          state.query = '';
          applyAutoStartSelection({ resetSelection: true, dayKey: state.activeDay, selectDefaultExtra: true });
          modesToLoad.forEach((sourceMode) => saveSheetImportCache(sourceMode, { freshImport: true }));
          persistAllModeSelectionCache();
          refreshModeUiAfterDataChange();
        } else {
          rebuildVehicleSelectionAfterLoad();
          state.selected.clear();
          state.query = '';
          applyAutoStartSelection({ resetSelection: true, dayKey: state.activeDay, selectDefaultExtra: true });
          modesToLoad.forEach((sourceMode) => saveSheetImportCache(sourceMode, { freshImport: true }));
          render();
          renderVehicles();
          renderDepots();
          renderStartLocations();
          saveLocal();
        }
      } catch (err) {
        console.error('Sheet load failed', err);
        showError('Не удалось загрузить данные из Google Sheets: ' + err.message);
        setActionButtonState(btn, 'idle');
      } finally {
        setSyncingUi(false);
        if (loadSucceeded) {
          setActionButtonState(btn, 'success', { successText: 'Обновлено' });
          const notifyText = modesToLoad.length > 1
            ? 'Расписание и справочники загружены (HoReCa + Галереи).'
            : `Данные «${getModeConfig(modesToLoad[0]).label}» обновлены.`;
          showNotify(notifyText, 'success');
          setTimeout(() => setActionButtonState(btn, 'idle'), 2000);
        }
      }
    }

    function renderStartLocations() {
      if (!dom.startTableWrap || !dom.startCount) return;
      if (isAllMode()) {
        if (!shouldLoadPersistedForMode(ALL_MODE)) {
          dom.startCount.textContent = '0 точек';
          dom.startTableWrap.innerHTML = '<p class="empty-hint">Сначала обновите данные из Google Sheets.</p>';
          return;
        }
        const merged = getMergedUniqueStartLocations();
        dom.startCount.textContent = `${merged.length} точек`;
        if (!merged.length) {
          dom.startTableWrap.innerHTML = '<p class="empty-hint">Нет стартовых точек. Нажмите «Добавить точку».</p>';
          return;
        }
        const list = document.createElement('div');
        list.className = 'config-list';
        const intro = document.createElement('p');
        intro.className = 'muted empty-hint';
        intro.style.marginBottom = '12px';
        intro.textContent = 'Одинаковые точки показаны один раз. Правки сохраняются локально и попадут в экспорт (не в Google Sheets).';
        list.appendChild(intro);
        merged.forEach((entry) => {
          const row = entry.record;
          const titleRaw = (row.ref && String(row.ref).trim()) || (row.id && String(row.id).trim()) || 'Новая точка';
          const sourceNote = formatSourceLabels(entry.sourceLabels);
          const latVal = row['point.lat'];
          const lonVal = row['point.lon'];
          const subtitleParts = [];
          if (sourceNote) subtitleParts.push(sourceNote);
          if (latVal !== '' && lonVal !== '') subtitleParts.push(`Координаты: ${latVal}, ${lonVal}`);
          if (row.time_window) subtitleParts.push(`Окно: ${row.time_window}`);
          if (row.type) subtitleParts.push(`Тип: ${row.type}`);
          const subtitle = subtitleParts.length ? subtitleParts.join(' • ') : (row.id ? `ID: ${row.id}` : 'Данные не заданы');
          const fieldsHtml = START_LOCATIONS_FIELDS.map((cfg) => makeConfigField(
            cfg.label,
            cfg.key,
            row[cfg.key] ?? '',
            cfg
          )).join('');
          const card = document.createElement('div');
          card.className = 'config-card';
          card.innerHTML = `
            <div class="config-card__header">
              <div>
                <h3 class="config-card__title">${escapeHtml(titleRaw)}</h3>
                <p class="config-card__subtitle muted">${escapeHtml(subtitle)}</p>
              </div>
              <button class="mini-btn" data-act="del">Удалить</button>
            </div>
            <div class="config-grid">${fieldsHtml}</div>
          `;
          const titleEl = card.querySelector('.config-card__title');
          const subtitleEl = card.querySelector('.config-card__subtitle');
          const syncHeader = () => {
            const item = entry.record;
            const currentTitle = (item.ref && String(item.ref).trim()) || (item.id && String(item.id).trim()) || 'Новая точка';
            const parts = [];
            if (sourceNote) parts.push(sourceNote);
            const currentLat = item['point.lat'];
            const currentLon = item['point.lon'];
            if (currentLat !== '' && currentLon !== '') parts.push(`Координаты: ${currentLat}, ${currentLon}`);
            if (item.time_window) parts.push(`Окно: ${item.time_window}`);
            if (item.type) parts.push(`Тип: ${item.type}`);
            titleEl.textContent = currentTitle;
            subtitleEl.textContent = parts.length ? parts.join(' • ') : (item.id ? `ID: ${item.id}` : 'Данные не заданы');
          };
          card.querySelectorAll('[data-field]').forEach((input) => {
            const eventName = input.type === 'checkbox' ? 'change' : 'input';
            input.addEventListener(eventName, (ev) => {
              const el = ev.target;
              const key = el.getAttribute('data-field');
              if (!key) return;
              let value;
              if (el.type === 'number') value = el.value === '' ? '' : Number(el.value);
              else value = el.value;
              setMergedStartField(entry, key, value);
              syncHeader();
              render();
            });
          });
          const delBtn = card.querySelector('[data-act="del"]');
          if (delBtn) {
            delBtn.addEventListener('click', () => {
              if (confirm('Удалить эту стартовую точку из всех источников?')) {
                deleteMergedStart(entry);
              }
            });
          }
          syncHeader();
          list.appendChild(card);
        });
        dom.startTableWrap.innerHTML = '';
        dom.startTableWrap.appendChild(list);
        return;
      }
      if (!shouldLoadPersistedForMode(getActiveMode())) {
        dom.startCount.textContent = '0 точек';
        dom.startTableWrap.innerHTML = '<p class="empty-hint">Сначала нажмите «Обновить из Google Sheets», чтобы загрузить стартовые точки.</p>';
        return;
      }
      const rows = dataStore.startLocations || [];
      dom.startCount.textContent = `${rows.length} точек`;
      if (!rows.length) {
        dom.startTableWrap.innerHTML = '<p class="empty-hint">Нет стартовых точек. Нажмите «Добавить точку», чтобы создать первую.</p>';
        return;
      }
      const list = document.createElement('div');
      list.className = 'config-list';
      rows.forEach((row, index) => {
        const card = document.createElement('div');
        card.className = 'config-card';
        const titleRaw = (row.ref && String(row.ref).trim()) || (row.id && String(row.id).trim()) || `Точка #${index + 1}`;
        const latVal = row['point.lat'];
        const lonVal = row['point.lon'];
        const subtitleParts = [];
        if (latVal !== '' && lonVal !== '') subtitleParts.push(`Координаты: ${latVal}, ${lonVal}`);
        if (row.time_window) subtitleParts.push(`Окно: ${row.time_window}`);
        if (row.type) subtitleParts.push(`Тип: ${row.type}`);
        const subtitle = subtitleParts.length ? subtitleParts.join(' • ') : (row.id ? `ID: ${row.id}` : 'Данные не заданы');
        const fieldsHtml = START_LOCATIONS_FIELDS.map((cfg) => makeConfigField(
          cfg.label,
          cfg.key,
          row[cfg.key] ?? '',
          cfg
        )).join('');
        card.innerHTML = `
          <div class="config-card__header">
            <div>
              <h3 class="config-card__title">${escapeHtml(titleRaw)}</h3>
              <p class="config-card__subtitle muted">${escapeHtml(subtitle)}</p>
            </div>
            <button class="mini-btn" data-act="del">Удалить</button>
          </div>
          <div class="config-grid">
            ${fieldsHtml}
          </div>
        `;
        const titleEl = card.querySelector('.config-card__title');
        const subtitleEl = card.querySelector('.config-card__subtitle');
        const syncHeader = () => {
          const item = rows[index];
          const currentTitle = (item.ref && String(item.ref).trim()) || (item.id && String(item.id).trim()) || `Точка #${index + 1}`;
          const currentLat = item['point.lat'];
          const currentLon = item['point.lon'];
          const parts = [];
          if (currentLat !== '' && currentLon !== '') parts.push(`Координаты: ${currentLat}, ${currentLon}`);
          if (item.time_window) parts.push(`Окно: ${item.time_window}`);
          if (item.type) parts.push(`Тип: ${item.type}`);
          const currentSubtitle = parts.length ? parts.join(' • ') : (item.id ? `ID: ${item.id}` : 'Данные не заданы');
          titleEl.textContent = currentTitle;
          subtitleEl.textContent = currentSubtitle;
        };
        card.querySelectorAll('[data-field]').forEach((input) => {
          const eventName = input.type === 'checkbox' ? 'change' : 'input';
          input.addEventListener(eventName, (ev) => {
            const el = ev.target;
            const key = el.getAttribute('data-field');
            if (!key) return;
            if (el.type === 'number') {
              rows[index][key] = el.value === '' ? '' : Number(el.value);
            } else {
              rows[index][key] = el.value;
            }
            saveLocal();
            syncHeader();
            render();
          });
        });
        const delBtn = card.querySelector('[data-act="del"]');
        if (delBtn) {
          delBtn.addEventListener('click', () => {
            rows.splice(index, 1);
            saveLocal();
            renderStartLocations();
            render();
          });
        }
        syncHeader();
        list.appendChild(card);
      });
      dom.startTableWrap.innerHTML = '';
      dom.startTableWrap.appendChild(list);
      const validStart = new Set((dataStore.startLocations || []).map((_, index) => startLocationUid(_, index)));
      let removed = false;
      state.selected.forEach((uid) => {
        if (uid.startsWith('start::') && !validStart.has(uid)) {
          state.selected.delete(uid);
          removed = true;
        }
      });
      if (removed) updateSelectedCount();
    }

    // ===== Конфигурация таблиц =====
    function renderVehicles() {
      if (!dom.vehTableWrap || !dom.vehCount) return;
      if (isAllMode()) {
        if (!shouldLoadPersistedForMode(ALL_MODE)) {
          dom.vehCount.textContent = '0 записей';
          dom.vehTableWrap.innerHTML = '<p class="empty-hint">Сначала обновите данные из Google Sheets (загрузятся HoReCa и Галереи).</p>';
          return;
        }
        updateVehicleCountBadge();
        const merged = getMergedUniqueVehicles();
        if (!merged.length) {
          dom.vehTableWrap.innerHTML = '<p class="empty-hint">Записей пока нет. Обновите данные из Google Sheets.</p>';
          return;
        }
        const list = document.createElement('div');
        list.className = 'config-list';
        const intro = document.createElement('p');
        intro.className = 'muted empty-hint';
        intro.style.marginBottom = '12px';
        intro.textContent = 'Уникальные водители. Правки сохраняются локально и попадут в экспорт (не в Google Sheets).';
        list.appendChild(intro);
        merged.forEach((entry, index) => {
          const row = entry.row;
          const isExportSelected = isVehicleExportSelectedInAllMode(entry);
          const card = document.createElement('div');
          card.className = 'config-card' + (isExportSelected ? ' config-card--export-on' : '');
          const titleRaw = (row.ref && String(row.ref).trim()) || (row.id && String(row.id).trim()) || `Запись #${index + 1}`;
          const sourceNote = formatSourceLabels(entry.sourceLabels);
          const subtitleParts = [];
          if (sourceNote) subtitleParts.push(sourceNote);
          const start = row.start_at && String(row.start_at).trim();
          if (start) subtitleParts.push(`Начало: ${start}`);
          const finish = row.finish_at && String(row.finish_at).trim();
          if (finish) subtitleParts.push(`Финиш: ${finish}`);
          const shiftWindow = row['shifts.0.time_window'] && String(row['shifts.0.time_window']).trim();
          if (shiftWindow) subtitleParts.push(`Смена: ${shiftWindow}`);
          const capacity = row['capacity.weight_kg'];
          if (capacity != null && capacity !== '') subtitleParts.push(`Грузоподъемность: ${capacity} кг`);
          const depots = row.depot_id && String(row.depot_id).trim();
          if (!subtitleParts.length && depots) subtitleParts.push(`Склады: ${depots}`);
          const subtitleRaw = subtitleParts.length ? subtitleParts.join(' • ') : (row.id ? `ID: ${String(row.id).trim()}` : 'Данные не заданы');
          const fieldsHtml = VEHICLE_FIELD_CONFIG.map((cfg) => makeConfigField(
            cfg.label,
            cfg.key,
            row[cfg.key] ?? '',
            cfg
          )).join('');
          const flagsHtml = VEHICLE_FLAG_CONFIG.map((cfg) => makeConfigField(
            cfg.label,
            cfg.key,
            row[cfg.key],
            { type: 'checkbox' }
          )).join('');
          card.innerHTML = `
            <div class="config-card__header">
              <label class="config-export-toggle">
                <input type="checkbox" data-act="export" ${isExportSelected ? 'checked' : ''} />
                <span>В экспорт</span>
              </label>
              <div class="config-card__titles">
                <h3 class="config-card__title">${escapeHtml(titleRaw)}</h3>
                <p class="config-card__subtitle muted">${escapeHtml(subtitleRaw)}</p>
              </div>
              <button type="button" class="mini-btn" data-act="del">Удалить</button>
            </div>
            <div class="config-grid">${fieldsHtml}</div>
            <div class="config-flags">${flagsHtml}</div>
          `;
          const exportToggle = card.querySelector('[data-act="export"]');
          if (exportToggle) {
            exportToggle.addEventListener('change', (ev) => {
              const checked = ev.target.checked;
              setVehicleExportSelectedInAllMode(entry, checked);
              card.classList.toggle('config-card--export-on', checked);
              updateVehicleCountBadge();
            });
          }
          const titleEl = card.querySelector('.config-card__title');
          const subtitleEl = card.querySelector('.config-card__subtitle');
          const syncHeader = () => {
            const item = entry.row;
            const chunks = [];
            if (sourceNote) chunks.push(sourceNote);
            const startVal = item.start_at && String(item.start_at).trim();
            if (startVal) chunks.push(`Начало: ${startVal}`);
            const finishVal = item.finish_at && String(item.finish_at).trim();
            if (finishVal) chunks.push(`Финиш: ${finishVal}`);
            const shiftVal = item['shifts.0.time_window'] && String(item['shifts.0.time_window']).trim();
            if (shiftVal) chunks.push(`Смена: ${shiftVal}`);
            const cap = item['capacity.weight_kg'];
            if (cap != null && cap !== '') chunks.push(`Грузоподъемность: ${cap} кг`);
            const dep = item.depot_id && String(item.depot_id).trim();
            if (!chunks.length && dep) chunks.push(`Склады: ${dep}`);
            titleEl.textContent = (item.ref && String(item.ref).trim()) || (item.id && String(item.id).trim()) || `Запись #${index + 1}`;
            subtitleEl.textContent = chunks.length ? chunks.join(' • ') : (item.id ? `ID: ${String(item.id).trim()}` : 'Данные не заданы');
          };
          card.querySelectorAll('[data-field]').forEach((input) => {
            const eventName = input.type === 'checkbox' ? 'change' : 'input';
            input.addEventListener(eventName, (ev) => {
              const el = ev.target;
              const key = el.getAttribute('data-field');
              if (!key) return;
              let value;
              if (el.type === 'checkbox') value = el.checked;
              else if (el.type === 'number') value = el.value === '' ? '' : Number(el.value);
              else value = el.value;
              setMergedVehicleField(entry, key, value, el.type);
              syncHeader();
            });
          });
          const delBtn = card.querySelector('[data-act="del"]');
          if (delBtn) {
            delBtn.addEventListener('click', () => {
              if (confirm('Удалить этого водителя из всех источников?')) {
                deleteMergedVehicle(entry);
              }
            });
          }
          syncHeader();
          list.appendChild(card);
        });
        dom.vehTableWrap.innerHTML = '';
        dom.vehTableWrap.appendChild(list);
        return;
      }
      if (!shouldLoadPersistedForMode(getActiveMode())) {
        dom.vehCount.textContent = '0 записей';
        dom.vehTableWrap.innerHTML = '<p class="empty-hint">Сначала нажмите «Обновить из Google Sheets», чтобы загрузить водителей.</p>';
        return;
      }
      const rows = dataStore.vehicles || [];
      pruneVehicleSelection();
      updateVehicleCountBadge();
      if (!rows.length) {
        dom.vehTableWrap.innerHTML = '<p class="empty-hint">Записей пока нет. Нажмите «Добавить запись» или обновите данные из Google Sheets.</p>';
        return;
      }
      const selectedIds = getSelectedVehicleIds();
      const list = document.createElement('div');
      list.className = 'config-list';
      rows.forEach((row, index) => {
        const uid = vehicleUid(row, index);
        const isExportSelected = selectedIds.has(uid);
        const card = document.createElement('div');
        card.className = 'config-card' + (isExportSelected ? ' config-card--export-on' : '');
        const titleRaw = (row.ref && String(row.ref).trim()) || (row.id && String(row.id).trim()) || `Запись #${index + 1}`;
        const subtitleParts = [];
        const start = row.start_at && String(row.start_at).trim();
        if (start) subtitleParts.push(`Начало: ${start}`);
        const finish = row.finish_at && String(row.finish_at).trim();
        if (finish) subtitleParts.push(`Финиш: ${finish}`);
        const shiftWindow = row['shifts.0.time_window'] && String(row['shifts.0.time_window']).trim();
        if (shiftWindow) subtitleParts.push(`Смена: ${shiftWindow}`);
        const capacity = row['capacity.weight_kg'];
        if (capacity != null && capacity !== '') subtitleParts.push(`Грузоподъемность: ${capacity} кг`);
        const depots = row.depot_id && String(row.depot_id).trim();
        if (!subtitleParts.length && depots) subtitleParts.push(`Склады: ${depots}`);
        const subtitleFallback = row.id ? `ID: ${String(row.id).trim()}` : 'Данные не заданы';
        const subtitleRaw = subtitleParts.length ? subtitleParts.join(' • ') : subtitleFallback;
        const fieldsHtml = VEHICLE_FIELD_CONFIG.map((cfg) => makeConfigField(
          cfg.label,
          cfg.key,
          row[cfg.key] ?? '',
          cfg
        )).join('');
        const flagsHtml = VEHICLE_FLAG_CONFIG.map((cfg) => makeConfigField(
          cfg.label,
          cfg.key,
          row[cfg.key],
          { type: 'checkbox' }
        )).join('');
        card.innerHTML = `
          <div class="config-card__header">
            <label class="config-export-toggle">
              <input type="checkbox" data-act="export" ${isExportSelected ? 'checked' : ''} />
              <span>В экспорт</span>
            </label>
            <div class="config-card__titles">
              <h3 class="config-card__title">${escapeHtml(titleRaw)}</h3>
              <p class="config-card__subtitle muted">${escapeHtml(subtitleRaw)}</p>
            </div>
            <button type="button" class="mini-btn" data-act="del">Удалить</button>
          </div>
          <div class="config-grid">
            ${fieldsHtml}
          </div>
          <div class="config-flags">
            ${flagsHtml}
          </div>
        `;
        const exportToggle = card.querySelector('[data-act="export"]');
        if (exportToggle) {
          exportToggle.addEventListener('change', (ev) => {
            const checked = ev.target.checked;
            if (checked) getSelectedVehicleIds().add(uid);
            else getSelectedVehicleIds().delete(uid);
            card.classList.toggle('config-card--export-on', checked);
            saveLocal();
            updateVehicleCountBadge();
          });
        }
        const titleEl = card.querySelector('.config-card__title');
        const subtitleEl = card.querySelector('.config-card__subtitle');
        const syncHeader = () => {
          const item = rows[index];
          const title = (item.ref && String(item.ref).trim()) || (item.id && String(item.id).trim()) || `Запись #${index + 1}`;
          const subtitleChunks = [];
          const start = item.start_at && String(item.start_at).trim();
          if (start) subtitleChunks.push(`Начало: ${start}`);
          const finish = item.finish_at && String(item.finish_at).trim();
          if (finish) subtitleChunks.push(`Финиш: ${finish}`);
          const shiftWindowValue = item['shifts.0.time_window'] && String(item['shifts.0.time_window']).trim();
          if (shiftWindowValue) subtitleChunks.push(`Смена: ${shiftWindowValue}`);
          const capacityValue = item['capacity.weight_kg'];
          if (capacityValue != null && capacityValue !== '') subtitleChunks.push(`Грузоподъемность: ${capacityValue} кг`);
          const depotsVal = item.depot_id && String(item.depot_id).trim();
          if (!subtitleChunks.length && depotsVal) subtitleChunks.push(`Склады: ${depotsVal}`);
          const subtitle = subtitleChunks.length ? subtitleChunks.join(' • ') : (item.id ? `ID: ${String(item.id).trim()}` : 'Данные не заданы');
          titleEl.textContent = title;
          subtitleEl.textContent = subtitle;
        };
        card.querySelectorAll('[data-field]').forEach((input) => {
          const eventName = input.type === 'checkbox' ? 'change' : 'input';
          input.addEventListener(eventName, (ev) => {
            const el = ev.target;
            const key = el.getAttribute('data-field');
            if (!key) return;
            if (el.type === 'checkbox') {
              rows[index][key] = el.checked;
            } else if (el.type === 'number') {
              rows[index][key] = el.value === '' ? '' : Number(el.value);
            } else {
              const prevUid = vehicleUid(rows[index], index);
              const wasSelected = getSelectedVehicleIds().has(prevUid);
              rows[index][key] = el.value;
              if (key === 'id' && wasSelected) {
                getSelectedVehicleIds().delete(prevUid);
                getSelectedVehicleIds().add(vehicleUid(rows[index], index));
                saveLocal();
              }
            }
            saveLocal();
            syncHeader();
          });
        });
        const delBtn = card.querySelector('[data-act="del"]');
        if (delBtn) {
          delBtn.addEventListener('click', () => {
            getSelectedVehicleIds().delete(uid);
            rows.splice(index, 1);
            saveLocal();
            renderVehicles();
          });
        }
        syncHeader();
        list.appendChild(card);
      });
      dom.vehTableWrap.innerHTML = '';
      dom.vehTableWrap.appendChild(list);
    }

    function renderDepots() {
      if (!dom.depTableWrap || !dom.depCount) return;
      if (!shouldLoadPersistedForMode(getActiveMode())) {
        dom.depCount.textContent = '0 складов';
        dom.depTableWrap.innerHTML = '<p class="empty-hint">Сначала нажмите «Обновить из Google Sheets», чтобы загрузить склады.</p>';
        return;
      }
      const depotEntries = isAllMode()
        ? getMergedUniqueDepotsWithMeta()
        : (dataStore.depots || []).map((depot) => ({ depot, sourceLabels: [] }));
      dom.depCount.textContent = `${depotEntries.length} складов`;
      if (!depotEntries.length) {
        dom.depTableWrap.innerHTML = '<p class="empty-hint">Складов пока нет. Нажмите «Добавить склад», чтобы создать первый.</p>';
        return;
      }
      const list = document.createElement('div');
      list.className = 'config-list';
      depotEntries.forEach((entry, index) => {
        const row = entry.depot;
        const card = document.createElement('div');
        card.className = 'config-card';
        const titleRaw = (row.ref && String(row.ref).trim()) || (row.id && String(row.id).trim()) || `Склад #${index + 1}`;
        const latVal = row['point.lat'];
        const lonVal = row['point.lon'];
        const hasLat = latVal !== undefined && latVal !== null && latVal !== '';
        const hasLon = lonVal !== undefined && lonVal !== null && lonVal !== '';
        const sourceNote = formatSourceLabels(entry.sourceLabels);
        let subtitleRaw = hasLat && hasLon
          ? `Координаты: ${latVal}, ${lonVal}`
          : row.time_window && String(row.time_window).trim()
            ? `Время работы: ${row.time_window}`
            : `ID: ${row.id ? String(row.id).trim() : 'не задан'}`;
        if (sourceNote) subtitleRaw = `${sourceNote} • ${subtitleRaw}`;
        const fieldsHtml = [
          makeConfigField('ID', 'id', row.id ?? '', { placeholder: 'например 1' }),
          makeConfigField('Название', 'ref', row.ref ?? '', { placeholder: 'например Основной склад' }),
          makeConfigField('Широта', 'point.lat', row['point.lat'] ?? '', { type: 'number', step: 'any' }),
          makeConfigField('Долгота', 'point.lon', row['point.lon'] ?? '', { type: 'number', step: 'any' }),
          makeConfigField('Время работы', 'time_window', row.time_window ?? '', { placeholder: 'например 09:00-21:00' })
        ].join('');
        card.innerHTML = `
          <div class="config-card__header">
            <div>
              <h3 class="config-card__title">${escapeHtml(titleRaw)}</h3>
              <p class="config-card__subtitle muted">${escapeHtml(subtitleRaw)}</p>
            </div>
            ${isAllMode() ? '' : '<button class="mini-btn" data-act="del">Удалить</button>'}
          </div>
          <div class="config-grid">
            ${fieldsHtml}
          </div>
        `;
        const titleEl = card.querySelector('.config-card__title');
        const subtitleEl = card.querySelector('.config-card__subtitle');
        const syncHeader = () => {
          const item = row;
          const title = (item.ref && String(item.ref).trim()) || (item.id && String(item.id).trim()) || `Склад #${index + 1}`;
          const lat = item['point.lat'];
          const lon = item['point.lon'];
          const hasLatVal = lat !== undefined && lat !== null && lat !== '';
          const hasLonVal = lon !== undefined && lon !== null && lon !== '';
          let subtitle = hasLatVal && hasLonVal
            ? `Координаты: ${lat}, ${lon}`
            : (item.time_window && String(item.time_window).trim())
              ? `Время работы: ${item.time_window}`
              : `ID: ${item.id ? String(item.id).trim() : 'не задан'}`;
          if (sourceNote) subtitle = `${sourceNote} • ${subtitle}`;
          titleEl.textContent = title;
          subtitleEl.textContent = subtitle;
        };
        card.querySelectorAll('[data-field]').forEach((input) => {
          if (isAllMode()) input.disabled = true;
          const eventName = input.type === 'checkbox' ? 'change' : 'input';
          input.addEventListener(eventName, (ev) => {
            if (isAllMode()) return;
            const el = ev.target;
            const key = el.getAttribute('data-field');
            if (!key) return;
            if (el.type === 'checkbox') row[key] = el.checked;
            else if (el.type === 'number') row[key] = el.value === '' ? '' : Number(el.value);
            else row[key] = el.value;
            saveLocal();
            syncHeader();
          });
        });
        const delBtn = card.querySelector('[data-act="del"]');
        if (delBtn) {
          delBtn.addEventListener('click', () => {
            dataStore.depots.splice(index, 1);
            saveLocal();
            renderDepots();
          });
        }
        syncHeader();
        list.appendChild(card);
      });
      dom.depTableWrap.innerHTML = '';
      dom.depTableWrap.appendChild(list);
    }

    // ===== Экспорт/импорт настроек =====
    function exportSettings() {
      const payload = {
        vehicles: dataStore.vehicles,
        depots: dataStore.depots,
        startLocations: dataStore.startLocations,
        selectedVehicleIds: Array.from(getSelectedVehicleIds())
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'settings.json';
      document.body.appendChild(link);
      link.click();
      URL.revokeObjectURL(link.href);
      link.remove();
    }

    async function importSettings(file) {
      try {
        const text = await file.text();
        const json = JSON.parse(text);
        if (Array.isArray(json.vehicles)) {
          dataStore.vehicles = json.vehicles.map((item) => normalizeVehicleRecord(item)).filter(Boolean);
          if (Array.isArray(json.selectedVehicleIds)) {
            getActiveStores().selectedVehicleIds = new Set(json.selectedVehicleIds.map((id) => String(id)));
          } else {
            rebuildVehicleSelectionAfterLoad();
          }
        }
        if (Array.isArray(json.depots)) dataStore.depots = json.depots.map((item) => normalizeDepotRecord(item)).filter(Boolean);
        const settingsStart = json.startLocations || json.start_locations || json.startData || json.StartData;
        dataStore.startLocations = Array.isArray(settingsStart)
          ? settingsStart.map((item) => normalizeStartRecord(item)).filter(Boolean)
          : [];
        pruneVehicleSelection();
        saveLocal();
        renderVehicles();
        renderDepots();
        renderStartLocations();
        render();
      } catch (err) {
        showError('Не удалось импортировать настройки: ' + err.message);
      }
    }

    function updateModeUI() {
      const mode = getActiveMode();
      const config = getModeConfig(mode);
      dom.modeButtons.forEach((btn) => {
        const active = btn.dataset.mode === mode;
        btn.classList.toggle('active', active);
        btn.setAttribute('aria-pressed', active ? 'true' : 'false');
      });
      if (dom.dataSourceLabel) {
        if (isAllMode(mode)) {
          const labels = SOURCE_MODE_IDS.map((id) => getModeConfig(id).label).join(' + ');
          dom.dataSourceLabel.textContent = `Источники: ${labels}`;
        } else {
          dom.dataSourceLabel.textContent = `Источник: ${config.label}`;
        }
      }
      updateSheetCacheUi();
    }

    function setActiveBusinessMode(mode) {
      if (!isValidBusinessMode(mode)) return;
      if (mode === state.activeMode) {
        updateModeUI();
        return;
      }
      state.activeMode = mode;
      try {
        localStorage.setItem('activeMode', mode);
      } catch (_) {}
      if (isAllMode(mode)) {
        SOURCE_MODE_IDS.forEach((sourceId) => {
          ensureEmptyModeUntilSheetSync(sourceId);
          loadLocalForMode(sourceId);
        });
        loadLocalForMode(ALL_MODE);
      } else {
        ensureEmptyModeUntilSheetSync(mode);
        loadLocalForMode(mode);
      }
      updateModeUI();
      renderVehicles();
      renderDepots();
      renderStartLocations();
      applyAutoStartSelection({ dayKey: state.activeDay, selectDefaultExtra: isDefaultExtraSelected() });
      render();
      updateSelectedCount();
      requestAnimationFrame(() => {
        if (!shouldShowSheetOnboarding()) {
          if (dom.sheetOnboarding && !dom.sheetOnboarding.hidden) dismissSheetOnboarding(true);
        } else if (!dom.sheetOnboarding || dom.sheetOnboarding.hidden) {
          showSheetOnboarding();
        } else {
          positionSheetOnboarding();
        }
      });
    }

    // ===== Слушатели =====
    dom.modeButtons.forEach((btn) => {
      btn.addEventListener('click', () => {
        const mode = btn.dataset.mode;
        if (mode) setActiveBusinessMode(mode);
      });
    });
    dom.tabs.forEach((btn) => btn.addEventListener('click', () => setActiveDay(btn.dataset.day)));
    if (dom.loadJsonBtn) {
      dom.loadJsonBtn.addEventListener('click', () => {
        if (dom.jsonFile && dom.jsonFile.files && dom.jsonFile.files[0]) {
          loadLocalJson(dom.jsonFile.files[0]);
        } else {
          showError('Выберите файл JSON.');
        }
      });
    }

    if (dom.errorClose) {
      dom.errorClose.addEventListener('click', () => clearError());
    }

    if (dom.notifyClose) {
      dom.notifyClose.addEventListener('click', () => hideNotify());
    }

    ensureActionButton(dom.loadSheetBtn);

    if (dom.loadSheetBtn) {
      dom.loadSheetBtn.addEventListener('click', () => loadFromGoogleSheets());
    }

    if (dom.settingsExportBtn) dom.settingsExportBtn.addEventListener('click', exportSettings);
    if (dom.settingsImportBtn && dom.settingsImportInput) {
      dom.settingsImportBtn.addEventListener('click', () => dom.settingsImportInput.click());
      dom.settingsImportInput.addEventListener('change', () => {
        if (dom.settingsImportInput.files && dom.settingsImportInput.files[0]) {
          importSettings(dom.settingsImportInput.files[0]);
        }
      });
    }

    if (dom.vehSelectAllBtn) {
      dom.vehSelectAllBtn.addEventListener('click', () => {
        setAllVehiclesSelected(true);
        renderVehicles();
      });
    }
    if (dom.vehClearSelectionBtn) {
      dom.vehClearSelectionBtn.addEventListener('click', () => {
        setAllVehiclesSelected(false);
        renderVehicles();
      });
    }

    if (dom.vehAddBtn) dom.vehAddBtn.addEventListener('click', () => {
      if (isAllMode()) {
        addVehicleToAllSources();
        return;
      }
      dataStore.vehicles.push({
        id: '',
        ref: '',
        'capacity.weight_kg': '',
        tags: '',
        start_at: '',
        finish_at: '',
        visit_depot_at_start: false,
        return_to_depot: false,
        depot_id: '',
        'shifts.0.id': '',
        'shifts.0.time_window': '',
        allow_different_depots_in_route: false,
        max_middle_depots: '',
        depots_only_at_run_beginning: false,
        starting_depot_id: '',
        middle_depot_id: ''
      });
      const idx = dataStore.vehicles.length - 1;
      getSelectedVehicleIds().add(vehicleUid(dataStore.vehicles[idx], idx));
      saveLocal();
      renderVehicles();
    });

    if (dom.vehClearBtn) dom.vehClearBtn.addEventListener('click', () => {
      if (confirm('Очистить все записи Vehicles?')) {
        if (isAllMode()) {
          SOURCE_MODE_IDS.forEach((sourceMode) => {
            const store = getStoreForMode(sourceMode);
            store.vehicles = [];
            store.selectedVehicleIds = new Set();
            saveLocalForMode(sourceMode);
          });
          renderVehicles();
          return;
        }
        dataStore.vehicles = [];
        getActiveStores().selectedVehicleIds = new Set();
        saveLocal();
        renderVehicles();
      }
    });

    if (dom.depAddBtn) dom.depAddBtn.addEventListener('click', () => {
      dataStore.depots.push({ id: '', ref: '', 'point.lat': '', 'point.lon': '', time_window: '' });
      saveLocal();
      renderDepots();
    });

    if (dom.depClearBtn) dom.depClearBtn.addEventListener('click', () => {
      if (confirm('Очистить все склады?')) {
        dataStore.depots = [];
        saveLocal();
        renderDepots();
      }
    });

    if (dom.startAddBtn) dom.startAddBtn.addEventListener('click', () => {
      if (isAllMode()) {
        addStartToAllSources();
        return;
      }
      (dataStore.startLocations || (dataStore.startLocations = [])).push({
        id: '',
        ref: '',
        'point.lat': '',
        'point.lon': '',
        time_window: '07:00:00-22:00:00',
        type: 'garage',
        address: '',
        comments: '',
        phone: ''
      });
      saveLocal();
      renderStartLocations();
      render();
    });

    if (dom.startClearBtn) dom.startClearBtn.addEventListener('click', () => {
      if (confirm('Очистить все стартовые точки?')) {
        if (isAllMode()) {
          SOURCE_MODE_IDS.forEach((sourceMode) => {
            const store = getStoreForMode(sourceMode);
            store.startLocations = [];
            saveLocalForMode(sourceMode);
          });
          renderStartLocations();
          render();
          return;
        }
        dataStore.startLocations = [];
        saveLocal();
        renderStartLocations();
        render();
      }
    });

    // ===== Инициализация =====
    (function init() {
      const yandexLink = document.getElementById('yandexPlanningLink');
      if (yandexLink && APP.yandexPlanningUrl) {
        yandexLink.href = APP.yandexPlanningUrl;
      }
      migrateLegacyLocalStorage();
      migrateHorecaStaleCache();
      expireStaleSheetCaches();
      try {
        const savedMode = localStorage.getItem('activeMode');
        if (savedMode && MODE_CONFIG[savedMode]) {
          state.activeMode = savedMode;
        } else {
          state.activeMode = DEFAULT_MODE;
        }
      } catch (_) {
        state.activeMode = DEFAULT_MODE;
      }
      updateModeUI();
      ensureEmptyModeUntilSheetSync('horeca');
      applyConfigSeedForAllModes();
      ensureEmptyModeUntilSheetSync(state.activeMode);
      if (isAllMode(state.activeMode)) {
        SOURCE_MODE_IDS.forEach((id) => loadLocalForMode(id));
        loadLocalForMode(ALL_MODE);
      } else {
        loadLocal();
      }
      applyConfigSeedIfEmpty();
      ensureVehicleSelectionDefault();
      renderVehicles();
      renderDepots();
      renderStartLocations();
      renderSchemaDocs();
      const jsDay = new Date().getDay();
      const mapJsToCanon = { 1: 'monday', 2: 'tuesday', 3: 'wednesday', 4: 'thursday', 5: 'friday', 6: 'saturday', 0: 'sunday' };
      setActiveDay(mapJsToCanon[jsDay] || 'monday');
      runSelfTestsIfEnabled();
      ensureEmptyModeUntilSheetSync(state.activeMode);
      refreshModeUiAfterDataChange();
      initSheetOnboarding();
      initExtraOrderModal();
      startSheetCacheUiTimer();
      updateSheetCacheUi();
      window.app = {
        state,
        dataStore,
        getScheduleData,
        storesByMode,
        APP_CONFIG: APP,
        MODE_CONFIG,
        MODE_IDS,
        SOURCE_MODE_IDS,
        ALL_MODE,
        getModeConfig,
        setActiveBusinessMode,
        showSheetOnboarding,
        dismissSheetOnboarding,
        EXPORT_SCHEMAS
      };
    })();

    function runSelfTestsIfEnabled() {
      let enabled = false;
      try {
        enabled = new URLSearchParams(window.location.search).get('test') === '1';
      } catch (_) {}
      if (!enabled) return;

      const results = [];
      const assert = (name, condition) => {
        results.push({ name, pass: !!condition });
        if (!condition) console.error('[TEST FAIL]', name);
      };
      const snapshotModes = {};
      MODE_IDS.forEach((modeId) => {
        snapshotModes[modeId] = JSON.parse(JSON.stringify(storesByMode[modeId]));
      });
      const savedMode = state.activeMode;
      try {
        state.activeMode = 'gallery';
        const galleryTw = getDefaultTimeWindow();
        assert('default time_window gallery', ((row) => row.time_window || getDefaultTimeWindow())({}) === galleryTw);
        state.activeMode = 'horeca';
        assert('horeca export prefix', getModeConfig().exportFilePrefix === 'HoReCa_заказы');
        state.activeMode = savedMode;
        const parsedHoreca = parseSheetDataset({ понедельник: [] }, SHEET_NAME_BY_DAY);
        assert('parseSheetDataset custom sheet names', parsedHoreca.schedule.monday.length === 0);
        setActiveBusinessMode('gallery');
        const gallerySchedule = getScheduleData();
        setActiveBusinessMode('horeca');
        assert('mode switch preserves per-mode cache', gallerySchedule !== getScheduleData() || state.activeMode === 'horeca');
        state.activeMode = savedMode;
        try { localStorage.setItem('activeMode', savedMode); } catch (_) {}
        updateModeUI();
        loadLocal();

        dataStore.depots = [{ id: '1', ref: 'Склад 1', 'point.lat': 55.75, 'point.lon': 37.61, time_window: '09:00-21:00' }];
        const depAoA = buildDepotAoA();
        assert('Depot headers length', depAoA[0].length === 5);
        assert('Depot body length', depAoA.length >= 2 && depAoA[1].length === 5);

        dataStore.vehicles = [{
          id: 'v1',
          ref: 'Машина 1',
          'capacity.weight_kg': 1200,
          tags: 'изотерма',
          start_at: 'depot:1',
          finish_at: 'depot:2',
          visit_depot_at_start: true,
          return_to_depot: true,
          depot_id: '1',
          'shifts.0.id': 'shift:1',
          'shifts.0.time_window': '09:00-18:00',
          allow_different_depots_in_route: true,
          max_middle_depots: 2,
          depots_only_at_run_beginning: false,
          starting_depot_id: '1',
          middle_depot_id: '2'
        }];
        dataStore.vehicles = [
          { id: 'v1', ref: 'Водитель 1', start_at: '', finish_at: '', visit_depot_at_start: false, return_to_depot: false, depot_id: '', 'shifts.0.id': '', 'shifts.0.time_window': '', allow_different_depots_in_route: false, max_middle_depots: '', depots_only_at_run_beginning: false, starting_depot_id: '', middle_depot_id: '' },
          { id: 'v2', ref: 'Водитель 2', start_at: '', finish_at: '', visit_depot_at_start: false, return_to_depot: false, depot_id: '', 'shifts.0.id': '', 'shifts.0.time_window': '', allow_different_depots_in_route: false, max_middle_depots: '', depots_only_at_run_beginning: false, starting_depot_id: '', middle_depot_id: '' }
        ];
        getActiveStores().selectedVehicleIds = new Set(['veh:id::v1']);
        const vehAoA = buildVehiclesAoA();
        const vehCols = EXPORT_SCHEMAS.vehicles.columns.length;
        assert('Vehicles headers length', vehAoA[0].length === vehCols);
        assert('Vehicles export only selected', vehAoA.length === 3 && vehAoA[2][0] !== 'v2');

        const before = { vehicles: [{ id: 'v1' }], depots: [{ id: 'd1' }], startLocations: [{ id: 's1', ref: 'Тест', 'point.lat': 55.1, 'point.lon': 37.1, time_window: '07:00:00-22:00:00', type: 'garage' }] };
        state.activeMode = 'gallery';
        dataStore.vehicles = before.vehicles;
        dataStore.depots = before.depots;
        dataStore.startLocations = before.startLocations;
        saveLocal();
        dataStore.vehicles = [];
        dataStore.depots = [];
        dataStore.startLocations = [];
        loadLocal();
        assert('loadLocal restores vehicles length', dataStore.vehicles.length >= 1);
        assert('loadLocal restores depots length', dataStore.depots.length >= 1);
        assert('loadLocal restores startLocations length', dataStore.startLocations.length >= 1);
        state.activeMode = savedMode;

        const special = ".+?^${}()|[]\\";
        let ok = true;
        try { highlight('a.+?b [x] (y) \\', special); } catch (err) { ok = false; }
        assert('highlight handles special chars safely', ok === true);
        const res = highlight('Тест .+? скобки ( ) и [квадратные] и \\ бэкслэш', '[квадратные]');
        assert('highlight actually marks text', /<mark>\[квадратные\]<\/mark>/.test(res));
      } finally {
        MODE_IDS.forEach((modeId) => {
          storesByMode[modeId] = JSON.parse(JSON.stringify(snapshotModes[modeId]));
        });
        state.activeMode = savedMode;
        try { localStorage.setItem('activeMode', savedMode); } catch (_) {}
        MODE_IDS.forEach((modeId) => ensureEmptyModeUntilSheetSync(modeId));
      }
    }
